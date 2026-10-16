"""Attendance workbook classification app.

Streamlit UI with Polars-based attendance classification and surgical XLSX edits.
Only the explicitly selected attendance result cells are changed in the output.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import re
import zipfile
from typing import Any, Iterable
import xml.etree.ElementTree as ET

import polars as pl


RULES: dict[int, dict[str, float]] = {
    # Thresholds are in hours. Inclusive lower bounds avoid gaps between codes.
    9: {"absence_below": 5.0, "half_from": 5.0, "short_from": 8.0, "present_from": 8.75, "present_plus_through": 17.0},
    10: {"absence_below": 5.0, "half_from": 5.0, "short_from": 9.0, "present_from": 9.75, "present_plus_through": 18.0},
    11: {"absence_below": 5.0, "half_from": 5.0, "short_from": 10.0, "present_from": 10.75, "present_plus_through": 20.0},
    12: {"absence_below": 6.0, "half_from": 6.0, "short_from": 10.0, "present_from": 11.75, "present_plus_through": 22.0},
}

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


class WorkbookSafetyError(ValueError):
    """Raised when the workbook cannot be changed with the promised scope."""


@dataclass
class ClassificationResult:
    row_numbers: list[int]
    proposed: list[str | None]
    messages: list[str]


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def parse_hours(value: Any, *, excel_day_fraction: bool = False) -> float | None:
    """Parse a duration as hours; strings like '10:45' mean 10 hours 45 minutes."""
    if _is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if hasattr(value, "total_seconds"):
        hours = value.total_seconds() / 3600
    elif isinstance(value, (int, float)):
        hours = float(value) * (24 if excel_day_fraction else 1)
    else:
        raw = str(value).strip().lower().replace("hours", "").replace("hour", "").replace("hrs", "").replace("hr", "").strip()
        match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*:\s*(\d{1,2})(?::(\d{1,2}))?", raw)
        if match:
            h, m = float(match.group(1)), int(match.group(2))
            seconds = int(match.group(3) or 0)
            if m >= 60 or seconds >= 60:
                return None
            hours = h + m / 60 + seconds / 3600
        else:
            try:
                hours = float(raw)
            except (TypeError, ValueError):
                return None
    if hours < 0:
        return None
    return hours


def parse_shift(value: Any) -> int | None:
    hours = parse_hours(value)
    if hours is None or not hours.is_integer():
        return None
    shift = int(hours)
    return shift if shift in RULES else None


def classify_attendance(shift_value: Any, worked_value: Any) -> tuple[str | None, str | None]:
    """Return (code, review message), without making assumptions for invalid rows."""
    shift = parse_shift(shift_value)
    hours = parse_hours(worked_value)
    if shift is None:
        return None, "Missing or unsupported shift length (supported: 9, 10, 11, 12 hours)."
    if hours is None:
        return None, "Missing or invalid worked-hours value."
    rule = RULES[shift]
    if hours < rule["absence_below"]:
        return "A", None
    if hours < rule["short_from"]:
        return "HF", None
    if hours < rule["present_from"]:
        return "SP", None
    if hours <= rule["present_plus_through"]:
        return "P", None
    return "PP", None


def classify_records(
    rows: Iterable[dict[str, Any]],
    *,
    shift_column: str,
    hours_column: str,
    in_column: str,
    out_column: str,
    excel_day_fraction: bool = False,
) -> ClassificationResult:
    """Classify records using Polars expressions/data columns for calculations."""
    materialized = list(rows)
    row_numbers = [int(r["_excel_row"]) for r in materialized]
    shifts = [r.get(shift_column) for r in materialized]
    worked = [parse_hours(r.get(hours_column), excel_day_fraction=excel_day_fraction) for r in materialized]
    in_values = [r.get(in_column) for r in materialized]
    out_values = [r.get(out_column) for r in materialized]

    frame = pl.DataFrame({"shift": shifts, "worked_hours": worked, "in_punch": in_values, "out_punch": out_values}, strict=False)
    codes: list[str | None] = []
    messages: list[str] = []
    for index, record in enumerate(materialized):
        has_in, has_out = not _is_blank(in_values[index]), not _is_blank(out_values[index])
        if has_in != has_out:
            codes.append("MP")
            messages.append("Single punch: " + ("in-punch missing." if not has_in else "out-punch missing."))
            continue
        if not has_in and not has_out:
            codes.append(None)
            messages.append("No punches are present; row left unchanged for review.")
            continue
        code, message = classify_attendance(frame["shift"][index], frame["worked_hours"][index])
        codes.append(code)
        messages.append(message or "Ready")
    return ClassificationResult(row_numbers, codes, messages)


def _xml_tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def _column_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        raise WorkbookSafetyError(f"Invalid Excel cell reference: {cell_ref}")
    n = 0
    for char in letters.group(0):
        n = n * 26 + ord(char) - 64
    return n


def _column_letters(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def _workbook_sheet_paths(book: zipfile.ZipFile) -> dict[str, str]:
    try:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError) as exc:
        raise WorkbookSafetyError("This file is not a supported Excel .xlsx workbook.") from exc
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")}
    result: dict[str, str] = {}
    for sheet in workbook.findall(f".//{_xml_tag('sheet')}"):
        rid = sheet.attrib.get(f"{{{REL_NS}}}id")
        if rid not in targets:
            continue
        target = targets[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = str(Path("xl") / target)
        result[sheet.attrib["name"]] = target.replace("\\", "/")
    return result


def _shared_strings(book: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(t.text or "" for t in item.findall(f".//{_xml_tag('t')}")) for item in root.findall(_xml_tag("si"))]


def _cell_value(cell: ET.Element, strings: list[str]) -> Any:
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(f".//{_xml_tag('t')}"))
    value = cell.find(_xml_tag("v"))
    if value is None:
        return None
    raw = value.text or ""
    if kind == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return None
    if kind == "b":
        return raw == "1"
    if kind in ("str", "e"):
        return raw
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def read_sheet_rows(file_bytes: bytes, sheet_name: str, header_row: int = 1) -> tuple[list[str], list[dict[str, Any]], int]:
    """Read cell values while retaining actual worksheet row numbers."""
    with zipfile.ZipFile(BytesIO(file_bytes)) as book:
        paths = _workbook_sheet_paths(book)
        if sheet_name not in paths:
            raise WorkbookSafetyError(f"Worksheet not found: {sheet_name}")
        try:
            root = ET.fromstring(book.read(paths[sheet_name]))
        except (KeyError, ET.ParseError) as exc:
            raise WorkbookSafetyError("Unable to read the selected worksheet.") from exc
        strings = _shared_strings(book)

    row_elements = root.findall(f".//{_xml_tag('sheetData')}/{_xml_tag('row')}")
    by_row: dict[int, dict[int, Any]] = {}
    max_col = 0
    for row in row_elements:
        row_idx = int(row.attrib["r"])
        values: dict[int, Any] = {}
        for cell in row.findall(_xml_tag("c")):
            col = _column_number(cell.attrib["r"])
            max_col = max(max_col, col)
            values[col] = _cell_value(cell, strings)
        by_row[row_idx] = values
    if header_row not in by_row:
        raise WorkbookSafetyError(f"Header row {header_row} is empty or missing.")
    headers: list[str] = []
    seen: set[str] = set()
    for col in range(1, max_col + 1):
        raw = by_row[header_row].get(col)
        name = str(raw) if not _is_blank(raw) else f"Column_{_column_letters(col)}"
        if name in seen:
            name = f"{name}__{_column_letters(col)}"
        seen.add(name)
        headers.append(name)
    records: list[dict[str, Any]] = []
    for row_idx in sorted(r for r in by_row if r > header_row):
        values = by_row[row_idx]
        record = {headers[col - 1]: values.get(col) for col in range(1, len(headers) + 1)}
        record["_excel_row"] = row_idx
        records.append(record)
    return headers, records, max_col


def _target_sheet_path(file_bytes: bytes, sheet_name: str) -> str:
    with zipfile.ZipFile(BytesIO(file_bytes)) as book:
        paths = _workbook_sheet_paths(book)
        if sheet_name not in paths:
            raise WorkbookSafetyError(f"Worksheet not found: {sheet_name}")
        return paths[sheet_name]


def _write_inline_string(cell: ET.Element, value: str) -> None:
    # Retain cell attributes, including style, while replacing only its contents.
    for child in list(cell):
        cell.remove(child)
    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, _xml_tag("is"))
    text = ET.SubElement(inline, _xml_tag("t"))
    if value != value.strip():
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value


def _set_cell(root: ET.Element, cell_ref: str, value: str) -> None:
    sheet_data = root.find(_xml_tag("sheetData"))
    if sheet_data is None:
        raise WorkbookSafetyError("Worksheet has no sheet data section.")
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    assert match is not None
    col, row_num = _column_number(cell_ref), int(match.group(2))
    row = next((r for r in sheet_data.findall(_xml_tag("row")) if int(r.attrib["r"]) == row_num), None)
    if row is None:
        row = ET.Element(_xml_tag("row"), {"r": str(row_num)})
        preceding = next((i for i, existing in enumerate(sheet_data.findall(_xml_tag("row"))) if int(existing.attrib["r"]) > row_num), None)
        sheet_data.insert(preceding if preceding is not None else len(sheet_data), row)
    cell = next((c for c in row.findall(_xml_tag("c")) if c.attrib.get("r") == cell_ref), None)
    if cell is None:
        cell = ET.Element(_xml_tag("c"), {"r": cell_ref})
        existing_cells = row.findall(_xml_tag("c"))
        insert_at = next((i for i, c in enumerate(existing_cells) if _column_number(c.attrib["r"]) > col), None)
        row.insert(insert_at if insert_at is not None else len(row), cell)
    if cell.find(_xml_tag("f")) is not None:
        raise WorkbookSafetyError(f"Target cell {cell_ref} contains a formula; refusing to overwrite it.")
    _write_inline_string(cell, value)


def _extend_dimension(root: ET.Element, column: int, row_number: int) -> None:
    dimension = root.find(_xml_tag("dimension"))
    if dimension is None:
        return
    current = dimension.attrib.get("ref", "A1")
    parts = current.split(":")
    end_ref = parts[-1]
    end_match = re.fullmatch(r"([A-Z]+)(\d+)", end_ref)
    if not end_match:
        return
    end_col, end_row = _column_number(end_ref), int(end_match.group(2))
    new_end = f"{_column_letters(max(end_col, column))}{max(end_row, row_number)}"
    dimension.attrib["ref"] = f"{parts[0]}:{new_end}" if len(parts) > 1 else f"A1:{new_end}"


def create_modified_workbook(
    file_bytes: bytes,
    *,
    sheet_name: str,
    classifications: dict[int, str],
    target_column: int | None,
    append_column: bool = False,
    header_row: int = 1,
    appended_header: str = "Attendance Classification",
) -> bytes:
    """Patch only the selected worksheet XML cells; all other ZIP parts stay byte-identical."""
    if not zipfile.is_zipfile(BytesIO(file_bytes)):
        raise WorkbookSafetyError("Upload a valid .xlsx workbook.")
    sheet_path = _target_sheet_path(file_bytes, sheet_name)
    headers, records, max_col = read_sheet_rows(file_bytes, sheet_name, header_row)
    if append_column:
        if target_column is not None:
            raise WorkbookSafetyError("Choose an existing target column or append one, not both.")
        target_column = max_col + 1
    if target_column is None or target_column < 1:
        raise WorkbookSafetyError("Select an existing classification column or explicitly choose to append one.")
    if not append_column and target_column > max_col:
        raise WorkbookSafetyError("Selected target column does not exist. Opt in to append a classification column.")
    if append_column:
        with zipfile.ZipFile(BytesIO(file_bytes)) as book:
            root = ET.fromstring(book.read(sheet_path))
        if root.find(f".//{_xml_tag('mergeCell')}") is not None:
            raise WorkbookSafetyError("Cannot safely append a column to a worksheet with merged cells.")
        if root.find(f".//{_xml_tag('tablePart')}") is not None:
            raise WorkbookSafetyError("Cannot safely append a column to a worksheet containing an Excel table.")

    with zipfile.ZipFile(BytesIO(file_bytes)) as book:
        original_members = {item.filename: book.read(item.filename) for item in book.infolist()}
    sheet_root = ET.fromstring(original_members[sheet_path])
    if append_column:
        _set_cell(sheet_root, f"{_column_letters(target_column)}{header_row}", appended_header)
        _extend_dimension(sheet_root, target_column, header_row)
    for row_number, code in classifications.items():
        if code not in {"PP", "P", "SP", "HF", "A", "MP"}:
            raise WorkbookSafetyError(f"Unsupported attendance code: {code}")
        _set_cell(sheet_root, f"{_column_letters(target_column)}{row_number}", code)
        if append_column:
            _extend_dimension(sheet_root, target_column, row_number)
    original_members[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
        for name, payload in original_members.items():
            out_zip.writestr(name, payload)
    result = output.getvalue()
    verify_workbook_patch(file_bytes, result, sheet_name, classifications, target_column, append_column, header_row, appended_header)
    return result


def _canonical_sheet_without_targets(xml_bytes: bytes, target_refs: set[str], *, ignore_dimension: bool = False) -> bytes:
    root = ET.fromstring(xml_bytes)
    for cell in list(root.findall(f".//{_xml_tag('sheetData')}/{_xml_tag('row')}/{_xml_tag('c')}")):
        if cell.attrib.get("r") in target_refs:
            parent_row = next(row for row in root.findall(f".//{_xml_tag('sheetData')}/{_xml_tag('row')}") if cell in list(row))
            parent_row.remove(cell)
    if ignore_dimension:
        dimension = root.find(_xml_tag("dimension"))
        if dimension is not None:
            root.remove(dimension)
    return ET.tostring(root, encoding="utf-8")


def verify_workbook_patch(
    original: bytes,
    output: bytes,
    sheet_name: str,
    classifications: dict[int, str],
    target_column: int,
    append_column: bool,
    header_row: int,
    appended_header: str,
) -> None:
    with zipfile.ZipFile(BytesIO(original)) as before, zipfile.ZipFile(BytesIO(output)) as after:
        before_names, after_names = set(before.namelist()), set(after.namelist())
        if before_names != after_names:
            raise WorkbookSafetyError("Verification failed: workbook package structure changed.")
        sheet_path = _workbook_sheet_paths(before)[sheet_name]
        for name in before_names - {sheet_path}:
            if before.read(name) != after.read(name):
                raise WorkbookSafetyError(f"Verification failed: unrelated workbook part changed ({name}).")
        refs = {f"{_column_letters(target_column)}{r}" for r in classifications}
        if append_column:
            refs.add(f"{_column_letters(target_column)}{header_row}")
        if _canonical_sheet_without_targets(before.read(sheet_path), refs, ignore_dimension=append_column) != _canonical_sheet_without_targets(after.read(sheet_path), refs, ignore_dimension=append_column):
            raise WorkbookSafetyError("Verification failed: non-target worksheet cells or structure changed.")
        strings_before, strings_after = _shared_strings(before), _shared_strings(after)
        root_after = ET.fromstring(after.read(sheet_path))
        for row, code in classifications.items():
            ref = f"{_column_letters(target_column)}{row}"
            cell = root_after.find(f".//{_xml_tag('c')}[@r='{ref}']")
            if cell is None or _cell_value(cell, strings_after) != code:
                raise WorkbookSafetyError(f"Verification failed: classification was not written to {ref}.")
        if append_column:
            ref = f"{_column_letters(target_column)}{header_row}"
            cell = root_after.find(f".//{_xml_tag('c')}[@r='{ref}']")
            if cell is None or _cell_value(cell, strings_after) != appended_header:
                raise WorkbookSafetyError("Verification failed: appended header was not written correctly.")


def _get_streamlit():
    try:
        import streamlit as st
    except ImportError as exc:
        raise RuntimeError("Install the app dependencies with: pip install streamlit polars openpyxl") from exc
    return st


def run_app() -> None:
    st = _get_streamlit()
    st.set_page_config(page_title="Attendance Workbook Cleaner", layout="wide")
    st.title("Attendance Workbook Classification")
    st.caption("Applies only the attendance codes and single-punch rule. The uploaded workbook is never overwritten.")
    uploaded = st.file_uploader("Upload an Excel .xlsx workbook", type=["xlsx"])
    if not uploaded:
        st.info("Upload a workbook to preview it and choose the attendance columns.")
        return
    file_bytes = uploaded.getvalue()
    try:
        with zipfile.ZipFile(BytesIO(file_bytes)) as book:
            sheet_names = list(_workbook_sheet_paths(book))
    except (zipfile.BadZipFile, WorkbookSafetyError) as exc:
        st.error(str(exc))
        return
    sheet_name = st.selectbox("Worksheet", sheet_names)
    header_row = st.number_input("Header row", min_value=1, value=1, step=1)
    try:
        headers, rows, max_col = read_sheet_rows(file_bytes, sheet_name, int(header_row))
    except WorkbookSafetyError as exc:
        st.error(str(exc))
        return
    if not headers:
        st.error("No columns were found on this worksheet.")
        return
    preview_frame = pl.DataFrame([{k: v for k, v in row.items() if k != "_excel_row"} for row in rows], strict=False) if rows else pl.DataFrame({h: [] for h in headers})
    st.subheader("Read-only preview")
    st.dataframe(preview_frame.head(20), use_container_width=True)
    col1, col2 = st.columns(2)
    with col1:
        shift_column = st.selectbox("Shift length column", headers)
        hours_column = st.selectbox("Worked hours column", headers)
        in_column = st.selectbox("In-punch column", headers)
    with col2:
        out_column = st.selectbox("Out-punch column", headers)
        append = st.checkbox("Append one classification column (only if no suitable column exists)", value=False)
        if append:
            target_column = None
            appended_header = st.text_input("New classification column heading", value="Attendance Classification")
        else:
            target_name = st.selectbox("Existing classification column to update", headers)
            target_column = headers.index(target_name) + 1
            appended_header = "Attendance Classification"
    st.subheader("Classification rules")
    st.dataframe(pl.DataFrame([{"Shift hours": s, **{k: v for k, v in rule.items()}} for s, rule in RULES.items()]), hide_index=True, use_container_width=True)
    st.caption("A single nonblank punch produces MP. Rows with no punches or missing/invalid shift or worked hours are left unchanged and flagged.")
    excel_fraction = st.checkbox("Interpret numeric worked-hours values as Excel day fractions", value=False)
    if not st.button("Preview proposed changes", type="primary"):
        return
    if len({shift_column, hours_column, in_column, out_column}) < 4:
        st.error("Select separate columns for shift length, worked hours, in-punch, and out-punch.")
        return
    if target_column is not None and headers[target_column - 1] in {shift_column, hours_column, in_column, out_column}:
        st.error("The output column must be different from the shift, hours, in-punch, and out-punch columns.")
        return
    try:
        result = classify_records(rows, shift_column=shift_column, hours_column=hours_column, in_column=in_column, out_column=out_column, excel_day_fraction=excel_fraction)
    except Exception as exc:  # Clear user feedback for malformed workbooks/data.
        st.error(f"Could not classify these rows: {exc}")
        return
    target_mapping = {record["_excel_row"]: result.proposed[i] for i, record in enumerate(rows) if result.proposed[i] is not None}
    current_codes = [None if append else record.get(headers[target_column - 1]) for record in rows]
    review = pl.DataFrame({"Excel row": result.row_numbers, "Current code": current_codes, "Proposed code": result.proposed, "Status": result.messages}, strict=False)
    st.dataframe(review, use_container_width=True)
    st.write(f"Ready to update {len(target_mapping)} classification cells. {len(rows) - len(target_mapping)} rows will remain unchanged and are flagged for review.")
    if not st.button("Create separate output workbook"):
        return
    try:
        output_bytes = create_modified_workbook(file_bytes, sheet_name=sheet_name, classifications=target_mapping, target_column=target_column, append_column=append, header_row=int(header_row), appended_header=appended_header)
    except (WorkbookSafetyError, zipfile.BadZipFile, ET.ParseError) as exc:
        st.error(f"No output was created: {exc}")
        return
    out_name = f"{Path(uploaded.name).stem}_attendance_classified.xlsx"
    st.success("Workbook verification passed. Only the approved classification cells were changed.")
    st.download_button("Download classified workbook", data=output_bytes, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    run_app()
