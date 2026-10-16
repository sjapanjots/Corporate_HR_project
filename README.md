# Attendance Workbook Classification

This Streamlit app applies the attendance codes from the handwritten rules to an `.xlsx` workbook. Polars is used for the row-based attendance calculations.

## Run

From this folder, install the dependencies and start Streamlit:

```powershell
python -m pip install -r requirements.txt
streamlit run App.py
```

## Use

1. Upload an `.xlsx` workbook and select a worksheet and header row.
2. Map the shift length, worked-hours, in-punch, and out-punch columns.
3. Select an existing attendance classification column or explicitly opt in to appending one.
4. Review the thresholds and proposed row changes, then create and download a separate output workbook.

Only classification cells with a valid proposed code are updated. A single punch receives `MP`. Rows with no punches or missing/unsupported required values are left unchanged and flagged for review. The source upload is never overwritten.

The app patches the selected worksheet XML and verifies that other workbook parts and all non-target worksheet cells remain unchanged. If it cannot verify the workbook, it does not provide an output file.

## Attendance codes

| Shift | A | HF | SP | P | PP |
|---|---:|---:|---:|---:|---:|
| 9 hours | `< 5` | `5 to < 8` | `8 to < 8:45` | `8:45 to 17` | `> 17` |
| 10 hours | `< 5` | `5 to < 9` | `9 to < 9:45` | `9:45 to 18` | `> 18` |
| 11 hours | `< 5` | `5 to < 10` | `10 to < 10:45` | `10:45 to 20` | `> 20` |
| 12 hours | `< 6` | `6 to < 10` | `10 to < 11:45` | `11:45 to 22` | `> 22` |

Hours in the table are worked hours; time values such as `10:45` mean 10 hours and 45 minutes. Numeric Excel day fractions can be enabled in the app when the worked-hours column uses that representation.
