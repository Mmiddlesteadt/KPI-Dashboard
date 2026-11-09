# WTDC monthly customer report builder

Turns the internal monthly KPI workbook into one self-contained, read-only HTML
report for the customer. Everything runs on a WTDC computer. Nothing is uploaded
or hosted anywhere.

## Rule number one
This repository holds code only. Real customer workbooks go in `input/` and
finished reports come out in `output/`. Both folders are blocked from Git by
`.gitignore`, so customer data can never be pushed by accident. The only
workbook in the repo is the dummy file in `sample_data/`.

## One-time setup
1. Install Python 3.10 or newer.
2. In this folder, run: `pip install -r requirements.txt`

## Each month
1. Save the month's workbook into `input/`.
2. Run: `python build_report.py input/Campari_October_2026_KPI.xlsx`
3. Read the checks printed in the terminal. Every line must say OK.
   Any CHECK line means fix the workbook and run again.
4. Open the report from `output/` in a browser and skim it.
5. Send it to the customer, then archive the report and its workbook
   on the internal drive.

## Try it with dummy data
`python build_report.py sample_data/Sample_Campari_September_2026_KPI.xlsx`

## Folder layout
| Path | What it is |
| --- | --- |
| `build_report.py` | Reads the workbook, runs the checks, writes the report |
| `templates/report_template.html` | The report design, the same for every month |
| `sample_data/` | Dummy workbook for testing changes |
| `input/` | Real workbooks (never committed) |
| `output/` | Finished reports (never committed) |
