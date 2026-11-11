# WTDC monthly customer report builder

Turns the internal monthly KPI workbook into one self-contained, read-only HTML
report for the customer. Everything runs on a WTDC computer. Nothing is uploaded
or hosted anywhere.

## Rule number one
This repository holds code only. Real customer workbooks go in `input/` and
finished reports come out in `output/`. Both folders are blocked from Git by
`.gitignore`. This prevents ordinary Git staging; it is not a security boundary
and can be overridden with `git add -f`. Do not force-add customer data. The only
workbook in the repo is the dummy file in `sample_data/`.

## One-time setup
1. Install Python 3.10 or newer.
2. In this folder, run: `pip install -r requirements.txt`

## Each month
1. Save the month's workbook into `input/`.
2. Run: `python build_report.py input/Campari_October_2026_KPI.xlsx`
3. Read the checks printed in the terminal. Any CHECK blocks report generation
   and exits with code 2. Fix the workbook and run again. For internal review,
   add `--draft` to generate a file with a DRAFT filename and visible banner;
   this still exits with code 2. Do not send drafts or stale reports from a
   previous run. Only send the file named by a successful current build.
4. Open the report from `output/` in a browser and skim it.
5. Send it to the customer, then archive the report and its workbook
   on the internal drive.

## Try it with dummy data
`python build_report.py sample_data/Sample_Campari_September_2026_KPI.xlsx --draft`

The supplied sample has two orders with confirmation dates after pick dates.
These remain visible in the order register but are excluded from step averages.
Creation-to-delivery time is calculated directly from the endpoint dates.

## Delivery architecture

Use GitHub for reusable code and VS Code for development. Generate a single HTML
file on a WTDC computer and send it through your approved customer delivery
channel. No GitHub Pages, hosted app, database, cloud build, analytics, CDN, or
external fonts are needed. The finished file includes its own data, styles,
scripts and font. A Content Security Policy blocks network connections.

Customers can search, sort, expand product/receipt details, compare years and
print. They cannot upload or edit data through the interface. A downloaded HTML
file is not tamper-proof: recipients can modify their own copy in an editor.
Keep the approved original internally if you need an authoritative record.

Every section has calculation notes; Definitions lists source worksheets and
the workbook SHA-256 fingerprint. Filters affect their table only. Report totals
stay fixed. Aging uses only supplied aging rows older than 365 days at period
end; the rest of closing inventory is not assumed to have verified receipt ages.
Orders are scoped by delivery date, including orders created in prior months.

## Validation

Run `python -m unittest -v`. Before delivery, open the generated HTML locally in
Edge or Chrome and review desktop/mobile layout, search, sort, row expansion,
year comparison and Print / Save PDF. Check offline behavior with networking off.
The parser expects this workbook's worksheet names and column layout. Excel
formulas must have current cached values: recalculate and save in Excel first.

## Folder layout
| Path | What it is |
| --- | --- |
| `build_report.py` | Reads the workbook, runs the checks, writes the report |
| `templates/report_template.html` | The report design, the same for every month |
| `sample_data/` | Dummy workbook for testing changes |
| `input/` | Real workbooks (never committed) |
| `output/` | Finished reports (never committed) |
