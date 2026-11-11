# WTDC KPI Dashboard

## Open the project

The primary Windows checkout is `C:\Users\Marjorie\KPI-Dashboard`.
Double-click `KPI-Dashboard.code-workspace`, or open that folder in VS Code.
Use this checkout for edits, commits and pushes. The earlier OneDrive workspace
is historical and is not the working location.

## Daily workflow

1. Open the workspace and VS Code Terminal.
2. On first setup, or after dependency changes, run `npm ci`.
3. Edit `app/page.tsx` (dashboard), `app/globals.css` (design), or
   `app/interactive-report.ts` (customer export).
4. Run `npm test`. This builds `index.html` and checks the standalone bundle.
5. Run `npm run preview`, then open http://127.0.0.1:4173. Refresh after rebuilding.
   Stop the preview with Ctrl+C. You can also open `index.html` directly offline.
6. Review Source Control in VS Code, commit the intended code files, then Push
   or Sync Changes. Work on a feature branch and merge through a pull request.

VS Code's Terminal > Run Task also provides install, build, verify and preview.
Ctrl+Shift+B runs the default build. The workspace uses LF line endings.

## Architecture

This is a browser-only React application, compiled into one portable HTML file.
It needs no backend, account, database, cloud storage or hosted runtime.
The local preview binds only to 127.0.0.1 and serves only the compiled dashboard;
it does not expose the input or output directories.

| Path | Purpose |
| --- | --- |
| `app/page.tsx` | Original dashboard UI, Excel parsing and interactions |
| `app/globals.css` | Original dashboard styling |
| `app/interactive-report.ts` | Preloaded read-only customer HTML export |
| `app/offline-entry.tsx` | Browser entry point |
| `public/wtdc-logo.png` | WTDC logo, embedded during build |
| `scripts/` | Build, verification and restricted local preview |
| `index.html` | Generated, committed standalone internal application |
| `package-lock.json` | Exact dependency resolution; use npm ci |
| `input/` | Local customer workbooks; ignored by Git |
| `output/` | Local reports and verification artifacts; ignored by Git |
| `sample_data/` | Approved dummy workbook only |
| `.github/workflows/verify.yml` | Code-only CI build and verification |

No Next.js server or Cloudflare/Sites runtime is used. Node 22.13+ is required
only for development. Customers do not need Node, Python or the Excel workbook.
Do not edit generated `index.html` manually; rebuild it from `app/`.

## Monthly customer delivery

Open the internal dashboard, choose the month's Excel workbook, review the data,
then select Export interactive report. Send the resulting HTML through WTDC's
approved delivery channel. Customers receive a fixed data snapshot and can
filter it without uploading files. The internal dashboard starts blank.

The current export retains the original site's export design, which differs
from its internal dashboard design. The internal application's component,
stylesheet, logo and export function came from Sites source commit
`2150586d2693d60133b8870838550738f6adad5d`.

Customer workbooks, generated customer reports, credentials and real data must
never be committed or used in cloud CI. Keep them in ignored folders. Gitignore
prevents ordinary staging, but can be overridden; inspect staged changes before
pushing. Downloaded HTML is editable by its recipient, so retain the approved
original internally when an authoritative record is needed.

## Verification scope

`npm test` verifies compilation, JavaScript syntax, embedded assets and rendered
blank-screen sections. It does not replace browser QA. Before a customer release,
check workbook loading, filters, charts, export/reopening offline, and printing.
CI runs only code and dummy-data checks; it does not publish or deploy reports.

## Earlier Python generator

`build_report.py`, `templates/report_template.html`, `requirements.txt` and
`test_report.py` are a separate legacy report generator with a different layout.
They are not needed for the primary React application. If maintaining it, use a
local Python virtual environment and run `python -m unittest -v`. It blocks output
on source issues unless `--draft` is requested. Do not mix its output with the
primary dashboard's customer exports.
