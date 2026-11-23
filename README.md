# New Home Sales Weekly Import

Reads the weekly SoCal new home sales report PDF and updates `New_Home_Sales_Database.xlsx`:
**Gross Sales** (Wk's Sales), **Cancellations** (Wk's Cans) and **Net Sales** (Gross - Cancellations, live formulas).
Net Sales also carries Project Label, Pace Label (`46 of 79 sold | 0.00/wk (4 wk) | 0.23/wk (13 wk) | 0.52/wk (LOJ)`)
and LOJ Pace (life-of-job: average of all weekly net sales).

## Setup (once)
1. Install Python 3.10+ (python.org).
2. Clone this repo, then copy your `New_Home_Sales_Database.xlsx` into the repo folder next to `app.py`.
   Workbooks and PDFs are git-ignored, so your data never gets committed.
3. Double-click **Run New Home Sales Import (Windows).bat** or **Run New Home Sales Import (Mac).command**
   (or run `pip install -r requirements.txt` then `streamlit run app.py`). A browser tab opens.

## Every week
1. In the app, the database path defaults to `New_Home_Sales_Database.xlsx` in this folder.
2. Drop in the new weekly report PDF (and GPN report, if any). Review: totals check (every section vs the PDF's TOTALS row), market totals, and the change list.
3. Click **Import into database**. The workbook is saved in place; a download copy is also offered.
   Close the workbook in Excel before importing so it can be saved.

## Backfilling older reports
Reports can be imported in any order, including weeks before 03/02/2025. Drop several PDFs in at once;
each lands in its own week column, and the grids extend back automatically. Older reports never overwrite
current names/units/status - those always come from the newest report. Flags (NEW / DROPPED / RETURNED)
are re-evaluated across the whole timeline after every import, so backfilling doesn't create false flags.
If an older PDF used a different layout, the totals check will show a mismatch before you import.

## Great Park Neighborhoods (GPN) report
Drop the GPN "General Distribution Report" PDF into the same uploader - it's recognised automatically.
- GPN numbers **replace** the weekly report for the same project and week (e.g. Moonlight, Isla, Rhea), in either import order.
  The weekly report row is kept in Weekly_Data as "Weekly report (superseded by GPN)" for audit.
- Projects the weekly report doesn't carry (Toll Brothers at Great Park) get their weekly numbers from GPN; their Tracking becomes "GPN".
- Names are matched automatically (builder + collection, e.g. "KB, Moonlight" -> Moonlight at Luna Park). If one
  doesn't match, add a Name_Map row with the alias exactly as GPN prints it (e.g. `KB, Moonlight`) and the Community ID.
- Collections that sold out before the report week and aren't in the database are skipped.
- GPN numbers show in bold teal in the weekly tabs. Each neighborhood is checked against the report's own total row.

## Latitude / Longitude
All three weekly tabs show Latitude and Longitude (8 decimals) after City. The weekly report has no addresses, so:
- Every project starts at its **city center** (grey italic) from the bundled `ca_places.csv` (GeoNames).
- The app's **Locations** tab has **Find locations**, which looks each project up by name on OpenStreetMap
  (one lookup per second, needs internet). Matches are checked to be within 25 km of the city.
- Coordinates typed in (Locations tab, Communities tab, or any weekly tab) are marked Manual and never overwritten.
The Communities tab's Geo Source column says which level each project has.

## Projects that didn't report (N/R)
When The report lists a project with Traffic "N/A", it didn't report that week. The tool records it as **N/R**
(on the report, no numbers), not as 0 sales, so it is left out of totals and pace and is not flagged as sold out.

## What gets flagged
| Event | Meaning |
|---|---|
| NEW | First appearance on the weekly report (also when the report notes it as "New") |
| ADDED | First week tracked here, in a submarket with no earlier data |
| DROPPED | On last week's report, missing this week - likely sold out / temporarily sold out. Shows SOLD OUT from then on; never deleted |
| RETURNED | Back on the report after being absent |
| FLAG S/O, FLAG TSO | Still listed, but the report marks it Sold Out / Temporarily Sold Out |
| POSSIBLE RENAME | A new and a dropped project with the same builder and similar names. Add a Name_Map row and re-import |

## Your editable parts (kept on every import)
- **Project Label** (Net Sales column F): your custom project name. Edit it in the app's **Project Labels** tab,
  directly in Net Sales column F, or in the Communities tab - whichever you changed is kept.
- **Communities** tab, yellow columns: Project Label, Total Lots, Grand Opening, Tracking (Weekly report / Manual / GPN), My Notes.
- **Name_Map**: alias -> existing Community ID.
- **Manual_Entries**: weekly gross/cans for communities not on the weekly report (Tracking = Manual, e.g. Toll Brothers at Great Park).

The three weekly grids are regenerated on each import, so don't type into them.

## Online database later
`Weekly_Data` (one row per community per week) and `Communities` are already table-shaped; the app's Database tab exports both as CSV. `sales_db.py` is the only module that touches storage - swap `load_db`/`save_db` for database calls.

## Files
- `weekly_report_parser.py` - PDF -> records (word-position based; validates every section's totals)
- `gpn_parser.py` - Great Park Neighborhoods report -> records (validates each neighborhood total)
- `geocode.py` + `ca_places.csv` - city-center coordinates and OpenStreetMap lookup
- `sales_db.py` - matching, new/dropped/returned logic, legacy import
- `workbook_writer.py` - builds the Excel tabs
- `app.py` - Streamlit front end
- `build_initial.py` - one-time rebuild from a PDF + the old v4 trends workbook

## Updating the code
`git pull` in this folder picks up new versions. Your workbook isn't touched by git.
