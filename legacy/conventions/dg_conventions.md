US Durable Goods workbook: conventions

Script: dg_update_workbook.py (standard library only; edits the file XML so the 20 charts survive; also rewrites the Summary sheet). Master copy is the project doc (claude/ folder); it is NOT in the owner's folder. Stage only the workbook from the folder; write the script out from the project. Master workbook: US_Durable_Goods.xlsx in the owner's connected project folder. Sample input: dg_values_2026-09.csv (real FRED values, Apr-Sep 2026, read in the built-in browser Nov 9, 2026).

Status (Nov 9, 2026)
Built Nov 9, 2026. September 2026 added to the master the same day (backup backups/US_Durable_Goods_before_2026-09.xlsx). Workbook now ends Sep-26.
Next run: October 2026 data, after the full M3 report (FRED next release Nov 25, 2026; confirm on the Census calendar). Next release: Nov 25, 2026 (October data) per FRED "Next Release Date"; confirm on the Census calendar.
What it is

US Manufacturers' New Orders (Census Bureau M3 / advance durable goods), millions of dollars, seasonally adjusted (FRED units line, Nov 9, 2026). 10 data sheets, one per FRED series, after a Summary sheet (11 sheets in all), in this order: DG=DGORDER | DG_Ex_Transport=ADXTNO | DG_Ex_Defense=ADXDNO | Comp_and_Electronic_Prod=A34SNO | Elec_Equip_Apps_Comps=A35SNO | Machinery=A33SNO | Fab_Metal_Prods=A32SNO | Primary_Metals=A31SNO | Transportation_Equipment=A36SNO | Furniture_and_Rel_Prods=A37SNO. Series titles were checked against the FRED pages on Nov 9, 2026 (all ten match the sheet names).

Layout (verified Nov 9, 2026)
Data sheets: row 1 B Index, C MoM, D YoY (A1 empty). Data from row 2, NEWEST FIRST; A = month (mmm-yy). Values hard-coded, no formulas.
24 rows per sheet (Sep-24 to Aug-26). The window grows: a new row is inserted at row 2 and nothing is dropped.
2 line charts per sheet (20 in all: Index and YoY), ranges already run to row 1005/1006, so the script leaves chart parts untouched.
MoM, YoY are fractions shown as 0.00%.
Summary sheet: hard-coded MoM/YoY, 10 rows in data-sheet order. B = latest month MoM, C = prior month MoM, D = latest YoY, E = prior YoY; headers are text like "Sep-2026 MoM" (shared strings). The script rewrites it on every run from the data sheets and checks each cell equals its source. --summary-only --in-place refreshes it alone (backup ..._before_YYYY-MM_summary.xlsx). Label typo "Durbale" in A3 left as is.
Source and how to read it
FRED (owner's choice; it republishes the Census release). Read in the BUILT-IN BROWSER (site access for fred.stlouisfed.org must be granted with request_access). Do not use WebFetch on FRED (stale copies, see ip_conventions.md); the shell cannot reach it.
In the browser, navigate to https://fred.stlouisfed.org/data/DGORDER.txt, then in javascript_tool: for each id, fetch('/data/'+id+'.txt'), parse with DOMParser, take body.innerText, regex /(\d{4}-\d{2})-01\s+([\d.,]+)/g, sort by month, keep the last 6. Output series,month,value lines (this worked for all 10 on Nov 9). Also read /series/<id> for "Updated:" and "Next Release Date".
Check that the latest month is the same for all 10. NOTE: A35SNO showed "Updated Oct 27, 2026" while the other nine showed "Nov 3, 2026" - its September value (18854) may be the advance-report figure and may be revised at the next release. The run is still allowed; mention it.
If the browser cannot be reached or a series has no new month, STOP and tell the owner; never fill from another site.
Release schedule

Advance durable goods report about the 25th-27th (8:30 a.m. ET), full M3 report about the 3rd-5th of the following month. FRED Sep data: Nov 3, 2026. Next: Nov 25, 2026 (unverified; FRED). Run after the full M3 report so all ten series have the new month. Revisions: each release restates recent months; the annual revision restates more.

values.csv

series,month,value; series = FRED id; months YYYY-MM. For EVERY series the new month and the previous months FRED shows, up to 6 months in all (new month + 5 prior). Must include the new month and the one before it.

Calculation (new row)

MoM = index / previous month index - 1; YoY = index / same month last year index - 1 (previous index = the revised value). New row date = the 1st of the month. Any prior month in values.csv whose index differs from the workbook by more than 0.5 (million $) is overwritten. The script then recomputes ONLY the cells that depend on it (that month's MoM and YoY, the next month's MoM, the month 12 later's YoY).

Validation (STOP = nothing written)
Workbook: sheets Summary + the 10 in order, 20 charts, months consecutive newest-first, no empty index, all sheets end on the same month.
values.csv: known series only, numbers, no duplicates, all 10 series, new month = workbook latest + 1 for EVERY series, previous month present, nothing older than 6 months or after the new month.
Prior months must exist in the workbook; a revision above 5% stops. New value positive and within 30% of the previous month (Transportation_Equipment moved +22% in Apr-26 on aircraft orders, hence 30%).
Proof after building: months only +1, nothing changed except revised indexes and their dependent cells; every non-sheet part except the Summary sheet and sharedStrings (its headers) is byte-identical; Summary cells are checked equal to the data sheets; Summary headers must look like 'Mon-YYYY MoM/YoY' or it stops; output re-opens with 20 charts.
Known problems in the existing data (reported every run, NOT changed)
Apr-26 to Aug-26 rows are dated the 26th (Furniture: Mar-26 to Aug-26), not the 1st (typed "Aug-26" in Excel). Months are matched by year-month, so harmless to the script. New rows use the 1st.
Those rows' MoM and YoY do NOT equal the index ratios on any sheet (DG Aug-26 YoY shows -13.82%, the index gives +8.43%; 102 cells in all on Nov 9). Revised months get recomputed (correct) values, so the first run corrects 100 of them; the rest are listed in the audit.
Summary sheet was stale (Mar/Feb 2026) and had a blank Furniture row until Nov 9, 2026; now refreshed.
Regression test (Nov 9, 2026)

Copy of the master + dg_values_2026-09.csv: Sep-26 added to 10 sheets (DG 345,672, MoM 1.91%, YoY 10.04%), 50 revisions (Apr-Aug 2026, all under 1.2%), 100 MoM/YoY cells recomputed, only sheet2-11 XML changed, Summary identical, 20 charts, opens in LibreOffice. Also tested (all STOP, rc=2): missing series, one series lagging, +53% move, 9.9% revision, text value, unknown series, duplicate row, missing previous month, re-running on the output (needs Oct-26). --in-place makes backups/US_Durable_Goods_before_2026-09.xlsx identical to the original.

Compare script

Not built. Parallel run: if the owner keeps a manual workbook, ask for it and match cell for cell (index first, then MoM/YoY).

Open items
Summary sheet is now updated by the script (owner approved Nov 9, 2026; refreshed to Sep/Aug 2026).
Decide whether to repair the day-26 dates and the 102 wrong MoM/YoY cells (before/after list first).
Confirm the Nov 25 release on the Census calendar. Excel (not just LibreOffice) check of the first real output.
Trust criteria in 01_INDICATOR_INDEX.md apply.