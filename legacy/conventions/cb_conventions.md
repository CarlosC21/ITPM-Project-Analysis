Corporate Bond Indices - US workbook: conventions

Script: cb_update_workbook.py (standard library only; edits the sheet XML so the 20 charts survive). Master copy of the script is the project doc (claude/ folder); it is NOT in the owner's folder. Stage only the workbook; write the script out from the project. Master workbook: Corporate Bond Indices - US.xlsx in the owner's connected project folder. Sample input: cb_values_2026-11-06.csv (real FRED values, Oct 13 - Nov 6, 2026, read in the built-in browser Nov 9, 2026).

Status (Nov 9, 2026)
Built Nov 9, 2026. Master updated the same day: 2026-10-23 to 2026-11-05 added (backup backups/Corporate Bond Indices - US_before_2026-10-23.xlsx). Master now ends 2026-11-05.
2026-11-06 is NOT addable yet: FRED had no DGS10 value for it on Nov 9 (Treasury series lags the ICE series by about a day).
What it is

Daily US corporate bond yields from FRED, newest first, plus six yield spreads (percentage points, not bps).

Sheet	Content
US Corp Yield Indices	B date, C US10YR (DGS10), D ICE_BoA_AAA (BAMLC0A1CAAAEY), E ICE_BoA_BBB (BAMLC0A4CBBBEY), F ICE_BoA_CCC_or_below (BAMLH0A3HYCEY). Hidden column A empty. Header row 2, data from row 3. 17,002 rows (10Y from 1962-01-02; ICE series from 1996-12-31). 8 charts.
US Corp Yield Spreads	A date, B AAA-BBB, C AAA-CCC, D BBB-CCC, E 10yr-AAA, F 10yr-BBB, G 10yr-CCC. Header row 2, data from row 3, hard-coded values (no formulas). Starts 2026-01-01 (207 rows), then blank formatted rows to row 10115. 12 charts.
Sheet2	The owner's scratch area (aaa, bbb, ccc, 10yr pasted, spread formulas). Not touched.
Sheet1	Empty. Not touched.
Spread = first minus second as named (AAA-BBB = AAA - BBB; 10yr-AAA = 10Y - AAA). Values are the raw float difference, same as Excel's formula.	
Source and how to read it
FRED (owner's choice). Series: DGS10, BAMLC0A1CAAAEY, BAMLC0A4CBBBEY, BAMLH0A3HYCEY. Percent, not seasonally adjusted, daily. Titles checked on the pages Nov 9, 2026.
Read in the BUILT-IN BROWSER (site access for fred.stlouisfed.org via request_access). The shell cannot reach FRED; WebFetch of FRED is unreliable.
Do NOT use /data/<id>.txt: DGS10 shows only the first 1,000 rows (1962-65) and the ICE series only 3 years from 2023-11-10 (FRED limits ICE history).
Use the graph CSV, from javascript_tool on any fred.stlouisfed.org page: fetch('/graph/fredgraph.csv?id=<ID>&cosd=<workbook latest date minus 7 days>&coed=<today>') for each of the four ids. Columns observation_date,<ID>; a blank value means no data that day. Merge by DATE (never by row position) into date,US10YR,AAA,BBB,CCC. Worked on Nov 9, 2026 for all four.
If the browser cannot be reached or a series has nothing new, STOP; never fill from another site.
Release schedule

Every US business day. ICE series appear the next morning; DGS10 (Fed H.15) can lag another day. On Nov 9 (Mon): ICE to Nov 6, DGS10 to Nov 5. FRED shows the loaded dates on each series page ("Last Updated"); exact times unverified.

Rules for which dates are added
A date is added only when all four values exist. The owner's workbook leaves out days where the Treasury market is closed but ICE has a value (e.g. 2026-10-12, Columbus Day); the script does the same and lists them as skipped.
Newest dates where any series is still blank are "pending": not added, reported, picked up on the next run.
A blank in the middle of the file that is not "10Y blank, other three present" STOPs.
values.csv

date,US10YR,AAA,BBB,CCC, YYYY-MM-DD, blank for no value. Must include the workbook's latest date and ideally 3-5 earlier dates (overlap check), then every new date up to today.

Calculation and writing

New rows go in at row 3 of BOTH data sheets (newest first), everything below shifts down. Indices keeps its history; on Spreads the same number of trailing blank rows is dropped so the sheet keeps its size (stops if those rows hold data). Date = Excel serial, styles copied from existing rows (date s=8; yields s=1; spreads s=7). Nothing else is written.

Validation (STOP = nothing written)
Workbook: sheet names and order, header labels on both sheets, 20 charts, newest date equal on both sheets.
values.csv: header, numeric values within range (10Y/AAA/BBB 0-20, CCC 0-60), no duplicate or weekend dates, workbook's latest date present.
Overlap: the workbook's newest 5 dates must equal FRED exactly, otherwise STOP (workbook or FRED has changed; ask the owner). Older overlap differences are only reported.
One-day move above 1.0 point (10Y, AAA, BBB) or 3.0 (CCC) STOPs. AAA <= BBB <= CCC order is a warning.
Proof after building: old rows on both sheets identical, shifted by the number of new rows; new rows equal the input and the six spreads; only the two data sheets differ from the original, all other parts (charts, styles, Sheet2) byte-identical.
--in-place writes backups/Corporate Bond Indices - US_before_<first new date>.xlsx (checked identical to the original) before replacing the master.
Data repair (Nov 9, 2026, owner asked for it) and what is still open

Repaired in the master (backup backups/Corporate Bond Indices - US_before_repair_2026-11-09.xlsx; full cell-by-cell list in cb_repair_log_2026-11-09.csv; one-off script cb_repair.py):

AAA/BBB/CCC (and one 10Y) values from 2026-05-07 to 2026-10-08 set to FRED by DATE: 160 Indices cells on 70 dates. Cause looked like pasting by row position around days when only the ICE series have a value (May 25, May 31, Jun 19, Jul 3, Sep 7, Oct 12).
2026-10-09 inserted on both sheets (it was missing).
306 Spreads cells recomputed from the corrected yields (every Spreads row from 2026-04-01 now equals Indices).
Chart ranges: 9 charts started at row 57 / 37 / 4 (the newest 34-54 days were outside them); all now start at row 3. "Last 5 Years" charts still end at row 1454 (a rolling window of about 1,450 rows, as the owner built it), full-history charts end at 10115. Checked before repairing: for every month from 1962 to 2026-04 the workbook equals FRED in row count and in the sum of every series (after setting aside the filler rows below). Only 2026-05 to 2026-10 differed. Left as they are (owner's earlier practice, not errors):
816 filler rows from 1962 to 2026-02: US holidays and month-end weekends (e.g. 2026-02-28, 2026-01-31, 2025-11-30) that carry a value, which FRED does not have for those days. The owner stopped adding such rows after Feb 2026; the script follows the newer practice (only days when all four series have a value).
ICE values exist for weekend month-ends and US holidays on FRED but are left out from March 2026 on (nothing is added for days when DGS10 is blank).
The Spreads sheet starts at 2026-01-01 although its full-history charts reach row 10115; spreads for earlier years (ICE data from 1996) are not in the workbook. Ask the owner if they want them back (about 7,800 rows).
The 10Y full-history Indices charts end at row 10115 while the sheet now has 17,013 rows, so they stop around 1983.
Regression test (Nov 9, 2026)

Copy of the master + cb_values_2026-11-06.csv: 10 rows added (2026-10-23 to 2026-11-05) to each sheet, 2026-11-06 reported pending, proof passed, opens in LibreOffice, re-running on the output adds nothing. 9 failure cases STOP (missing overlap date, overlap mismatch, 1.5-point jump, text value, duplicate date, hole in the middle, bad header, weekend date) or report nothing to add (only incomplete dates).

Compare script

Not built. Parallel run: if the owner updates a manual workbook the same days, ask for it and compare cell for cell (Indices B:F, Spreads A:G).

Open items
Owner to decide: fix the Jun-Oct 2026 values and the missing Oct 9 (before/after list first); chart ranges.
Confirm the daily vs weekly rhythm the owner wants (script handles any number of new days per run).
First real run on the master; Excel (not just LibreOffice) check of the output.
Trust criteria in 01_INDICATOR_INDEX.md apply.