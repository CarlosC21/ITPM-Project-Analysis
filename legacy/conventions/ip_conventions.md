US Industrial Production workbook: conventions

Script: ip_update_workbook.py (standard library only; edits the file XML so the 42 charts survive). Master copy is the project doc; it is NOT in the owner's folder (the owner has only allowed that for the Employment script). Stage nothing from the folder except the workbook; write the script out from the project doc. Master workbook: US_Industrial_Production.xlsx in the owner's connected project folder (added Nov 9, 2026). Sample input: ip_values_2026-09.csv (real FRED values for May-Sep 2026, read in the built-in browser Nov 9, 2026).

Status (Nov 9, 2026)
Built and tested on COPIES only. The master has not been changed. Workbook ends Sep-26 (data of the Oct 16 release).
First real run: October 2026 data, G.17 release scheduled Nov 17, 2026 (FRED showed "next release Nov 17, 2026").
Owner decisions (Nov 9): (1) overwrite revised prior months within the last 5 months; (2) report the existing bad rows (below), fix nothing.
What it is

US Industrial Production, Federal Reserve G.17 release, index 2017=100. 21 sheets, one per FRED series, in this order: Industrial_Production=INDPRO | Manufacturing=IPMAN | Food_Bev_Tob=IPG311A2S | Textiles=IPG313A4S | Chemicals=IPG325S | Plastics_Rubber=IPG326S | Apparel_Leather=IPG315A6S | Wood=IPG321S | Paper=IPG322S | Printing_Related=IPG323S | Petrol_Coal=IPG324S | Nonmetallic_Minerals=IPG327S | Primary_Metals=IPG331S | Fabricated_Metals=IPG332S | Machinery=IPG333S | Comp_Elec=IPG334S | ElecEquip_Appliances_Components=IPG335S | MotorVehicles_Parts=IPG3361T3S | Aero_MiscTransport=IPG3364T9N | Furniture_Related=IPG337S | Misc=IPG339S. All are seasonally adjusted EXCEPT IPG3364T9N (Aero), which is NOT seasonally adjusted (FRED title/page, Nov 9, 2026); the owner already used it, so it stays. Sheet order = the owner's list of 21 FRED pages.

Layout (verified Nov 9, 2026)
Row 1: B Index, C MoM, D YoY (A1 empty). Data from row 2, NEWEST FIRST; A = month (mmm-yy). Values hard-coded, no formulas.
25 rows per sheet (Sep-24 to Sep-26) except Aero_MiscTransport, which keeps history down to Jan-1973 (645 rows). The window grows: a new row is inserted at row 2 and nothing is dropped.
2 line charts per sheet (42), ranges already run to row 1006, so the script leaves chart parts untouched.
MoM, YoY are fractions shown as 0.00%.
Source and how to read it
FRED (owner's choice; it republishes G.17). The Fed's G.17 PDF shows ONE decimal only, so it cannot feed 4-decimal cells; use it only as a cross-check of the first decimal (Table 4, SA; its 3364-9 row is seasonally adjusted, unlike the workbook's).
DO NOT use WebFetch on fred.stlouisfed.org/series/...: on Nov 9, 2026 it returned stale copies (14 of 21 series stopped at Jun-Aug while the live pages had Sep). Use the BUILT-IN BROWSER (site access on fred.stlouisfed.org must be granted; the shell and the WebFetch of the bulk CSV are blocked).
In the browser, open https://fred.stlouisfed.org/data/IPG325S.txt (any series page; it is an HTML table), then run in javascript_tool: for each id, fetch('/data/'+id+'.txt'), parse with DOMParser, take body.innerText, regex /(\d{4}-\d{2})-01\s+([\d.]+)/g, sort by month and keep the last 5. Output lines series,month,value.
EXCEPTION: /data/INDPRO.txt only lists up to 2002 in that parse. Read INDPRO from fetch('/series/INDPRO') text: regex /((?:Jan|...|Dec) 20\d\d)[:\s]+([\d.]{6,})/g, first 5 distinct months (seen Nov 9: Sep 2026 103.9506 ...). Check that "Last Updated" and the latest month are the same for all 21.
If the browser cannot be reached or a series has no new month yet, STOP and tell the owner; never fill from another site.
Release schedule

Mid-month, about 9:15 a.m. ET (G.17). Sep data: Oct 16, 2026. Next: Nov 17, 2026 (Oct data) per FRED; confirm on the Fed calendar. Revisions: each release restates recent months (and the Fed restates more in its annual revision; see Stops).

values.csv

series,month,value; series = FRED id; months YYYY-MM. For EVERY series the new month and the previous months FRED shows, up to 5 months in all (new month + 4 prior). Must include the new month and the one before it.

Calculation (new row)

MoM = index / previous month index - 1; YoY = index / same month last year index - 1 (previous index = the revised value). The new row's date is the 1st of the month. Revisions: any of the prior months in values.csv whose index differs from the workbook by more than 0.00005 is overwritten. The script then recomputes ONLY the cells that depend on it: that month's MoM and YoY, the next month's MoM, the month 12 later's YoY. Everything else is untouched. Nov 9 dry run: 63 revisions (May/Jun/Jul-26) and 147 recomputed cells.

Validation (STOP = nothing written)
Workbook: 21 sheets in order, 42 charts, months consecutive newest-first (all rows), no empty index, every sheet ends on the same month.
values.csv: known series only, numbers, no duplicates, all 21 series, new month = workbook latest + 1 for EVERY series (if one series lags, e.g. Aero, the run stops and names it), previous month present, nothing older than 5 months or after the new month.
Prior months must exist in the workbook; a revision above 5% stops (annual revision, wrong series or typo).
New value positive and within 20% of the previous month.
Proof after building: months only +1, nothing changed except revised indexes and their dependent cells; every non-sheet part of the file is byte-identical; output re-opens with 42 charts.
Known problems in the existing data (reported every run, NOT changed - owner decision Nov 9)
Apr-26 to Sep-26 rows are dated the 26th (e.g. 26-Sep-2026), not the 1st (typed "Sep-26" in Excel). Months are matched by year-month, so this is harmless to the script. New rows use the 1st.
Those same six rows have MoM and YoY that do NOT equal the index ratios on any sheet (INDPRO Sep-26 MoM shows -2.76%, the index gives +0.66%). Older rows match. The run's audit lists the count (210 cells on Nov 9) and the first 12. Revised months get recomputed (correct) values, so after the first run some of these rows will be right and others not.
May/Jun/Jul-26 indexes were old vintages vs live FRED (e.g. INDPRO Jul 102.9939 vs 103.2006); the first run overwrites them.
Primary_Metals Sep-26 holds 101.944 where FRED has 101.9444 (will be listed as a tiny revision).
Misc!E1 holds a stray "s". Aero_MiscTransport carries 620 extra history rows below row 26.
Chart caches are not rebuilt (ranges run to row 1006; Excel refreshes on open). Tested in LibreOffice, not Excel.
Regression test (Nov 9, 2026)

Copy the master, remove row 2 on every sheet (Sep-26) to rewind to Aug-26, run with ip_values_2026-09.csv. Expected: Sep-26 indexes equal the master's (except Primary_Metals 101.944 -> 101.9444), 63 revisions, rows older than the window unchanged, 42 charts, opens in LibreOffice. Also tested: missing series, one series lagging, +51% move, 10.8% revision, unknown series, duplicate row, text value, missing previous month (all STOP, rc=2), and --in-place (backup identical to original). Running the sample csv on the real master stops with "Oct-26 is not in values.csv" (correct).

Compare script

Not built. Parallel run: if the owner keeps a manual workbook, ask for it and match it cell for cell (index first, then MoM/YoY).

Open items
Trust criteria in 01_INDICATOR_INDEX.md apply (3 consecutive months matching the manual workbook).
Confirm Nov 17 release time on the Fed calendar.
Decide later whether to repair the six day-26 rows (before/after list first).