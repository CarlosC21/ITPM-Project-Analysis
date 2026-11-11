US CPI workbook: conventions

Script: cpi_update_workbook.py (no other scripts; it works at file level so the 20 charts are kept). Master workbook: US_CPI.xlsx in the owner's connected project folder. Folder access is untested.

Layout (verified Nov 2026)
10 sheets: CPI_All, CPI_All_Less_Food_Energy, Food_Beverages, Energy, Housing, Apparel, Transportation, Medical_Care, Recreation, Education.
Each sheet: header in row 1 (B Index, C MoM, D YoY); data from row 2, NEWEST FIRST; A = month (format mmm-yy). Two charts per sheet read A/B and A/D, rows 2 to the last row. A run inserts the new month as row 2 and extends every chart range by one row.
Values are hard-coded, no formulas. Index is the NOT seasonally adjusted (NSA) CPI-U index, 1982-84=100 (the Aug-26 values in all sheets but Education equal BLS's Aug-26 NSA figures).
New rows get the first of the month as the date.
Calculation (new rows)

MoM = index / previous-month index - 1. YoY = index / same-month-last-year index - 1. Both from the source's own index values, not from the workbook's older rows (see existing problems).

Monthly inputs: values.csv

One line for each of the ten sheets: sheet,month,prev_index,index,year_ago_index,bls_12m_pct (example: cpi_values_2026-09.csv). The three index values per series are the three index columns of the BLS CPI news release: same month last year, previous month, new month.

Source rows (BLS release, September 2026 edition checked): Table 1 has All items, All items less food and energy, Energy, Apparel. Table 3 (special aggregates) has Food and beverages, Housing, Transportation, Medical care, Recreation. Table 1 lists the published 12-month % for the first group; put it in bls_12m_pct as a cross-check, otherwise leave blank.
Sheet to BLS row: Food_Beverages = "Food and beverages"; Housing = "Housing" (not Shelter); Transportation = "Transportation" (not Transportation services); Medical_Care = "Medical care".
Release timing and page URLs are unverified for future months; confirm on BLS's schedule. In one test session only a BLS archive page found through search could be opened, the "current release" address could not. If the page cannot be opened, the owner pastes Tables 1 and 3.
Run

python cpi_update_workbook.py US_CPI.xlsx values.csv --in-place A dated backup goes to backups/US_CPI_before_<YYYY-MM>.xlsx first. --out FILE leaves the original.

Validation (STOP = nothing written)

All sheets end on the same month and the month in the CSV is the next one; prev_index equals the workbook's newest index within 0.0005; the computed YoY rounds to the published 12-month % when given. --accept-mismatch SHEET lets one sheet's prev_index differ; only use it when the owner says so.

State after the owner's Nov 6 corrections
Done by the owner: Oct-25 (x, not reported) and Nov-25 rows added; index values corrected; Education_Communication replaced by the Education series.
Done in the restatement: all dates set to the 1st of the month; MoM recalculated from the Index wherever the previous month is in the sheet; Sep-26 YoY recalculated against the Sep-25 row; Oct-25 index cell emptied (the "x" is kept as a note in column E) so charts show a gap instead of zero; chart caches rebuilt; the Food_Beverages YoY chart now plots Food_Beverages.
Missing months: Nov-25 has no MoM (Oct-25 is x) and no YoY (no Nov-24 row). Sep-25 MoM/YoY are the old values (no Aug-25 or Sep-24 row to check them).
OPEN: YoY for Oct-25 to Aug-26 cannot be recalculated without 2024-25 index values. Dec-25 to Mar-26 agree with BLS-type figures; Apr-26 to Aug-26 do not (CPI_All Jul-26 shows 4.80%, BLS reported 3.4% for July 2026). Do not change them without the year-ago index values from the owner.
Chart ranges now cover rows 2 to 14; each run extends them by one row.
The file has a dead link to D:\ITPM Work Done Part 2\US_CPI.xlsx; Excel may ask about it.