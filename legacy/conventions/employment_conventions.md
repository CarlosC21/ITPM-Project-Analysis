US Employment Situation workbook: conventions

Script: employment_update_workbook.py (standard library only; edits the file XML so the 22 charts survive). A copy of the script is in the owner's connected project folder (added Nov 6, 2026 with the owner's permission; the only extra file allowed there). The project doc claude/employment_update_workbook.py is the master copy; if the two differ, tell the owner. Master workbook: US_Employment_Report.xlsx in the owner's connected project folder (folder read/write TESTED Nov 6, 2026). Sample input: employment_values_2026-10.csv (real October 2026 values, checked against the BLS release Nov 6, 2026).

Status (Nov 6, 2026)
October 2026 is in the workbook (newest row on top of all 11 sheets). Workbook now ends Oct-26.
The one-time --restate is DONE (dates are the 1st; Unemployment_Rate has Oct-25 and Nov-25 rows). Do not run it again.
October was run with --accept-previous-mismatch (see Flag below) because the copy given to the run already held the revised Aug/Sep totals. That was a one-off. NOVEMBER AND LATER RUN WITHOUT THE FLAG. If the "holds the previous release" check stops next month, find out why before using the flag.
Backup of the pre-October workbook: backups/US_Employment_Report_before_2026-10.xlsx (the owner may delete the backups folder after validating; the script recreates it and names each backup by the month being added).
Layout (verified Nov 6, 2026)
11 sheets, same layout: row 1 headers (B Index, C MoM, D YoY); data from row 2, NEWEST FIRST; A = month (mmm-yy, first of the month). Two line charts per sheet (level from column B, YoY from column D), rows 2 to the last row.
Values are hard-coded, no formulas. All series are seasonally adjusted. History starts Sep-24.
Sheet -> FRED id -> BLS Table B-1 row (levels in thousands of persons): Unemployment_Rate = UNRATE (Table A-1, percent) | Total_Nonfarm_Payrolls = PAYEMS "Total nonfarm" | Total_Government = USGOVT "Government" | Total_Private = USPRIV "Total private" | Goods_Producing = USGOOD "Goods-producing" | Construction = USCONS | Manufacturing = MANEMP | Trade_Transport_Utilities = USTPU "Trade, transportation, and utilities" | Financials = USFIRE "Financial activities" | Professional_Business_Services = USPBS | Service_Producing = SRVPRD, NOT a B-1 row: the script derives it as Total nonfarm - Goods-producing (equals Private service-providing + Government; checked).
Unemployment_Rate: Oct-25 is blank (never published: no household survey), note in column E. Its YoY and the Oct-26 YoY stay blank by design (Oct-26 YoY will stay blank; Oct-27 YoY will also be blank because the year-ago month is blank). Nov-25 now holds a rate in the owner's workbook, so Nov-26 YoY will be calculated.
Calculation (new rows)
Level sheets: MoM = level / previous month - 1; YoY = level / same month last year - 1.
Unemployment_Rate: percentage-point difference shown as a percent (4.1 vs 4.2 = -0.10%), the rule the sheet followed before May-26. YoY stays blank while the year-ago month is blank.
Monthly source and flow
Read the BLS Employment Situation release for the new month. HOW TO OPEN IT (learned Nov 6, 2026): a direct WebFetch of a bls.gov URL typed from memory is refused ("not on the allowlist"). First run WebSearch with allowed_domains=["bls.gov"], query like "The Employment Situation <Month> <Year> BLS"; the results include https://www.bls.gov/news.release/pdf/empsit.pdf (latest release) and .../empsit.nr0.htm. WebFetch that result link: offset 0 gives the headline, the revisions paragraph ("from X to Y" for the two prior months) and Table A-1; offset 100000 gives Table B-1 (levels, seasonally adjusted: year-ago month, the two prior months, new month). The latest-release URL only shows the latest month; archived releases are under /news.release/archives/. Do not use FRED or other sites for the figures. If the release cannot be opened, ask the owner to paste Table B-1, Table A-1 and the headline/revisions paragraph.
Write values.csv (format below). Show the owner the rows (a table of the four months by series is fine) before running.
Stage the master workbook AND employment_update_workbook.py from the owner's folder with device_stage_files (if the script is not there, write it out from the project doc). Re-stage and compare mtime if hours have passed.
Test first: python3 -I employment_update_workbook.py US_Employment_Report.xlsx values.csv --out updated.xlsx (--in-place would write into the workspace copy, not the owner's folder; use --out for cloud sessions).
Check updated.xlsx (newest row, 22 charts, history unchanged). Then, in this order: a) copy the ORIGINAL staged workbook to backups/US_Employment_Report_before_<YYYY-MM>.xlsx in the owner's folder (device_bash mkdir -p the backups folder, then device_commit_files); verify with sha256sum that it equals the original; b) only then commit the updated file over the master with expectedMtimeMs = the staged mtime; verify the checksum of the master equals updated.xlsx. If the backup cannot be written: stop, do not replace the master.
FRED is only a later cross-check (on release day it still showed the previous month). data.bls.gov, download.bls.gov, api.bls.gov and fred.stlouisfed.org are blocked from the shell (403).
values.csv

key,month,value, months as YYYY-MM. Keys: UR, TOTAL, PRIVATE, GOODS, CONSTRUCTION, MANUFACTURING, PSP (private service-providing), TTU, FINANCIAL, PBS, GOVERNMENT. For each key: the new month and the two prior months (revised values). Year-ago month for every key except UR. Then: HEADLINE_CHANGE,<new month>,<n>; for each of the two prior months ORIGINAL_CHANGE,<m>,<n> and REVISED_CHANGE,<m>,<n> (total nonfarm change in thousands as the release text states). Example for the November 2026 run (Dec 4): new month 2026-11, prior months 2026-10 and 2026-09, year-ago 2025-11; ORIGINAL_CHANGE 2026-10 should be 165 (what the workbook holds) and ORIGINAL_CHANGE 2026-09 is 60.

Validation (STOP = nothing written)
Workbook: dates are the 1st, months consecutive, all 11 sheets end on the same month; the new month is the next one.
values.csv complete, nothing extra, numbers valid.
Identities for every month supplied: Total = Private + Government; Private = Goods + Private service-providing; Total - Goods = Private service-providing + Government (tolerance 1).
Components outside the identities: Goods - Construction - Manufacturing stays within 15 of the previous month; Private service-providing - TTU - Financial - PBS within 250; monthly % move below 1%; revision to a prior level below 300.
Total change: new - previous = HEADLINE_CHANGE; revised previous - prior = REVISED_CHANGE; the workbook's own change for the previous month = ORIGINAL_CHANGE (the workbook holds the last release).
Year-ago level in the release equals the workbook's (STOP means history needs a refresh, e.g. after the annual benchmark).
Unemployment rate for the two prior months equals the workbook (it is revised only once a year).
After building: history unchanged except the two prior months, 22 charts present, XML well formed.
Flag: --accept-previous-mismatch (off by default; one run only)

Lets the run continue when the workbook's previous-month change equals the REVISED figure instead of the originally published one (the workbook was updated by hand with revised numbers). It still stops if the workbook's change equals neither. It adds a warning to the output. Use only when the owner says so, for that one month; never as a routine. Tested Nov 6, 2026: without the flag the October run stopped, with it the run passed.

Revisions

BLS revises the two prior months each release. The script overwrites those two months' levels (all level sheets), recomputes their MoM/YoY, and prints old -> new. Nothing older is touched. If the older month the release implies differs from the workbook, that is a warning, not a change.

Known problems (not changed)
Jul-26: the workbook's total (158,858) is not the value the Oct 6 release implies (158,882, the Oct 2 edition value). The script warned on the October run (it compares Aug with Jul). It should not warn in the November run.
Unemployment_Rate Nov-25: the anomaly audit reports YoY empty although the levels give 0.30%. Owner said (Nov 6, 2026) to leave it empty; the audit will keep listing it as 1 anomaly, which is expected.
Annual benchmark and seasonal-factor revisions (published with the January report; confirm on the BLS schedule) restate history. The year-ago check will STOP; a refresh mode is not built.
Dead link to D:\ITPM Work Done Part 2\US_Employment_Report.xlsx (no formula uses it); Excel may ask about it.
Next release: November data on Fri Dec 4, 2026 (BLS CES release schedule page, read Nov 6, 2026).
Not tested in Excel itself (LibreOffice opens and renders it).