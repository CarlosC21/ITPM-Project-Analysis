CFTC Commitments of Traders: Metals & Energy (conventions)
1. What it is

Weekly positioning of money managers in 11 futures contracts (CFTC Disaggregated report, Futures Only). The workbook holds, per week and contract: open interest, money-manager longs, money-manager shorts. Net long/short as % of open interest = (longs - shorts) / open interest is calculated by formula.

2. Source
CFTC "Historical Compressed" page: https://www.cftc.gov/MarketReports/CommitmentsofTraders/HistoricalCompressed/index.htm Disaggregated Futures Only, latest year, Excel (f_year.xls, about 39 MB, 12,000 rows x 188 columns).
The owner downloads it and attaches it each week (automatic download not built; not tested whether the workspace can reach cftc.gov).
Verified Nov 9, 2026: the file read for 2026 ran to report date 2026-11-03; every week from 2026-01-06 to 2026-10-20 equals the workbook for all 11 contracts.
Not verified: when the year file is refreshed (CFTC normally releases Fridays about 3:30 p.m. ET; the owner's file, saved Nov 9, already held Tuesday Nov 3 data).
3. Contracts (matched by CFTC_Contract_Market_Code, never by name)
Sheet	Code	Name must contain
Crude Oil WTI	067411	CRUDE OIL, LIGHT SWEET-WTI
Crude Oil Brent	06765T	BRENT LAST DAY
Heating Oil	022651	NY HARBOR ULSD
Palladium	075651	PALLADIUM
Platinum	076651	PLATINUM
Silver	084691	SILVER
Gold	088691	GOLD
Copper	085692	COPPER
Aluminium	191693	ALUMINUM
Steel	192651	STEEL-HRC
Natural Gas	023651	NAT GAS NYME
4. Source columns kept (7)

Market_and_Exchange_Names, Report_Date_as_MM_DD_YYYY, CFTC_Contract_Market_Code, Open_Interest_All, M_Money_Positions_Long_ALL, M_Money_Positions_Short_ALL, Open_Interest_Old. Workbook use: OI = Open_Interest_All, Longs = M_Money_Positions_Long_ALL, Shorts = M_Money_Positions_Short_ALL. Open_Interest_Old is not stored in the workbook (it has no column). The script writes it to the extract csv and uses it as a check only (must not exceed Open_Interest_All). In the 2026 file it equals Open_Interest_All for all 484 rows.

5. Release schedule

Data are as of Tuesday; released Friday 3:30 p.m. ET (unverified here). Holiday weeks shift the report date by a day (workbook gaps are 6, 7 or 8 days).

6. Workbook layout (Commitment_of_Traders_Metals_&_Energy.xlsx)
Data - Energy & Metals: row 1 commodity names (B1, E1, H1 ... AF1), row 2 headers, A = date, then 3 columns per commodity in this order: WTI B:D, Brent E:G, Heating Oil H:J, Palladium K:M, Platinum N:P, Silver Q:S, Gold T:V, Copper W:Y, Aluminium Z:AB, Steel AC:AE, Natural Gas AF:AH. Typed values, newest week at the bottom (ascending), history kept (from 2006-06-13). Column AM holds the commodity names (lookup list), untouched.
11 commodity sheets (Crude Oil Brent, WTI, Heating Oil, Palladium, Platinum, Silver, Gold, Copper, Aluminium, Steel, Natural Gas): A3 downwards = formulas linking row for row to the data sheet (A date, B OI, C longs, D shorts, E net %). Each has two charts.
Energy & Metals Summary: latest 25 weeks (LARGE on the data sheet dates), values through INDIRECT/VLOOKUP. Its commodity order differs from the data sheet (WTI, Brent, Heating Oil, Natural Gas, Palladium, ...); the script reads it from row 3.
NOTES: the owner's manual instructions.
7. What the script does each week

python cot_update_workbook.py <workbook> <f_year.xls> [<previous_year.xls>] --out NEW.xlsx (or --in-place, which first writes backups/<name>_before_<first new date>.xlsx).

Reads the 11 contracts and 7 columns from the file.
Checks every week present in both files (all weeks of the year so far, 11 x 3 values) is identical.
Appends each week after the workbook's last date: one row to the data sheet, one row (A:E, same formulas) to each of the 11 commodity sheets. Styles are copied from the last row. Row N of every sheet is the same date.
Refreshes cached Summary values, removes calcChain and sets full recalculation on open (Excel recalculates when opened).
Edits the file at XML level, so charts, conditional formats and formatting are kept. Only the data sheet, the 11 commodity sheets, the 22 chart ranges, the Summary cache, workbook.xml/rels/content-types (calcChain) change.
In January the new year's file does not contain the last December week: pass last year's file as a second argument.
8. STOP checks (nothing is written, exit code 2)
A required column is missing, or FutOnly_or_Combined is not FutOnly.
A contract code is missing, or its name does not contain the expected text.
The 11 contracts do not share the same set of report dates; a duplicate date with different values.
Non-numeric or non-integer values; open interest <= 0; longs or shorts > open interest; Open_Interest_Old > Open_Interest_All.
Any week in both files differs from the workbook (history is never changed silently).
The workbook's last week is not in the source (continuity cannot be checked).
A gap of fewer than 5 or more than 9 days between consecutive new weeks.
Workbook layout differs (commodity names in row 1, Summary headers, row already filled). Warning only: open interest moves more than 30% in a week (likely a contract roll). If the workbook is already up to date the script says so and writes nothing.
9. Known problems in the existing workbook
Chart ranges were uneven (7 charts stopped at rows 1035 to 1049, others ran 30 to 120 rows past the data). Fixed Nov 9, 2026: the script now sets all 22 charts (series A3:A and E3:E) to end at the last data row on every run.
The data sheet autofilter covers A2:AH770 only (old range).
Palladium, Platinum, Silver and Gold sheets have empty styled rows below the data (to rows 1087, 1106, 1127, 1112); the script fills rows from 1066 and checks they are empty first.
Early history has blanks (for example Brent and WTI start later than 2006); not an error.
NOTES says "Sources (as of Mar-2021)".
10. Regression test

cot_values_2026-11-03.csv (extract of 2026-10-20, 10-27 and 11-03). Running the script on the Oct 20 workbook with the Nov 3 file must add 2 weeks, 22 contract-weeks; LibreOffice recalculation of the 22 new rows and the 25-row Summary matched the script's cached values (0 differences in 165 commodity cells and 575 Summary cells). Eleven failure cases tested (changed history value, missing contract, missing week, contract short of a week, renamed contract, long > OI, Combined file, missing column, text in a number, source older than the workbook, already up to date).

11. Compare script

Not built. If wanted: compare the owner's manual workbook with the automated one on the data sheet A3:AH and the 11 commodity sheets A:E, cell for cell.

12. Open items
Master updated Nov 9, 2026 (weeks 2026-10-27 and 2026-11-03). Owner to open it in Excel (checked in LibreOffice only).
Decide whether Open_Interest_Old should get a column in the workbook.
Automatic download of f_year.xls not built and not tested; the owner attaches it each week.