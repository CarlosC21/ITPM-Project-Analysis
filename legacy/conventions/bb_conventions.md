Bull & Bear Markets: S&P 500 (conventions)
1. What it is

Bull_Bear_Markets.xlsx tracks the S&P 500 against a bear-market line (20% below the rolling all-time high). Daily Open/High/Low/Close, a rolling index high, the bear-market level, and two chart series (Bull = close above the level, Bear = close at or below it). One chart on the S&P 500 sheet. The Definitions sheet lists 13 other indices (Nasdaq 100 ... ASX 200) but the workbook has no sheets for them (open question 1).

2. Source
Yahoo Finance chart data for ^GSPC, read in the BUILT-IN BROWSER (Claude asks for site access, scope "once"): https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?period1=<unix>&period2=<unix>&interval=1d then JSON.parse(document.body.innerText) in the page and print date,open,high,low,close. WebFetch, curl and the shell are refused (not on the allowlist).
Verified Nov 9, 2026: Oct 14, 15, 16 and 19 2026 (open, high, low, close) equal the workbook to the last digit (the workbook's float32-style numbers such as 7690.72998046875 are Yahoo's). So the workbook was built from Yahoo.
Not verified: Yahoo terms of use for automated retrieval; whether Yahoo restates old days; the pre-1950 history (flat O=H=L=C in the workbook).
Use COMPLETED days only. Today's bar is live until the 4:00 p.m. ET close (the Nov 9 bar was still moving at 3:53 p.m.).
3. Release schedule

Every US trading day, after the 4:00 p.m. ET close. Yahoo's daily bar is final shortly after.

4. Workbook layout (S&P 500 sheet)
Row 1 headers: A S&P 500 (date), B Open, C High, D Low, E Close, F Rolling Index High, G Rolling Bear Market Level, H Bull Market Index, I Bear Market Index, J Labels.
Data from row 2 (newest, on top) down to 1928-01-03. New days are inserted at row 2; every older row moves down.
F = MAX(C<row>:$C$<anchor>): highest High from that day back to the anchor row (1931-09-03, row 23954 on Nov 9, 2026; it moves down with each insert). Rows older than the anchor use MAX(C$<anchor>:$C<row>).
G = F*0.8. H = IF(E>G,E,#N/A). I = IF(E<=G,E,#N/A). J = =E<row> or =C<row> on 6 hand-picked label days.
Block M:O (from row 360): COVID-19 and Global Financial Crisis breach dates, typed by the owner; it moves with the rows.
The chart (J2:Q25) is fixed in place. Axis fixed from 2020-01-03 to 2026-11-11.
Window: growing; history kept.
5. What the script does (bb_update_workbook.py)

python bb_update_workbook.py <workbook> values.csv --out NEW.xlsx (or --in-place, which first writes backups/Bull_Bear_Markets_before_<first new date>.xlsx, checks it equals the original, then replaces the master).

Reads values.csv (date,open,high,low,close).
Checks the workbook's newest date, and every date in both, are identical to the last digit (history never changed silently).
Inserts one row per new day at row 2 with the same styles; F, G, H, I as formulas with cached values computed.
Moves every older row down; shifts all formulas (shared formulas too), the M:O block, the dimension and the rolling-high anchor reference.
Extends the 3 chart series ends by the number of new rows (starts kept; --align-chart sets all starts to row 2).
Removes calcChain and sets full recalculation on open. Edited at XML level; the chart and formats are kept.
Proof step re-reads the output: old rows equal and moved, formulas equal the shifted originals, new F/G/H/I recomputed, only the sheet, chart, workbook.xml, rels and content types changed.
6. STOP checks (nothing written, exit code 2)

Wrong sheets, headers, F2/G2 formula, chart count or range pattern; csv header; bad date or number; duplicate or weekend or future date; low/high not bracketing open/close; newest workbook date missing from the csv; any overlap day (newest 10 rows) differing from the workbook; a gap over 8 days between consecutive new days; a close-to-close move over 12%. Warning only: a gap of 5 to 8 days. "Nothing to add" if the csv has no day newer than the workbook.

7. Known problems in the existing workbook (reported, not fixed)
(Fixed Nov 9, 2026 at the owner's request: 20 duplicate trading days, 2026-06-08 to 2026-07-07, rows 94-113 removed.)
The chart series start at different rows (Close E2, Bear I55, Labels J356) against dates from A2.
Rows 222 and 223 (2026): low above open/close.
Rows below the anchor (1928-01-03 to 1931-09-02): the rolling high looks forward to 1931-09-03 instead of back.
Chart series ranges run 474 rows past the data.
The workbook must be closed in Excel before the master is replaced (a lock file ~$Bull_Bear_Markets.xlsx shows it is open).
8. Regression test

bb_values_2026-11-06.csv (Oct 14 to Nov 6 2026). Run on the Oct 19 workbook: adds 14 days (Oct 20 to Nov 6), proof passes; LibreOffice recalculation of F:I matched the cached values (336 cells checked, 0 differences). Eleven failure cases tested (changed overlap value, missing newest date, bad header, text in a price, high below open, 32% move, weekend date, duplicate date, 14-day hole, future date, already up to date); in-place on a copy writes a backup identical to the original.

9. Compare script

Not built. If wanted: match the owner's manual workbook with the automated one on A:E for every row.

10. Open items

Owner's answers, Nov 9, 2026:

Scope: S&P 500 only (the 13 other names on Definitions have no sheets and stay out of scope).
The 20 duplicate days were removed Nov 9, 2026 with bb_repair_duplicates.py (one-off; deleted the lower copy, rows 94-113).
Chart series starts (E2, I55, J356): not decided; default keeps them, --align-chart sets all to row 2.
DONE Nov 9, 2026: duplicates removed and 14 days added to the master (see 01_INDICATOR_INDEX.md log). Excel and say so;
11. First run (Nov 9, 2026)

Backups: backups/Bull_Bear_Markets_before_repair_2026-11-09.xlsx (identical to the original) and backups/Bull_Bear_Markets_before_2026-10-20.xlsx (after the repair). Master now ends 2026-11-06: 24,872 dated rows, no duplicates, F:I recalculated in LibreOffice (8,288 cells, 0 differences). Next run: 2026-11-09 and later days (read after the 4:00 p.m. ET close). Not yet opened in Excel by the owner.