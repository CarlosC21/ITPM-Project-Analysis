NFIB Small Business Optimism: components (built Nov 9, 2026)
1. What it is

NFIB Small Business Optimism - COMPONENTS.xlsx (connected project folder). 11 sheets: Total + 10 components. Per sheet: A date (26th of the month, newest at row 3), B value (1 decimal), C 3-month average (IF(B{r+2}="",#N/A,AVERAGE(B{r}:B{r+2}))), D YoY change B{r}-B{r+12}, E 3mma YoY C{r}-C{r+12}. Blocks at G:S and U:AG are empty headers; Total also has a summary block H83:K86 (latest, percentile rank, mean). Older rows (8-26) take A and B from an external workbook (NFIB ... COMPONENTS.xlsm, sheet Data) by formula; the newest rows are typed. 66 charts (6 per sheet) on ranges like $A$3:$A$3005. openpyxl DROPS CHARTS: the script edits the XML. History is kept (growing block, one row added per month; rows below the block must be empty).

2. Source (verified Nov 9, 2026)
Components (10): https://www.nfib-sbet.org/Indicators.html, Monthly, no filters, Seasonal Adjustment ticked (the default). Matches the workbook to 1 decimal. The page is JavaScript: WebFetch sees nothing. Use the BUILT-IN BROWSER (request_access for nfib-sbet.org first). Per indicator: set the selector, click SHOW RESULTS (10-20 s), read $('#indicatorGrid').data('kendoGrid').dataSource.data().toJSON().slice(-6) (field indexvalue, monthyear). Selecting an indicator re-ticks Seasonal Adjustment. A loop that drives the page's own controls worked; a direct POST to api.nfib-sbet.org from page script was blocked by the permission classifier: do not retry it.
Total: NOT the site's Optimism Index (site Aug-26 98.19 SA / 97.89 unadjusted vs workbook 98.8). Owner confirmed: headline index from the NFIB monthly report PDF (search "NFIB Small Business Optimism Index <Month Year>", fetch the nfib.com NFIB-SBET-Report-<Month>-<Year>.pdf result). Sep-26 = 97.9. The report's component table is whole numbers: do not use it.
Codes: expand_good=Expansion Plans; emp_count_change_expect=Employment Plans; inventory_expect=Inventory Plans; bus_cond_expect=Economic Expectations; sales_expect=Real Sales Expectations; cap_ex_expect=Capital Plans; inventory_current=Current Inventory; job_opening_unfilled=Current Job Openings; credit_access_expect=Credit Expectations; earn_change=Earnings Trend.
3. Release timing

Monthly report, 2nd Tuesday-ish (Sep data Oct 13-14, 2026). October data due about Nov 10, 2026 (unverified).

4. values.csv (sheet,month,value)

11 rows for the NEW month (raw site numbers; the script rounds half-up to 1 decimal) plus, optionally, 11 rows for the PRIOR month. Prior-month rows are only compared with the workbook (differences are warnings; seasonal factors can be revised). See nfib_values_2026-09.csv.

5. Script: nfib_update_workbook.py (stdlib only)

python nfib_update_workbook.py "<workbook>" values.csv --out test.xlsx then --in-place (writes backups/<name>before<YYYY-MM>.xlsx). --correct "Sheet,YYYY-MM,value" fixes one typed level (logged); never used unless the owner says so. STOPs: sheet names/order, 66 charts, months not consecutive, sheets disagree, rows below the block not empty, csv missing/extra/ wrong month, value out of range (Total 70-130, others -60..80) or a move > 6 (Total) / 20 (others) from last month. Also: rewrites the 3 formula columns with computed cached values, shifts only A:E down one row (other columns untouched), keeps external-link formulas (A unchanged, B's own-row $A reference moved), refreshes chart caches and the Total summary cache, deletes calcChain and sets fullCalcOnLoad so Excel recalculates when opened. Proves old months are unchanged.

6. Test results (Nov 9, 2026)

Engine reproduced every stored C/D/E value in the original (only difference: the text cell). Output recalculated by LibreOffice from stripped caches: 1,144 cells, 0 mismatches. Total summary block matched except K83 (LibreOffice lacks FILTER). Only A:E and the Total summary cells changed; 66 charts kept. Not opened in Excel from here: open it once and say if anything looks off.

7. Known problems
Aug-26 Economic Expectations was the text "9..3": corrected to 9.3 on owner's instruction (Sep-26 run). That changed Aug-26 3mma 13.30 -> 11.97, YoY #VALUE! -> -23.89, 3mma YoY -16.49 -> -17.83.
Total Aug-26: workbook 98.8, Sep report says August was 98.7 (kept 98.8).
Excel will ask to update links to the external .xlsm (as before).
8. Open items

Parallel run: compare against the owner's manual Sep-26 update (compare script not built for this workbook). Oct-26 data (about Nov 10): first run on release timing.