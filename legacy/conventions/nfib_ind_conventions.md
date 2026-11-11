NFIB Small Business Optimism: industries (built Nov 9, 2026)
1. What it is

NFIB Small Business Optimism - INDUSTRIES.xlsx (connected project folder). 10 sheets: Total + 9 industries (Construction, Manufacturing, Transportation, Wholesale, Retail, Agriculture, Finance, Services, Professions). Per sheet: A date (26th of the month for typed rows, newest at row 3), B level (1 decimal when typed), C 3-month average (IF(B{r+2}="",#N/A,AVERAGE(B{r}:B{r+2}))), D YoY change B{r}-B{r+12}, E 3mma YoY C{r}-C{r+12}. Header row 2 (B2 = sheet name, C2:E2 and G2/U2 are formulas on it). Blocks G:S and U:AG are empty headers. No summary block. Data is a growing block (Aug-24 to Sep-26 = rows 3-28 now); rows below must be empty. Rows 9-28 take A and B from an external workbook (NFIB ... INDUSTRIES.xlsm, sheet Data, by formula; older levels carry 3 decimals); rows 3-8 are typed. 60 charts (6 per sheet) on ranges like $A$3:$A$3006. openpyxl DROPS CHARTS: the script edits the XML. C:E hold Excel shared formulas; the script rewrites C:E in every data row as plain formulas with fresh cached values.

2. Source (verified Nov 9, 2026)
ONE page for all 10 sheets: https://www.nfib-sbet.org/Indicators.html, indicator "Small Business Optimism Index" (OPT_INDEX), Frequency Monthly, Seasonal Adjustment ticked (the default), no Employee Size / Region filter. Industry sheet = that one industry selected in the Industry box (codes 1 Construction, 2 Manufacturing, 3 Transportation, 4 Wholesale, 5 Retail, 6 Agriculture, 7 Finance, 8 Services, 9 Professions; 10 Other is not in the workbook). Total = NO industry selected (all industries). This is the site's own all-industry series (Sep-26 97.814, Aug 98.84, Jul 99.98), which matches the workbook Total (97.8, 98.8, 100.0). It is NOT the NFIB report headline used by the COMPONENTS workbook (Sep-26 97.9).
The page is JavaScript: WebFetch sees nothing. Use the BUILT-IN BROWSER (request_access for nfib-sbet.org first). Drive the page's own controls from page script: $('#selectIndustry').data('kendoMultiSelect').value(['1']) + .trigger('change'), click SHOW RESULTS, wait for $('#indicatorGrid').data('kendoGrid').dataSource.data() to change (fields monthyear, indexvalue; the grid object is briefly undefined while it reloads: guard with try/catch). ~10 s per industry, 30-60 s for Total. Read .toJSON().slice(-14). Values have 3 decimals; the script rounds half-up to 1 decimal. A direct POST to api.nfib-sbet.org from page script was blocked by the permission classifier in the COMPONENTS run: do not retry it.
3. Release timing

Same monthly NFIB release as COMPONENTS: about the 2nd Tuesday (Sep data Oct 13-14, 2026; Oct data about Nov 10, 2026, unverified). The page header says "LAST UPDATED 9/2026" until the new month is loaded: check it before reading.

4. values.csv (sheet,month,value)

10 rows for the NEW month (raw site numbers) plus, optionally, 10 rows for the PRIOR month, which are only compared with the workbook (difference > 0.051 is a warning; seasonal factors get revised). See nfib_ind_values_2026-09.csv.

5. Script: nfib_ind_update_workbook.py (stdlib only)

python nfib_ind_update_workbook.py "<workbook>" values.csv --out test.xlsx then --in-place (writes backups/<name>before<YYYY-MM>.xlsx, checked identical to the original before the master is replaced). --correct "Sheet,YYYY-MM,value" fixes one typed level (logged); never used unless the owner says so. The new month is worked out from the workbook (latest + 1). STOPs: sheet names/order, 60 charts, months not consecutive, sheets disagree, rows below the block not empty, csv missing/extra/wrong month/duplicate/non-numeric/unknown sheet, value outside 80-120 (Total) or 70-140 (industries), a move > 6 (Total) or > 30 (industries) from last month. Industry moves are large (Transportation Aug->Sep-26 was -21.8). Also: keeps external-link formulas (A unchanged, B's own-row $A reference moved), refreshes chart caches, deletes calcChain and sets fullCalcOnLoad so Excel recalculates on opening. Proves old A and B cells are unchanged and no shared-formula stub is left behind.

6. Test results (Nov 9, 2026)

Rewound a copy of the owner's workbook to Aug-26 (inverse of the monthly shift) and ran the script with the real Sep-26 site values: 1,350 A:E cells compared with the owner's manual Sep-26 workbook, 1 difference: Professions Sep-26 (manual 102.4, site 97.278 -> 97.3) and the three cells computed from it. Every other level, 3mma, YoY and 3mma YoY identical (the engine also reproduces all old cached values). Only A:E and the 60 charts, workbook.xml, content types and rels changed; calcChain removed. LibreOffice recalculated 780 C:E cells from stripped caches: 0 mismatches. 9 failure cases stop with nothing written; --in-place on a copy gave an identical backup. Not opened in Excel from here: open it once and say if anything looks off.

7. Known problems (reported, not fixed)
Professions Sep-26: workbook 102.4 equals the Jul-26 level; the site shows 97.278 (97.3). Likely a manual slip. Fix with --correct "Professions,2026-09,97.3" only if the owner says so (changes that month's 3mma 100.8 -> 99.1, YoY 3.01 -> -2.09, 3mma YoY -2.22 -> -3.92, and later months' 3mma/YoY that use it).
Transportation May-26 is the text "96..4" (site 96.39 -> 96.4): 3mma ignores it, YoY is #VALUE!. Fix with --correct "Transportation,2026-05,96.4" on the owner's word.
Jun-26 typed levels differ slightly from the site now (Construction 99.3 vs 99.353, Professions 102.6 vs 102.541): seasonal revisions or rounding; not compared (only the prior month is).
Excel will ask to update links to the external .xlsm.
8. Open items
The workbook already holds Sep-26 (typed by the owner), so the first real run is Oct-26 (about Nov 10).
Parallel run: compare the Oct-26 result with the owner's manual update (compare script not built for this workbook).
If the workbook is open in Excel (lock file ~$NFIB ... INDUSTRIES.xlsx seen Nov 9) it cannot be replaced: close it first.