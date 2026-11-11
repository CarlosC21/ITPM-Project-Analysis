ISM Services (Non-Manufacturing) industry lists: conventions

Read this before running anything for this indicator. Scripts are the same three as Manufacturing (ism_industry_parser.py, ism_update_workbook.py, optional ism_compare_workbooks.py); the workbook type is detected automatically. Sample text: ism_services_test_oct_2026.txt. Usage, backups, --in-place, --out, --window are as in ism_manufacturing_conventions.md.

Where the workbook lives

ISM_Non-Manufacturing_Index.xlsx in the owner's connected project folder (same folder as the Manufacturing master). Folder read/write through the connected computer is not tested yet.

Workbook layout (verified Nov 2026)
One sheet, "Sectors". Six blocks, titles in column A (matched ignoring case and repeated spaces): ISM Non-Manuf (row 3), Business Activity (26), New  Orders (49), EMPLOYMENT (72), DELIVERIES (95), INVENTORIES (118).
Header row holds month dates in B, D, F, H ...; industry rows start 3 rows below the header, exactly 18 per block. Each month is a pair: label, then score. Newest on the right.
Labels: ISM Non-Manuf, Business Activity, New Orders, Employment = Growth / Neutral / Contraction (even though the report says "increase/decrease"); Deliveries = Slower / Neutral / Faster; Inventories = Increasing / Neutral / Decreasing.
Rows are sorted by the newest score, highest first; ties keep their previous order.
Conditional formatting (text colour on labels, colour scale on scores) is copied to the new columns.
Scoring rule (confirmed by the owner)

First-listed industry gets the largest magnitude. Up side (Growth, Slower, Increasing): +(n - position + 1). Down side (Contraction, Faster, Decreasing): -(m - position + 1). Neutral (not named): 0. If a list is empty ("no industry reporting faster deliveries"), that side has no rows.

Report text the parser expects

Headings, one per index: ISM Non-Manufacturing -, Business Activity -, New Orders -, Employment -, Deliveries -, Inventories -. Handled wording: counts as digits or words; growth / contraction / increase / decrease / slower / faster; lists split on semicolons; "the only industry reporting a decrease ... is X" (no colon); "There was no industry reporting faster deliveries". Neutral = the 18 Services industries minus those named. Validation stops the run (nothing is written) if a stated count differs from the names parsed, a name is not one of the 18, an industry is in two lists, the month in the text is wrong, a section is missing, or a sentence type does not belong in that index.

Hidden helper area (columns R:AN) and how many months fit
The workbook has a hidden scratch area from the old manual process: pasted sentences in column S, TEXTSPLIT formulas, and rank/label tables in S:X. The updater does not read or change it.
New months go in H:I, J:K, L:M, N:O, P:Q (Oct-26 to Feb-27). The month after that would land on the helper area, so the run STOPS before writing anything. The owner must then move or delete the helper area (it is not needed for the automated process) before the next run.
Saving through the updater removes things Excel cannot keep in this format: the hidden form button (a leftover from an .xlsm), and the "spill" behaviour of the four TEXTSPLIT cells in the helper area (they become single-cell formulas). The visible tables are unaffected. The workbook still carries a dead link to ISM Non-Manufacturing Index.xlsm on D:; Excel may ask about it.
Existing data
14 cells (Arts, Entertainment & Recreation rows, B:G in four blocks) were VLOOKUP formulas into the helper area. The updater replaces them with their current values, because re-sorting would break their same-row references.
Employment, Jul 2026: positive scores are 1-6, 6, 7 (a duplicate 6, no 8) and negative scores have a duplicate -5. Reported by the audit, not fixed.
Expected counts for the October 2026 sample (new month)

ISM Non-Manuf 15 growth / 1 neutral / 2 contraction; Business Activity 13 / 3 / 2; New Orders 15 / 2 / 1; Employment 6 / 5 / 7; Deliveries 14 slower / 4 neutral / 0 faster; Inventories 10 increasing / 4 neutral / 4 decreasing.