ISM Manufacturing industry lists: conventions

Read this before running anything for this indicator. It records how the existing workbook works, so the scripts never have to be re-derived from a chat.

Files for this indicator
ism_industry_parser.py: turns report text into ranked rows and validates them.
ism_update_workbook.py: adds the next month to the workbook using those rows.
ism_compare_workbooks.py: optional; lists every difference between two workbooks.
ism_manufacturing_test_oct_2026.txt: October 2026 text with all seven headings, as a sample. Python 3 with openpyxl is required. The same three scripts also handle the Services workbook (see ism_services_conventions.md); the workbook type is detected from its block titles.
Running the updater

python ism_update_workbook.py workbook.xlsx report.txt [--in-place | --out new.xlsx] [--window N]

No month is typed. It is the month after the workbook's newest column, and the report text must name that month or the run stops.
Default: APPEND. The new month becomes a new pair of columns to the right of the newest one, and all history is kept. Cell formatting and conditional formatting are extended to the new columns. With --window N the workbook keeps only the newest N months instead (the oldest drops off and no columns are added). Only use a window if the owner asks.
--in-place: copies the original to backups/<name>_before_<YYYY-MM>.xlsx next to it, then replaces the workbook. The backup is made before anything is replaced. A run that stops makes no backup and changes nothing.
--out FILE: writes a separate file and leaves the original untouched.
Where the workbook lives

ISM_Manufacturing_Index.xlsx in the owner's connected project folder. The monthly run copies it into the workspace, runs the updater with --in-place there, then writes the backup and the updated master back to the folder (backup first). Reading and writing through the connected folder has not been tested yet.

Workbook layout (verified Nov 2026)
One sheet. Seven blocks. Block titles are in column A: ISM Manufacturing (row 2), NEW ORDERS (24), PRODUCTION (46), EMPLOYMENT (68), DELIVERIES (90), INVENTORIES (112), CUSTOMER INVENTORIES (134).
The header row holds month dates (mmm-yy) in columns B, D, F, H ... one per month.
Industry rows start 3 rows below the header, exactly 18 per block.
Each month is a pair of columns: label, then score. B:C, D:E, F:G, H:I ... Newest is on the right.
Labels by block:
Manufacturing, New Orders, Production, Employment: Growth / Neutral / Contraction
Deliveries: Slower / Neutral / Faster
Inventories: Higher / Neutral / Lower
Customer Inventories: Too High / Neutral / Too Low
Rows are sorted by the newest month's score, highest first; ties keep their previous order.
Months must be consecutive; a gap or out-of-order header stops the run.
Observation (not changed): the conditional-formatting rules on the label columns look for the words Growth and Contraction in every block, including Customer Inventories, so its Too High / Too Low labels are not colour-coded. New columns copy the existing rules as they are.
Scoring rule (verified against the real September 2026 report)

The first-listed industry in each ISM list gets the largest magnitude.

Growth side (Growth, Slower, Higher, Too High): +(n - position + 1), so first = +n, last = +1.
Contraction side (Contraction, Faster, Lower, Too Low): -(m - position + 1), so first = -m, last = -1.
Neutral (not named in either list): 0. Example: 12 growth industries means the first listed scores +12 and the last +1. Two contraction industries means the first listed scores -2 and the second -1.
Report text the parser expects

One heading line per index, followed by that index's paragraph: ISM Manufacturing -, New Orders -, Production -, Employment -, Supplier Deliveries - (or Deliveries -), Inventories -, Customer Inventories -. Any lines before the first heading are ignored. Each index words its sentences differently, and the parser handles:

Counts as digits or words ("12", "eight"), with or without "Of the 18 manufacturing industries,".
Growth/contraction, decline, decrease, slower/faster, higher/lower, too high/too low.
Lists after a colon, split on semicolons (industry names contain commas), or inside parentheses when only two are named.
A "no industry..." sentence (count 0) and a "...no change" sentence (count only). Neutral industries are inferred as the 18 minus those named. If the report states a "no change" count, it is cross-checked against the inferred count.
Validation (the run stops, and writes nothing, if any fail)
The count stated in each sentence matches the number of names parsed.
Every name is one of the 18 canonical industries.
No industry appears in two lists within the same index.
The month named in the text matches the month being added.
The workbook's months are consecutive, all blocks end on the same month, and each block has exactly the 18 canonical industries.
A placeholder or unrecognized text produces "no industry sentences recognized".
Comparing manual and automated workbooks

python ism_compare_workbooks.py manual.xlsx automated.xlsx [--csv diffs.csv] Cells are matched by block, industry and month (from the header date), not by column position. Labels are compared ignoring case. It reports months present in only one file, every differing cell, and exits with an error code if there are differences. A difference is not automatically a bug in the automation; the manual file has had errors too (see below). Check each one against the source report.

Known problems in the Nov 2026 data (reported, not fixed)
44 cells held broken =IFERROR(VLOOKUP(...#REF!...)) formulas that always returned the fallback ("Neutral" / 0). The updater converts them to those values, because re-sorting rows would break their same-row references.
Customer Inventories, Jul 2026: labels contradict scores (for example "Too High" with negative scores, "Neutral" with non-zero scores).
Aug 2026: Apparel (Manufacturing) and Computer (Employment) are Growth with a score of 0; Deliveries Paper Products is Neutral with a score of 4.
Sep 2026: Deliveries Textile Mills is Slower with a score of 0. The updater prints this audit on every run. It never silently changes history.
Optional reference: expected counts for the sample

Running the updater on the original Jul-Sep workbook with ism_manufacturing_test_oct_2026.txt should give these new-month counts: Manufacturing 12 growth / 3 neutral / 3 contraction; New Orders 11 / 3 / 4; Production 11 / 4 / 3; Employment 8 / 5 / 5; Deliveries 12 slower / 4 neutral / 2 faster; Inventories 7 higher / 5 neutral / 6 lower; Customer Inventories 0 too high / 7 neutral / 11 too low.

Not automated yet
Getting the report text each month (see 02_PROJECT_CONTEXT.md).
Headline index values (this workbook does not contain them).