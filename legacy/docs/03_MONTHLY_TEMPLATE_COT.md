Weekly run: CFTC Commitments of Traders, Metals & Energy

Paste this into a new chat (in the ITPM Project) and attach the file.

Update Commitment_of_Traders_Metals_&_Energy with the attached CFTC file (Disaggregated, Futures Only, latest year). In January also attach last year's file.

Steps for Claude:

Read cot_conventions.md and the whole of cot_update_workbook.py in the project first.
Stage the master Commitment_of_Traders_Metals_&_Energy.xlsx and cot_update_workbook.py from my folder.
Test: python cot_update_workbook.py <master copy> <attached file> --out test.xlsx --extract-csv cot_values_<last date>.csv. If it prints STOP or fails, tell me exactly what it said and write NOTHING to my folder.
Show me the new weeks (the printed table) and any WARNING lines. Do not pass anything to get around a STOP.
If fine: run with --in-place on the copy (it makes backups/<name>_before_<first new date>.xlsx), write that backup into a backups folder inside my project folder, confirm it exists, and only then replace the master. The workbook must be closed in Excel to be replaced.
Tell me: weeks added, anything odd (warnings, chart ranges), in a few lines.