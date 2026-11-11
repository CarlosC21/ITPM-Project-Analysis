Monthly template: US Employment Situation. Copy the text between the lines into a new chat in the project

You attach nothing and paste nothing. The workbook is read from your connected project folder and the figures are read from the BLS release. Requirements: the Claude desktop app open and linked to your computer. If the chat says it cannot reach the folder, attach the workbook to that chat and send again. If it cannot open the BLS release, it will ask you to paste Table B-1, Table A-1 and the headline/revisions paragraph.

--- copy from here ---

Run the monthly US Employment Situation update. Use the master workbook US_Employment_Report.xlsx in my connected project folder and the latest BLS Employment Situation release (find it by searching bls.gov, then fetch the result; do not take figures from FRED or any other site). Follow the project instructions and employment_conventions.md. Show me the values before running. Do not use --accept-previous-mismatch.

--- copy to here ---

Before you run it (about 2 minutes):

Is the new report out? Releases are 8:30 a.m. ET on a Friday. Next: Fri Dec 4, 2026 (November data). Check the BLS schedule each month.
Make sure the workbook in the folder is the version this process last wrote (ends Oct-26 before the November run). If you have edited it by hand since, tell the chat what you changed. Hand-entering revised Aug/Sep-type values is what caused the one-time stop in October.
Keep the Claude desktop app open and the computer awake and online.

What you should get back: the month added, the 11 new MoM/YoY values, the revisions to the two prior months (usually "none" or small), warnings, the anomaly count, and the backup file name (backups/US_Employment_Report_before_<YYYY-MM>.xlsx).

If it says STOP: nothing was written. Send the message to the chat. Typical causes: a figure mistyped in values.csv (headline or identity checks), the workbook not holding the last release, or (in February-ish) the annual benchmark changing history. Do not ask it to skip the check unless you understand why.

After the run:

Add a line to the parallel-run log in 01_INDICATOR_INDEX.md (the chat does this if you ask).
When FRED shows the month (it lags the release), compare your manual workbook with the automated one. The trust criterion is 3 months in a row matching cell for cell, every difference explained.
The backups folder: you can delete old backups once you are happy; each run creates a new one named for its month.
Your manual workbook is never touched.

Notes:

The year-ago Unemployment_Rate for Oct-26 is blank by design (no October 2025 household survey).
Annual benchmark (usually with the January report in February): expect a year-ago STOP; ask the chat to plan a history refresh.