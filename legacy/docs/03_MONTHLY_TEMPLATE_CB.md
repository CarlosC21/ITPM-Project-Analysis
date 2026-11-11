Monthly (or any-day) run: Corporate Bond Indices - US

Paste this into a new chat (project "ITPM Project") whenever you want the workbook brought up to date (daily data; any number of new days):

"Update Corporate Bond Indices - US with the latest FRED data."

Claude then: reads cb_conventions.md; reads the four FRED series in the built-in browser (graph CSV from the workbook's latest date); shows you the new dates and anything pending; stages the workbook from your folder and writes cb_update_workbook.py from the project; tests with --out; makes the backup in backups/ and confirms it; then replaces the master. Any STOP: nothing is written.