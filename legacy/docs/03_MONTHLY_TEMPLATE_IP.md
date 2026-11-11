Monthly run: US Industrial Production

Paste this into a new chat (project "ITPM Project") after the G.17 release (Oct data: Nov 17, 2026):

"Update US Industrial Production for the latest G.17 release."

Claude then: reads ip_conventions.md; reads the 21 FRED series in the built-in browser (last 5 months each); shows you the new month by series plus the revisions; stages the workbook from your folder and writes ip_update_workbook.py from the project; tests with --out; makes the backup in backups/ and confirms it; then replaces the master. Any STOP: nothing is written.