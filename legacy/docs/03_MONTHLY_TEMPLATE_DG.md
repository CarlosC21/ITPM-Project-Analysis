Monthly run: US Durable Goods

Paste this into a new chat (project "ITPM Project") after the full Census M3 report (Oct data: about Nov 25-Dec 4, 2026; check FRED "Updated"):

"Update US Durable Goods for the latest release."

Claude then: reads dg_conventions.md; reads the 10 FRED series in the built-in browser (last 6 months each); shows you the new month by series plus the revisions; stages the workbook from your folder and writes dg_update_workbook.py from the project; tests with --out; makes the backup in backups/ and confirms it; then replaces the master. Any STOP: nothing is written.