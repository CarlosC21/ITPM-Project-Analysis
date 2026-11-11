Project context: automating macro indicator updates
Goal

Replace manual copying into Excel with a repeatable monthly process, one indicator at a time. First target: ISM Manufacturing and Services, because they take the most time.

End-state vision

Macro indicators refresh on their release schedule, are validated, and feed the workbooks, so time goes to interpreting the data rather than copying it. Indicators that cannot be automated stay manual but are tracked in the same index. A freshness check flags any missed release.

Target flow

Scheduler triggers on release day -> pipeline requests the source -> source returns the data -> pipeline validates and transforms -> workbook is updated (with a backup) -> alert on any failure.

How it works today (ISM Manufacturing)

The owner pastes the report text into a new chat each month. The master workbook is read from the connected project folder, so it is not attached. A script parses the text, validates it, appends the new month to the workbook with matching formatting, and a dated backup of the previous version is saved first. The owner keeps a separate raw manual workbook until the process is trusted.

How it works today (US Employment)

No paste needed. Claude finds the BLS Employment Situation release through a bls.gov search, reads the headline, revisions and Tables A-1 and B-1, writes values.csv, shows it, runs employment_update_workbook.py on a copy, then writes the backup into the connected folder, confirms it, and only then replaces the master. Details: employment_conventions.md. Lesson from the first run: a hand-edited workbook (revised numbers entered before the script ran) trips the "workbook holds the previous release" check; the flag --accept-previous-mismatch exists for such a case but is a one-off, never routine.

Working approach

Each indicator moves from Manual to Parallel run to Trusted (see 01_INDICATOR_INDEX.md).

Status (Nov 6, 2026)
Done: ISM Manufacturing and Services parsers/updaters; US CPI updater (September 2026 first month); US Employment updater, restatement done, October 2026 added. Folder read/write through the connected folder now tested (employment).
Not done: automatic retrieval of the ISM source text; scheduling; all other indicators.
Next runs: ISM Manufacturing Dec 1, ISM Services Dec 3, Employment Dec 4 (first on release day), CPI mid-month.
Open blockers and questions
Source and licensing. ISM has no official public API; reports are web pages and PDFs, and historical data is a paid product. Check what the company already licenses (Bloomberg, LSEG, Haver, FactSet) before scraping, and confirm ISM's terms of use for automated retrieval.
The free short ISM PDF may omit the contraction industry lists. Confirm which source carries the full sentences; full press-release reprints include them.
Folder access depends on the Claude desktop app being open and linked to the computer. If it is not, the process falls back to an attached workbook and a downloaded result.
Scheduling a fully hands-off monthly run is not set up. It needs a legitimate way to fetch the report text and a reachable workbook location.
Seasonal-adjustment revisions: ISM revises history; decide whether to overwrite or snapshot. BLS: the annual benchmark (with the January report) will restate employment history; a refresh mode is not built.
IT constraints on API keys, outbound access and where scheduled jobs may run.
Ownership: who maintains each process when a source changes wording or format.
Web access in cloud chats: bls.gov pages open only after they appear in a bls.gov-restricted search; the project scripts live in the project and must be written out to the workspace each time (copying the script into the connected folder would let the chat stage it instead; owner's choice, not done).
Other indicators (later phases)

Likely easy: CPI, unemployment, payrolls, GDP, Fed funds, Treasury yields, M2, PPI, retail sales via FRED, BLS or BEA APIs. Moderate: jobless claims, housing starts, industrial production, sentiment. Hard or manual: paywalled surveys and consensus estimates. Verify each source before relying on it.