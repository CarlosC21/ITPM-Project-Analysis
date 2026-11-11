# ITPM Project Analysis (portfolio project)

## What this is

Portfolio project for data engineering roles. Goal: user types a stock symbol on a web page; the backend reads macro indicator data and highlights which indicators matter most for that stock.

Origin: working Python scripts that update monthly Excel workbooks (project "ITPM Project"). This repo turns that into a proper pipeline (database, orchestration, API, cloud).

## Indicators (11 workbooks, all built; see legacy/docs/01_INDICATOR_INDEX.md)

ISM Manufacturing, ISM Services, US CPI, US Employment, US Industrial Production, US Durable Goods, NFIB Components, NFIB Industries, Corporate Bond Indices, CFTC Commitments of Traders, Bull & Bear Markets (S&P 500).

- Status is "parallel run" or "built, tested on copies". None is "trusted"; only the owner changes that status.
- Sources differ per indicator (BLS, ISM pasted text, FRED, NFIB, CFTC, Yahoo). The index lists each one. Do not add or swap a source without asking.
- Not started: GDP Correlation, M2 Money Supply, UMCSI, US Building Permits, US Unemployment Claims. Not in the index yet.
- Read the indicator's conventions file in legacy/conventions/ before touching it. Do not re-derive layouts or rules.

## How to work with me (most important)

- I am learning. Work in small steps. Before each step say what it is and why it exists.
- Do not guess. If something is unclear or unverified, say so and ask.
- Plan first, wait for my approval, then change files.
- Never invent sources, series IDs, URLs or release dates.
- Do not hand-edit data values or silently fix historical data.
- If a validation check fails (STOP), report it exactly and write nothing.
- Back up before overwriting anything.

## Hard rules

- The master Excel workbooks live in the separate ITPM-Project folder and are mine. Never edit, move or delete them. Work on copies only.
- Workbooks and backups are not committed to git (.gitignore).
- Keep raw source text and tables unchanged. Revisions are new rows, not overwrites.
- Loads must be idempotent (re-running a month never duplicates data).

## Planned design (proposed, not final; confirm before building)

- Star schema: dim_indicator, dim_date, dim_sector, fact_observation; reference month kept separate from release date.
- Frequencies differ (daily, weekly, monthly): the schema must handle this.
- Layers: raw -> cleaned -> served.
- Order: schema -> extractors with automated data-quality checks -> orchestration -> API and web page -> relevance ranking by stock -> AWS and Terraform -> MCP server.
- Big data tools only where the data justifies them (e.g. stock price history), not for small macro series.

## Environment

- Windows, PowerShell, VS Code. Claude Pro plan (usage shared with chat), so keep sessions narrow.
- Use /clear between unrelated tasks and /compact when a task continues.

## Portfolio deliverables

README with architecture diagram, ERD, decisions log, tests, CI, and a "what failed and how it recovered" section (use the repair scripts and logs in legacy/repairs/).

## Monthly Excel updates

- Not done in this repo. Real monthly updates of the master workbooks stay in the
  "ITPM Project" chat, which has the connected folder and the browser access the
  scripts' sources need.
- Sample csv/txt files are not in this repo (kept in ITPM-Project). If a task needs
  test input, ask me, or build small synthetic fixtures; do not invent real values.
- Each conventions file describes the input a script expects.
