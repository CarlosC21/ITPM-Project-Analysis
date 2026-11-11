# legacy/

A frozen copy of the files from the "ITPM Project" (the Excel-workbook automation
this repo grew out of). The master copies of these files, the master workbooks and
the sample csv/txt inputs live in the separate ITPM-Project folder, not here.

- `docs/` - indicator index, project context, monthly chat templates.
- `conventions/` - one conventions file per workbook. Read it before touching that indicator.
- `scripts/` - the update scripts.
- `repairs/` - one-off repair scripts.

Paths such as `claude/...` and "the connected project folder" inside these files
refer to the ITPM Project, not to this repo.

## Changes made to this copy (Nov 11, 2026)

File names only; no contents were changed.

- Conventions files renamed from short names (`employment`, `cpi`, ...) to
  `<name>_conventions.md`, the names the index and the scripts use.
- Eight scripts renamed from `<name>.py` to `<name>_update_workbook.py`, the names
  the index uses.
- `docs/03_MONTHLY_TEMPLATE` renamed to `docs/03_MONTHLY_TEMPLATE_MANUFACTURING.md`.
  It is the ISM Manufacturing template; the index gives it no name, so this name
  was chosen to follow the other templates.

## Known issues in the docs (not corrected in the copies)

- `conventions/employment_conventions.md`, line 14: says the Oct-27 unemployment
  YoY "will also be blank because the year-ago month is blank". The script
  (`scripts/employment_update_workbook.py`, `expected()`) leaves YoY blank only
  when the year-ago level is blank; Oct-26 has a rate, so it will calculate Oct-27.
- `conventions/employment_conventions.md`, line 47: refers to the "Oct 6 release"
  and an "Oct 2 edition". The index dates the October 2026 data release Nov 6, 2026.
  Which dates are meant is an open question for the owner.

## Not verified

- Whether the scripts here are identical to the master copies in ITPM-Project.

## Missing

- `repairs/cb_repair_log_2026-11-09.csv` (listed in the index; to be copied from
  ITPM-Project by the owner).
- There is no CPI monthly template; the index does not list one either.
