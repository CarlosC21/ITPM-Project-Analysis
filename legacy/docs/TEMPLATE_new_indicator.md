Template: adding a new indicator

Name the file <indicator>_conventions.md, and prefix its scripts and test file the same way.

What it is and what the workbook holds (headline values, rankings, both?).
Source: where the data comes from, how to get it, licensing or terms, and whether it was verified. Mark anything unverified.
Release schedule and time zone.
Workbook layout: sheet, block titles and rows, columns per period, labels, units, sort order, window (rolling or growing).
Transformation rules: scoring, rounding, unit conversion, seasonal adjustment, revisions.
Validation checks that stop the run (counts, names, dates, ranges, duplicates).
Known problems in the existing data (reported, not fixed).
Regression test: a sample input and the expected output.
Compare script: how the manual and automated workbooks are matched cell for cell.
Open items. Then add or update the row in 01_INDICATOR_INDEX.md.