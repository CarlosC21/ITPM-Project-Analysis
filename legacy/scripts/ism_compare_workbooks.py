"""Compare a manual ISM workbook (Manufacturing or Services) against an automated one, cell by cell.
 
Usage:
    python ism_compare_workbooks.py manual.xlsx automated.xlsx [--csv diffs.csv]
 
Cells are matched by block, industry and month (taken from the header date), not by column
position, so the two files can have different month windows. Only months present in both files
are compared; months in only one file are listed. Labels are compared ignoring case and spaces.
Exits with status 1 if any difference is found.
"""
import argparse
import csv
import os
import sys
 
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl  # noqa: E402
import ism_update_workbook as U  # noqa: E402
 
 
def load(path):
    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    ws, wsv = wb.active, wbv.active
    data, months = {}, set()
    _, prof = U.detect_profile(ws)
    for title, *_ in prof["blocks"]:
        hdr = U.find_header(ws, title)
        dates, rows = U.read_block(ws, wsv, hdr, [], prof["industries"])
        for d in dates:
            months.add(d.strftime("%Y-%m"))
        for ind, vals in rows:
            for d, (label, score) in zip(dates, vals):
                data[(title, ind, d.strftime("%Y-%m"))] = (label, score)
    return data, months
 
 
def norm(label):
    return str(label).strip().lower()
 
 
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manual")
    ap.add_argument("automated")
    ap.add_argument("--csv")
    a = ap.parse_args()
    try:
        man, man_months = load(a.manual)
        aut, aut_months = load(a.automated)
    except ValueError as e:
        sys.exit(f"STOP: {e}")
    common = sorted(man_months & aut_months)
    print("Months only in manual file:   ", sorted(man_months - aut_months) or "none")
    print("Months only in automated file:", sorted(aut_months - man_months) or "none")
    print("Months compared:              ", common or "none")
    if not common:
        sys.exit("no months in common; nothing to compare")
    diffs, checked = [], 0
    for (title, ind, month), (ml, ms) in sorted(man.items(), key=lambda kv: (kv[0][2], kv[0][0], kv[0][1])):
        if month not in common:
            continue
        al, as_ = aut[(title, ind, month)]
        checked += 1
        if norm(ml) != norm(al) or ms != as_:
            diffs.append({"month": month, "block": title, "industry": ind,
                          "manual": f"{ml} / {ms}", "automated": f"{al} / {as_}"})
    print(f"\nCells compared: {checked}   Differences: {len(diffs)}")
    for d in diffs[:60]:
        print(f"  {d['month']} | {d['block']} | {d['industry']}: manual {d['manual']}  vs  automated {d['automated']}")
    if len(diffs) > 60:
        print(f"  ... {len(diffs) - 60} more (use --csv to see all)")
    if a.csv:
        with open(a.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["month", "block", "industry", "manual", "automated"])
            w.writeheader()
            w.writerows(diffs)
        print("Wrote", a.csv)
    sys.exit(1 if diffs else 0)
 
 
if __name__ == "__main__":
    main()