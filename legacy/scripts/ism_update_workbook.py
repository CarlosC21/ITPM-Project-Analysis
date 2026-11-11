"""Add the next month to an ISM industry workbook (Manufacturing or Services) using report text.
 
Usage:
    python ism_update_workbook.py workbook.xlsx report.txt [--in-place | --out new.xlsx] [--window N]
 
Which indicator: detected from the block titles in column A (override: --indicator
manufacturing|services).
 
Month: worked out from the workbook (the month after its newest column). The report text must
name that same month (for example "October") or the run stops.
 
Default mode is APPEND: the new month becomes a new pair of columns to the right of the newest
one, with cell formatting and conditional formatting extended to match. With --window N the
workbook keeps only the newest N months instead (the oldest drops off and no columns are added).
 
Output:
  --in-place  First copies the original to a backups/ folder next to it (never skipped), then
              replaces the workbook with the updated one.
  --out FILE  Writes a separate file and leaves the original untouched (default name is dated).
 
Sheet conventions this relies on (see ism_manufacturing_conventions.md):
  * 7 blocks, each a header row of month dates (labels in B, D, F, ...; score in the next column)
    followed by 2 spacer rows and exactly 18 industry rows.
  * Score: first-listed industry gets the largest magnitude. Growth-side = +(n - pos + 1),
    contraction-side = -(m - pos + 1), neutral = 0. (Verified against the September 2026 report.)
  * Rows are re-sorted by the newest score, descending, keeping prior order for ties.
  * Broken '=IFERROR(VLOOKUP(...#REF!...))' cells are replaced by the value they already
    evaluate to (their IFERROR fallback), because re-sorting would otherwise break them.
  * Prints an audit of anomalies found in the existing data. Never changes history.
"""
import argparse
import copy
import datetime
import os
import re
import shutil
import sys
 
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl  # noqa: E402
from openpyxl.formatting.rule import Rule  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
import ism_industry_parser as P  # noqa: E402
 
# Each block: (title in column A, parser index name, label for the up side, label for the down side).
# Titles are matched ignoring case and repeated spaces.
PROFILES = {
    "manufacturing": {
        "industries": P.INDUSTRIES,
        "blocks": [
            ("ISM Manufacturing", "Manufacturing PMI", "Growth", "Contraction"),
            ("NEW ORDERS", "New Orders", "Growth", "Contraction"),
            ("PRODUCTION", "Production", "Growth", "Contraction"),
            ("EMPLOYMENT", "Employment", "Growth", "Contraction"),
            ("DELIVERIES", "Supplier Deliveries", "Slower", "Faster"),
            ("INVENTORIES", "Inventories", "Higher", "Lower"),
            ("CUSTOMER INVENTORIES", "Customer Inventories", "Too High", "Too Low"),
        ],
    },
    "services": {
        "industries": P.SERVICES_INDUSTRIES,
        "blocks": [
            ("ISM Non-Manuf", "Services PMI", "Growth", "Contraction"),
            ("Business Activity", "Business Activity", "Growth", "Contraction"),
            ("New Orders", "New Orders", "Growth", "Contraction"),
            ("EMPLOYMENT", "Employment", "Growth", "Contraction"),
            ("DELIVERIES", "Supplier Deliveries", "Slower", "Faster"),
            ("INVENTORIES", "Inventories", "Increasing", "Decreasing"),
        ],
    },
}
UP = {"Growth", "Slower", "Higher", "Too High", "Increasing"}
DOWN = {"Contraction", "Faster", "Lower", "Too Low", "Decreasing"}
N_ROWS = 18
FALLBACK = re.compile(r',\s*("[^"]*"|-?\d+(?:\.\d+)?)\)\s*$')
 
 
def next_month(d):
    return datetime.datetime(d.year + (d.month == 12), d.month % 12 + 1, 1)
 
 
def pair_cols(k):
    """Columns (label, score) of the k-th month pair, 0-based: B:C, D:E, F:G, H:I ..."""
    return 2 + 2 * k, 3 + 2 * k
 
 
def cell_value(ws, wsv, row, col, converted):
    v = ws.cell(row, col).value
    if isinstance(v, str) and v.startswith("="):
        cached = wsv.cell(row, col).value
        if cached is None:
            m = FALLBACK.search(v)
            if not m or "#REF!" not in v:
                raise ValueError(f"cannot resolve formula at {ws.cell(row, col).coordinate}: {v}")
            tok = m.group(1)
            cached = tok.strip('"') if tok.startswith('"') else float(tok)
            cached = int(cached) if isinstance(cached, float) and cached.is_integer() else cached
        converted.append(ws.cell(row, col).coordinate)
        return cached
    return v
 
 
def _norm(v):
    return " ".join(str(v).split()).lower() if v is not None else ""
 
 
def detect_profile(ws, name=None):
    """Pick the profile whose block titles are all present in column A."""
    titles = {_norm(ws.cell(r, 1).value) for r in range(1, ws.max_row + 1)}
    for key, prof in PROFILES.items():
        if (name or key) == key and all(_norm(b[0]) in titles for b in prof["blocks"]):
            return key, prof
    raise ValueError("workbook block titles in column A match neither the Manufacturing nor the "
                     "Services layout" + (f" (asked for {name})" if name else ""))
 
 
def find_header(ws, title):
    rows = [r for r in range(1, ws.max_row + 1) if _norm(ws.cell(r, 1).value) == _norm(title)]
    if len(rows) != 1:
        raise ValueError(f"expected one header '{title}', found {len(rows)}")
    return rows[0]
 
 
def read_block(ws, wsv, hdr, converted, industries):
    dates = []
    while isinstance(ws.cell(hdr, pair_cols(len(dates))[0]).value, datetime.datetime):
        dates.append(ws.cell(hdr, pair_cols(len(dates))[0]).value)
    if not dates:
        raise ValueError(f"the block at row {hdr} has no month date headers in B, D, F ...")
    for a, b in zip(dates, dates[1:]):
        if b != next_month(a):
            raise ValueError(f"the block at row {hdr} has months out of order or with a gap: "
                             f"{a:%b-%y} then {b:%b-%y}")
    rows = []
    for r in range(hdr + 3, hdr + 3 + N_ROWS):
        ind = ws.cell(r, 1).value
        months = [(cell_value(ws, wsv, r, pair_cols(k)[0], converted),
                   cell_value(ws, wsv, r, pair_cols(k)[1], converted)) for k in range(len(dates))]
        rows.append([ind, months])
    if ws.cell(hdr + 3 + N_ROWS, 1).value not in (None, ""):
        raise ValueError(f"block at row {hdr} has more than {N_ROWS} industry rows")
    if sorted(r[0] for r in rows) != sorted(industries):
        raise ValueError(f"block at row {hdr}: industry names differ from the canonical 18")
    return dates, rows
 
 
def audit(name, dates, rows):
    notes = []
    for mi, d in enumerate(dates):
        tag = f"{name} {d:%b-%y}"
        pos, neg = [], []
        for ind, months in rows:
            label, score = months[mi]
            if label in UP and not score > 0:
                notes.append(f"{tag}: {ind} is '{label}' but score is {score}")
            elif label in DOWN and not score < 0:
                notes.append(f"{tag}: {ind} is '{label}' but score is {score}")
            elif str(label).lower() == "neutral" and score != 0:
                notes.append(f"{tag}: {ind} is '{label}' but score is {score}")
            (pos if isinstance(score, (int, float)) and score > 0 else neg if score < 0 else []).append(score)
        if sorted(pos) != list(range(1, len(pos) + 1)):
            notes.append(f"{tag}: positive scores {sorted(pos)} are not 1..{len(pos)}")
        if sorted(neg) != [-i for i in range(len(neg), 0, -1)]:
            notes.append(f"{tag}: negative scores {sorted(neg)} are not -{len(neg)}..-1")
    return notes
 
 
def new_month_scores(index, text, month_num, up_label, down_label, industries):
    rows, problems = P.parse_section(index, text, month_num, industries)
    if problems:
        raise ValueError(f"{index}: " + "; ".join(problems))
    allowed = {"Growth": {"growth", "contraction", "increase", "decrease"},
               "Slower": {"slower", "faster"}, "Higher": {"higher", "lower"},
               "Too High": {"too high", "too low"}, "Increasing": {"increase", "decrease"}}[up_label]
    for r in rows:
        if r["category"] not in allowed and r["direction"] != "neutral":
            raise ValueError(f"{index}: found a '{r['category']}' sentence, which does not belong "
                             f"in this index; the wording or the headings may be wrong")
    out = {}
    for direction in ("up", "down"):
        items = [r for r in rows if r["direction"] == direction]
        m = len(items)
        for r in items:
            mag = m - r["rank"] + 1
            out[r["industry"]] = ((up_label, mag) if direction == "up" else (down_label, -mag))
    for r in rows:
        if r["direction"] == "neutral":
            out[r["industry"]] = ("Neutral", 0)
    return out
 
 
# ---- formatting for appended columns ---------------------------------------------------------
 
def _cf_ranges(ws):
    for cf in ws.conditional_formatting:
        for cr in cf.sqref.ranges:
            yield cr, cf.rules
 
 
def _has_cf(ws, col, r0, r1):
    return any(cr.min_col <= col <= cr.max_col and not (cr.max_row < r0 or cr.min_row > r1)
               for cr, _ in _cf_ranges(ws))
 
 
def _collect_rules(ws, col, r0, r1):
    found = {}
    for cr, rules in _cf_ranges(ws):
        if cr.min_col == cr.max_col == col and cr.min_row >= r0 and cr.max_row <= r1:
            for rule in rules:
                key = (rule.type, rule.operator, rule.text,
                       repr(rule.colorScale) if rule.colorScale else None, repr(rule.dxf))
                found.setdefault(key, rule)
    return list(found.values())
 
 
def extend_formatting(ws, hdr, prev_cols, new_cols):
    """Copy cell styles and conditional formatting from the previous month pair to the new one."""
    r0, r1 = hdr + 3, hdr + 2 + N_ROWS
    notes = []
    for pc, nc in zip(prev_cols, new_cols):
        for r in range(hdr, r1 + 1):
            ws.cell(r, nc)._style = copy.copy(ws.cell(r, pc)._style)
        if _has_cf(ws, nc, r0, r1):
            notes.append(f"conditional formatting already present in column {get_column_letter(nc)}; left as is")
            continue
        letter = get_column_letter(nc)
        for rule in _collect_rules(ws, pc, r0, r1):
            if rule.type == "containsText" and rule.text:
                new = Rule(type="containsText", operator=rule.operator, text=rule.text,
                           dxf=copy.deepcopy(rule.dxf), stopIfTrue=rule.stopIfTrue,
                           formula=[f'NOT(ISERROR(SEARCH("{rule.text}",{letter}{r0})))'])
            else:
                new = copy.deepcopy(rule)
            new.priority = 0  # let openpyxl assign the next free priority
            ws.conditional_formatting.add(f"{letter}{r0}:{letter}{r1}", new)
    return notes
 
 
def backup_path(workbook, new_date):
    folder = os.path.join(os.path.dirname(os.path.abspath(workbook)), "backups")
    stem, ext = os.path.splitext(os.path.basename(workbook))
    path = os.path.join(folder, f"{stem}_before_{new_date:%Y-%m}{ext}")
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem}_before_{new_date:%Y-%m}_{n}{ext}")
        n += 1
    return path
 
 
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("report_text")
    ap.add_argument("month", nargs="?", help="optional YYYY-MM; must equal the month after the sheet's newest")
    ap.add_argument("--indicator", choices=sorted(PROFILES), help="default: detected from the workbook")
    ap.add_argument("--window", type=int, help="keep only the newest N months (default: keep all, append)")
    dest = ap.add_mutually_exclusive_group()
    dest.add_argument("--in-place", action="store_true", help="back up the original, then replace it")
    dest.add_argument("--out", help="write a separate file; the original is untouched")
    a = ap.parse_args()
 
    wb = openpyxl.load_workbook(a.workbook)
    wbv = openpyxl.load_workbook(a.workbook, data_only=True)
    ws, wsv = wb.active, wbv.active
    sections = P.split_sections(open(a.report_text, encoding="utf-8").read())
 
    converted, audit_notes, blocks = [], [], []
    try:
        pname, prof = detect_profile(ws, a.indicator)
        industries = prof["industries"]
        for title, index, up_l, down_l in prof["blocks"]:
            hdr = find_header(ws, title)
            dates, rows = read_block(ws, wsv, hdr, converted, industries)
            audit_notes += audit(title, dates, rows)
            blocks.append((title, index, hdr, dates, rows, up_l, down_l))
    except ValueError as e:
        sys.exit(f"STOP: {e}")
    print(f"Workbook layout: {pname}")
 
    if len({len(b[3]) for b in blocks}) != 1:
        sys.exit("STOP: blocks have different numbers of months")
    latest = {b[3][-1] for b in blocks}
    if len(latest) != 1:
        sys.exit("STOP: blocks end on different months: " + ", ".join(sorted(f"{d:%b-%y}" for d in latest)))
    latest = latest.pop()
    k = len(blocks[0][3])
    new_date = next_month(latest)
    if a.month:
        year, mon = (int(x) for x in a.month.split("-"))
        if datetime.datetime(year, mon, 1) != new_date:
            sys.exit(f"STOP: sheet ends at {latest:%b-%y}, so the month to add is {new_date:%Y-%m}, not {a.month}")
    mon = new_date.month
    if a.window is not None and a.window < k:
        sys.exit(f"STOP: sheet has {k} months, more than --window {a.window}; refusing to drop history")
    roll = a.window is not None and k >= a.window
    print(f"Adding {new_date:%b-%y} (sheet currently ends at {latest:%b-%y}); "
          + ("rolling the window" if roll else "appending a new month"))
 
    if not a.in_place:
        out = a.out or os.path.splitext(os.path.basename(a.workbook))[0] + f"_{new_date:%Y-%m}.xlsx"
        if os.path.abspath(out) == os.path.abspath(a.workbook):
            sys.exit("STOP: refusing to overwrite the input workbook (use --in-place for that)")
 
    plan = []
    for title, index, hdr, dates, rows, up_l, down_l in blocks:
        if index not in sections:
            sys.exit(f"STOP: report text has no section for {index}")
        try:
            new = new_month_scores(index, sections[index], mon, up_l, down_l, industries)
        except ValueError as e:
            sys.exit(f"STOP: {e}")
        if not roll:  # never write over other content (for example a helper area to the right)
            for c in pair_cols(k):
                busy = [ws.cell(r, c).coordinate for r in range(hdr, hdr + 3 + N_ROWS)
                        if ws.cell(r, c).value not in (None, "")]
                if busy:
                    sys.exit(f"STOP: the new month would go into column {get_column_letter(c)}, but "
                             f"{busy[0]} already holds something. Move the other content to the right "
                             f"first (nothing was written).")
        plan.append((title, index, hdr, dates, rows, new))
 
    summary, fmt_notes = [], []
    for title, index, hdr, dates, rows, new in plan:
        for r in rows:
            r[1] = (r[1][1:] if roll else r[1]) + [new[r[0]]]
        dates = (dates[1:] if roll else dates) + [new_date]
        rows.sort(key=lambda r: -r[1][-1][1])  # stable: ties keep previous order
        if not roll:
            fmt_notes += extend_formatting(ws, hdr, pair_cols(k - 1), pair_cols(k))
            for pc, nc in zip(pair_cols(k - 1), pair_cols(k)):  # same width as the previous month
                spans = [(d.min, d.max) for d in ws.column_dimensions.values() if d.min and d.max]
                if not any(lo <= nc <= hi for lo, hi in spans):
                    w = ws.column_dimensions[get_column_letter(pc)].width
                    if w:
                        ws.column_dimensions[get_column_letter(nc)].width = w
        for j, d in enumerate(dates):
            ws.cell(hdr, pair_cols(j)[0]).value = d
        for i, (ind, months) in enumerate(rows):
            r = hdr + 3 + i
            ws.cell(r, 1).value = ind
            for j, (label, score) in enumerate(months):
                lc, sc = pair_cols(j)
                ws.cell(r, lc).value, ws.cell(r, sc).value = label, score
        counts = {}
        for _, months in rows:
            counts[months[-1][0]] = counts.get(months[-1][0], 0) + 1
        summary.append(f"{title}: {counts}")
 
    if a.in_place:
        backup = backup_path(a.workbook, new_date)
        os.makedirs(os.path.dirname(backup), exist_ok=True)
        shutil.copy2(a.workbook, backup)  # backup is made before anything is replaced
        tmp = a.workbook + ".tmp"
        wb.save(tmp)
        os.replace(tmp, a.workbook)
        print("Backup of the original:", backup)
        print("Updated in place:", a.workbook)
    else:
        wb.save(out)
        print("Saved:", out)
    cols = f"{get_column_letter(pair_cols(k - 1 if roll else k)[0])}:{get_column_letter(pair_cols(k - 1 if roll else k)[1])}"
    print(f"New month is in columns {cols}; the workbook now holds "
          f"{k if roll else k + 1} months ({dates[0]:%b-%y} to {new_date:%b-%y}).")
    for n in sorted(set(fmt_notes)):
        print("Note:", n)
    print("\nNew month counts:")
    print("\n".join("  " + s for s in summary))
    print(f"\nFormula cells replaced by their current values (so re-sorting cannot break them): {len(converted)}")
    print(f"\nAnomalies found in the EXISTING data ({len(audit_notes)}):")
    for n in audit_notes:
        print("  -", n)
 
 
if __name__ == "__main__":
    main()