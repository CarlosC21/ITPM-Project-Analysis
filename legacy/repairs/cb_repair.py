#!/usr/bin/env python3
"""One-off repair of 'Corporate Bond Indices - US.xlsx' against FRED, matched BY DATE.
Usage: python cb_repair.py WORKBOOK fred_window.csv --out NEW.xlsx --log changes.csv
fred_window.csv = date,US10YR,AAA,BBB,CCC for every FRED date from the window start to the newest workbook date (blank = no value).
Only dates where all four exist are used. Rows before the window are never touched. Charts: range starts forced to row 3."""
import argparse, csv, datetime as dt, os, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cb_update_workbook as U
 
def setrow(x, n):
    x = re.sub(r'<row r="\d+"', '<row r="%d"' % n, x, count=1)
    return re.sub(r'<c r="([A-Z]+)\d+"', lambda m: '<c r="%s%d"' % (m.group(1), n), x)
 
def setcell(x, col, val):
    pat = re.compile(r'(<c r="%s\d+"[^>]*>)<v>[^<]*</v>' % col)
    if not pat.search(x): U.stop("cell %s has no value to replace" % col)
    return pat.sub(lambda m: m.group(1) + '<v>%r</v>' % val, x, count=1)
 
def main():
    ap = argparse.ArgumentParser(); ap.add_argument("wb"); ap.add_argument("fred"); ap.add_argument("--out", required=True); ap.add_argument("--log", required=True)
    a = ap.parse_args()
    try:
        book = U.Book(a.wb)
        fred = {}
        for d, v in U.read_values(a.fred):
            if None not in v: fred[d] = v
        start = min(fred)
        newest = U.from_serial(book.idx[0][1][0])
        if max(fred) != newest: U.stop("fred file must end on the workbook's newest date %s (ends %s)" % (newest, max(fred)))
        log = []
        parts = dict(book.parts)
        # ---------- Indices
        head, body, tail = U.split_sheet(parts[U.INDICES].decode())
        rows = U.rows_of(body); keep = [x for r, x in rows if r < U.FIRST]
        data = [(r, x, U.from_serial(U.cell_vals(x)["B"])) for r, x in rows if r >= U.FIRST]
        have = {d for r, x, d in data}
        out, ins = [], 0
        for r, x, d in data:
            if d >= start:
                if d not in fred: U.stop("Workbook date %s (in window) has no complete FRED values" % d)
                cv = U.cell_vals(x)
                for col, name, new in zip("CDEF", U.NAMES, fred[d]):
                    old = cv.get(col)
                    if old is None or abs(old - new) > 1e-9:
                        log.append(("Indices", d, name, old, new, "value corrected")); x = setcell(x, col, new)
            out.append((d, x))
        for d, v in fred.items():
            if d not in have:
                t, aa, bb, c = v
                x = ('<row r="0" spans="1:6" x14ac:dyDescent="0.25"><c r="B0" s="8"><v>%d</v></c><c r="C0" s="1"><v>%r</v></c><c r="D0" s="1"><v>%r</v></c>'
                     '<c r="E0" s="1"><v>%r</v></c><c r="F0" s="1"><v>%r</v></c></row>' % (U.serial(d), t, aa, bb, c))
                out.append((d, x)); ins += 1
                log.append(("Indices", d, "row", None, "%s" % v, "missing date inserted"))
        out.sort(key=lambda p: p[0], reverse=True)
        ix_final = {d: x for d, x in out}
        idx_new = "".join(setrow(x, U.FIRST + i) for i, (d, x) in enumerate(out))
        last = U.FIRST + len(out) - 1
        head = re.sub(r'<dimension ref="A1:F\d+"/>', '<dimension ref="A1:F%d"/>' % last, head)
        parts[U.INDICES] = (head + "".join(keep) + idx_new + tail).encode()
        # ---------- Spreads
        head, body, tail = U.split_sheet(parts[U.SPREADS].decode())
        rows = U.rows_of(body); n_ins = ins
        blanks = [x for r, x in rows[len(rows) - n_ins:]] if n_ins else []
        if any("<v>" in x for x in blanks): U.stop("Spreads trailing rows are not blank")
        rows = rows[:len(rows) - n_ins] if n_ins else rows
        keep = [x for r, x in rows if r < U.FIRST]
        sdata = [(U.from_serial(U.cell_vals(x)["A"]), x) for r, x in rows if r >= U.FIRST and U.cell_vals(x).get("A") is not None]
        sdone, sout = set(), []
        for d, x in sdata:
            if d >= start:
                t, aa, bb, c = fred[d]; exp = [aa - bb, aa - c, bb - c, t - aa, t - bb, t - c]
                cv = U.cell_vals(x)
                for col, name, e in zip("BCDEFG", U.SPREAD_HDR, exp):
                    old = cv.get(col)
                    if old is None or abs(old - e) > 1e-9:
                        log.append(("Spreads", d, name, old, e, "value corrected")); x = setcell(x, col, e)
                sdone.add(d)
            sout.append((d, x))
        have_s = {d for d, x in sdata}
        for d in fred:
            if d not in have_s and d >= min(have_s):
                t, aa, bb, c = fred[d]; sp = [aa - bb, aa - c, bb - c, t - aa, t - bb, t - c]
                x = '<row r="0" spans="1:7" x14ac:dyDescent="0.25"><c r="A0" s="8"><v>%d</v></c>%s</row>' % (
                    U.serial(d), "".join('<c r="%s0" s="7"><v>%r</v></c>' % (col, e) for col, e in zip("BCDEFG", sp)))
                sout.append((d, x)); log.append(("Spreads", d, "row", None, "6 spreads", "missing date inserted"))
        sout.sort(key=lambda p: p[0], reverse=True)
        snew = "".join(setrow(x, U.FIRST + i) for i, (d, x) in enumerate(sout))
        rest = "".join(setrow(x, U.FIRST + len(sout) + i) for i, x in enumerate(blanks_pad(rows, len(sdata), n_ins)))
        parts[U.SPREADS] = (head + "".join(keep) + snew + rest + tail).encode()
        # ---------- charts: range start -> row 3
        for i in list(range(1, 21)):
            nm = "xl/charts/chart%d.xml" % i
            s = parts[nm].decode()
            s2 = re.sub(r"(<c:f>'US Corp Yield (?:Indices|Spreads)'!\$[A-Z]\$)(?:57|37|4)(:\$[A-Z]\$\d+</c:f>)", r"\g<1>3\2", s)
            if s2 != s:
                parts[nm] = s2.encode(); log.append(("Chart", None, "chart%d" % i, "range started below row 3", "starts at row 3", "chart range fixed"))
        tmp = a.out + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for nm in book.order: zo.writestr(nm, parts[nm])
        os.replace(tmp, a.out)
        with open(a.log, "w", newline="") as f:
            w = csv.writer(f); w.writerow(["sheet", "date", "item", "old", "new", "action"])
            for s, d, k, o, n, act in sorted(log, key=lambda z: (z[0], str(z[1]), z[2])): w.writerow([s, d, k, o, n, act])
        print("changes logged:", len(log)); return 0
    except U.Stop as e:
        print("STOP:", e); return 2
 
def blanks_pad(rows, ndata, n_ins):
    # trailing blank styled rows to keep the sheet length (take them from the existing blank tail)
    tail = [x for r, x in rows if r >= U.FIRST][ndata:]
    return tail
 
if __name__ == "__main__":
    sys.exit(main())