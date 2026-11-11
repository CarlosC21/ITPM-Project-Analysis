#!/usr/bin/env python3
"""Update 'Corporate Bond Indices - US.xlsx' with new daily FRED values.
 
Usage:
  python cb_update_workbook.py WORKBOOK values.csv --out NEW.xlsx
  python cb_update_workbook.py WORKBOOK values.csv --in-place      (backup first)
  python cb_update_workbook.py WORKBOOK --audit                    (read-only checks)
 
values.csv: header  date,US10YR,AAA,BBB,CCC   (YYYY-MM-DD; blank = FRED has no value that day)
It must include the workbook's latest date (and ideally a few earlier ones) as an overlap check.
 
Standard library only. Edits sheet XML directly, so the 20 charts and everything else are
copied byte for byte. Exit code 2 = STOP (nothing written).
"""
import argparse, csv, datetime as dt, io, os, re, shutil, sys, zipfile
 
SPREADS = "xl/worksheets/sheet1.xml"   # 'US Corp Yield Spreads'
INDICES = "xl/worksheets/sheet3.xml"   # 'US Corp Yield Indices'
EPOCH = dt.date(1899, 12, 30)
FIRST = 3                               # first data row on both sheets
NAMES = ["US10YR", "AAA", "BBB", "CCC"]
MAXMOVE = {"US10YR": 1.0, "AAA": 1.0, "BBB": 1.0, "CCC": 3.0}   # largest plausible 1-day move (pct points)
RANGE = {"US10YR": (0, 20), "AAA": (0, 20), "BBB": (0, 20), "CCC": (0, 60)}
SPREAD_HDR = ["AAA-BBB", "AAA-CCC", "BBB-CCC", "10yr-AAA", "10yr-BBB", "10yr-CCC"]
INDEX_HDR = ["US10YR", "ICE_BoA_AAA", "ICE_BoA_BBB", "ICE_BoA_CCC_or_below"]
 
 
class Stop(Exception):
    pass
 
 
def stop(msg):
    raise Stop(msg)
 
 
def serial(d):
    return (d - EPOCH).days
 
 
def from_serial(n):
    return EPOCH + dt.timedelta(days=int(n))
 
 
ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
 
 
def split_sheet(xml):
    a = xml.index("<sheetData>") + len("<sheetData>")
    b = xml.index("</sheetData>")
    return xml[:a], xml[a:b], xml[b:]
 
 
def rows_of(body):
    return [(int(m.group(1)), m.group(0)) for m in ROW_RE.finditer(body)]
 
 
def cell_vals(rowxml):
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', rowxml, re.S):
        col, attrs, inner = m.group(1), m.group(2), m.group(3) or ""
        v = re.search(r"<v>([^<]*)</v>", inner)
        if v is None:
            out[col] = None
        elif 't="s"' in attrs:
            out[col] = ("s", int(v.group(1)))
        else:
            out[col] = float(v.group(1))
    return out
 
 
def shift_row(rowxml, n):
    def rr(m):
        return '<row r="%d"' % (int(m.group(1)) + n)
    rowxml = re.sub(r'<row r="(\d+)"', rr, rowxml, count=1)
    return re.sub(r'<c r="([A-Z]+)(\d+)"', lambda m: '<c r="%s%d"' % (m.group(1), int(m.group(2)) + n), rowxml)
 
 
def shared_strings(z):
    s = z.read("xl/sharedStrings.xml").decode("utf-8")
    return [re.sub(r"<[^>]+>", "", x) for x in re.findall(r"<si>(.*?)</si>", s, re.S)]
 
 
class Book:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.parts = {i.filename: self.z.read(i.filename) for i in self.z.infolist()}
        self.order = [i.filename for i in self.z.infolist()]
        wb = self.parts["xl/workbook.xml"].decode()
        names = re.findall(r'<sheet name="([^"]+)"', wb)
        if names != ["US Corp Yield Spreads", "Sheet2", "US Corp Yield Indices", "Sheet1"]:
            stop("Unexpected sheets %s" % names)
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode()
        for rid, tgt in (("rId1", "worksheets/sheet1.xml"), ("rId3", "worksheets/sheet3.xml")):
            if not re.search(r'Id="%s"[^>]*Target="%s"|Target="%s"[^>]*Id="%s"' % (rid, tgt, tgt, rid), rels):
                stop("Sheet mapping changed (%s)" % rid)
        nch = len([n for n in self.parts if re.match(r"xl/charts/chart\d+\.xml$", n)])
        if nch != 20:
            stop("Expected 20 charts, found %d" % nch)
        ss = shared_strings(self.z)
        sp = {r: cell_vals(x) for r, x in rows_of(split_sheet(self.parts[SPREADS].decode())[1]) if r <= 2}
        ix = {r: cell_vals(x) for r, x in rows_of(split_sheet(self.parts[INDICES].decode())[1]) if r <= 2}
        got = [ss[sp[2][c][1]] for c in "BCDEFG"]
        if got != SPREAD_HDR:
            stop("Spreads headers changed: %s" % got)
        got = [ss[ix[2][c][1]] for c in "CDEF"]
        if got != INDEX_HDR:
            stop("Indices headers changed: %s" % got)
        self.idx = self._read(INDICES, "BCDEF")
        self.spr = self._read(SPREADS, "ABCDEFG")
 
    def _read(self, part, cols):
        out = []
        for r, x in rows_of(split_sheet(self.parts[part].decode())[1]):
            if r < FIRST:
                continue
            v = cell_vals(x)
            if v.get(cols[0]) is None:
                continue
            out.append((r, [v.get(c) for c in cols]))
        return out
 
 
def audit(book, n=250):
    notes = []
    d = [from_serial(v[0]) for r, v in book.idx]
    bad = [str(d[i]) for i in range(len(d) - 1) if not d[i] > d[i + 1]]
    notes.append("Indices: %d dated rows, newest %s, oldest %s; order problems: %s" % (len(d), d[0], d[-1], bad[:5] or "none"))
    sd = [from_serial(v[0]) for r, v in book.spr]
    notes.append("Spreads: %d dated rows, newest %s, oldest %s" % (len(sd), sd[0], sd[-1]))
    if sd[0] != d[0]:
        notes.append("  !! newest dates differ between the two sheets")
    wk = [str(x) for x in d[:n] if x.weekday() >= 5]
    notes.append("Weekend dates in newest %d Indices rows: %s" % (n, wk or "none"))
    im = {from_serial(v[0]): v[1:] for r, v in book.idx}
    wrong = []
    for r, v in book.spr:
        dd = from_serial(v[0])
        x = im.get(dd)
        if not x or None in x:
            continue
        t, a, b, c = x
        exp = [a - b, a - c, b - c, t - a, t - b, t - c]
        if any(abs(v[1 + i] - exp[i]) > 1e-9 for i in range(6) if v[1 + i] is not None):
            wrong.append(str(dd))
    notes.append("Spreads rows that do not equal the Indices rows for the same date: %d %s" % (len(wrong), wrong[:12]))
    gaps = [(str(d[i + 1]), str(d[i])) for i in range(min(n, len(d) - 1)) if (d[i] - d[i + 1]).days > 4]
    notes.append("Calendar gaps over 4 days in newest %d rows: %s" % (n, gaps or "none"))
    return notes
 
 
def read_values(path):
    rows = []
    with open(path, newline="") as f:
        rd = csv.reader(f)
        hdr = [h.strip() for h in next(rd)]
        if hdr != ["date", "US10YR", "AAA", "BBB", "CCC"]:
            stop("values.csv header must be date,US10YR,AAA,BBB,CCC (got %s)" % hdr)
        seen = set()
        for ln, r in enumerate(rd, 2):
            if not r or not "".join(r).strip():
                continue
            if len(r) != 5:
                stop("values.csv line %d: expected 5 fields" % ln)
            try:
                d = dt.date.fromisoformat(r[0].strip())
            except ValueError:
                stop("values.csv line %d: bad date %r" % (ln, r[0]))
            if d in seen:
                stop("values.csv: duplicate date %s" % d)
            seen.add(d)
            vals = []
            for k, s in zip(NAMES, r[1:]):
                s = s.strip()
                if s in ("", "."):
                    vals.append(None)
                    continue
                try:
                    x = float(s)
                except ValueError:
                    stop("values.csv line %d: %s is not a number: %r" % (ln, k, s))
                lo, hi = RANGE[k]
                if not lo < x < hi:
                    stop("values.csv line %d: %s=%s outside %s-%s" % (ln, k, x, lo, hi))
                vals.append(x)
            rows.append((d, vals))
    rows.sort()
    return rows
 
 
def plan(book, rows, warnings):
    latest = from_serial(book.idx[0][1][0])
    if from_serial(book.spr[0][1][0]) != latest:
        stop("Newest date differs between Spreads (%s) and Indices (%s)" % (from_serial(book.spr[0][1][0]), latest))
    wb = {from_serial(v[0]): v[1:] for r, v in book.idx[:60]}
    byd = dict(rows)
    # overlap check: workbook's latest date must be in the file and equal
    if latest not in byd:
        stop("values.csv must include the workbook's latest date %s (overlap check)" % latest)
    for d in sorted(byd):
        if d in wb and d >= from_serial(book.idx[min(4, len(book.idx) - 1)][1][0]):
            for k, a, b in zip(NAMES, wb[d], byd[d]):
                if b is None or a is None or abs(a - b) > 1e-9:
                    stop("Overlap mismatch on %s %s: workbook %s, values.csv %s. The workbook may hold other values than FRED "
                         "(revision or a past paste error). Nothing written; ask the owner." % (d, k, a, b))
    older = [str(d) for d in sorted(byd) if d in wb and d < from_serial(book.idx[min(4, len(book.idx) - 1)][1][0]) and
             any(b is None or a is None or abs(a - b) > 1e-9 for a, b in zip(wb[d], byd[d]))]
    if older:
        warnings.append("Older overlap dates differ from FRED (reported, not changed): %s" % ", ".join(older))
    new = [(d, v) for d, v in rows if d > latest]
    complete, pending, skipped = [], [], []
    last_complete = max([d for d, v in new if None not in v], default=None)
    for d, v in new:
        if d.weekday() >= 5:
            stop("Weekend date %s in values.csv" % d)
        if None not in v:
            complete.append((d, v))
        elif last_complete is None or d > last_complete:
            pending.append((d, v))                      # newest dates FRED has not finished yet
        elif v[0] is None and None not in v[1:]:
            skipped.append((d, v))                      # US holiday: Treasury blank, ICE has a value (owner leaves these out)
        else:
            stop("Date %s has missing values in the middle of the file: %s" % (d, v))
    # day-to-day sanity
    prev = wb[latest]
    for d, v in complete:
        for k, a, b in zip(NAMES, prev, v):
            if abs(a - b) > MAXMOVE[k]:
                stop("%s moved %.2f in one day on %s (limit %.1f). Check the source." % (k, b - a, d, MAXMOVE[k]))
        if not (v[1] <= v[2] <= v[3]):
            warnings.append("%s: AAA/BBB/CCC not in order (%s)" % (d, v[1:]))
        prev = v
    return complete, pending, skipped
 
 
def build(book, complete, out):
    n = len(complete)
    newest_first = sorted(complete, reverse=True)
    ix_rows, sp_rows = [], []
    for i, (d, v) in enumerate(newest_first):
        r = FIRST + i
        t, a, b, c = v
        ix_rows.append('<row r="%d" spans="1:6" x14ac:dyDescent="0.25"><c r="B%d" s="8"><v>%d</v></c>'
                       '<c r="C%d" s="1"><v>%r</v></c><c r="D%d" s="1"><v>%r</v></c><c r="E%d" s="1"><v>%r</v></c>'
                       '<c r="F%d" s="1"><v>%r</v></c></row>' % (r, r, serial(d), r, t, r, a, r, b, r, c))
        sp = [a - b, a - c, b - c, t - a, t - b, t - c]
        cells = "".join('<c r="%s%d" s="7"><v>%r</v></c>' % (col, r, x) for col, x in zip("BCDEFG", sp))
        sp_rows.append('<row r="%d" spans="1:7" x14ac:dyDescent="0.25"><c r="A%d" s="8"><v>%d</v></c>%s</row>' % (r, r, serial(d), cells))
    parts = dict(book.parts)
    # Indices: insert, shift the rest down
    head, body, tail = split_sheet(book.parts[INDICES].decode())
    rows = rows_of(body)
    keep = "".join(x for r, x in rows if r < FIRST)
    shifted = "".join(shift_row(x, n) for r, x in rows if r >= FIRST)
    last = rows[-1][0] + n
    head = re.sub(r'<dimension ref="A1:F\d+"/>', '<dimension ref="A1:F%d"/>' % last, head)
    parts[INDICES] = (head + keep + "".join(ix_rows) + shifted + tail).encode()
    # Spreads: insert, shift down, drop the n trailing blank rows so the sheet keeps its size
    head, body, tail = split_sheet(book.parts[SPREADS].decode())
    rows = rows_of(body)
    tailrows = rows[-n:]
    if any("<v>" in x for r, x in tailrows):
        stop("Spreads sheet has data in its last %d rows; cannot insert" % n)
    rows = rows[:-n]
    keep = "".join(x for r, x in rows if r < FIRST)
    shifted = "".join(shift_row(x, n) for r, x in rows if r >= FIRST)
    parts[SPREADS] = (head + keep + "".join(sp_rows) + shifted + tail).encode()
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in book.order:
            zo.writestr(name, parts[name])
    os.replace(tmp, out)
 
 
def prove(old, new, complete):
    """Re-read both files; every old row must still be there, shifted by n, and the new rows must be right."""
    n = len(complete)
    a, b = Book(old), Book(new)
    if b.idx[n:] != [(r + n, v) for r, v in a.idx]:
        stop("PROOF FAILED: old Indices rows changed")
    if b.spr[n:] != [(r + n, v) for r, v in a.spr]:
        stop("PROOF FAILED: old Spreads rows changed")
    exp = sorted(complete, reverse=True)
    for i, (d, v) in enumerate(exp):
        r, row = b.idx[i]
        if r != FIRST + i or from_serial(row[0]) != d or [round(x, 9) for x in row[1:]] != [round(x, 9) for x in v]:
            stop("PROOF FAILED: new Indices row %d" % r)
        r2, row2 = b.spr[i]
        t, aa, bb, c = v
        e = [aa - bb, aa - c, bb - c, t - aa, t - bb, t - c]
        if r2 != r or from_serial(row2[0]) != d or any(abs(x - y) > 1e-9 for x, y in zip(row2[1:], e)):
            stop("PROOF FAILED: new Spreads row %d" % r2)
    za, zb = zipfile.ZipFile(old), zipfile.ZipFile(new)
    diff = [x for x in za.namelist() if za.read(x) != zb.read(x)]
    if sorted(diff) != sorted([SPREADS, INDICES]):
        stop("PROOF FAILED: unexpected parts changed: %s" % diff)
    if za.namelist() != zb.namelist():
        stop("PROOF FAILED: part list changed")
    return len(b.idx), len(b.spr)
 
 
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("values", nargs="?")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--out")
    g.add_argument("--in-place", action="store_true")
    ap.add_argument("--audit", action="store_true")
    a = ap.parse_args()
    try:
        book = Book(a.workbook)
        if a.audit:
            print("\n".join(audit(book)))
            return 0
        if not a.values or not (a.out or a.in_place):
            stop("Give values.csv and --out FILE or --in-place")
        warnings = []
        complete, pending, skipped = plan(book, read_values(a.values), warnings)
        print("Workbook newest date:", from_serial(book.idx[0][1][0]))
        if skipped:
            print("Skipped (Treasury market closed, 10Y blank):", ", ".join(str(d) for d, v in skipped))
        if pending:
            print("Not added yet (FRED has not published every series):",
                  "; ".join("%s %s" % (d, ["" if x is None else x for x in v]) for d, v in pending))
        if not complete:
            print("NOTHING TO ADD. No new date has all four values yet.")
            return 0
        print("New rows (%d): %s to %s" % (len(complete), complete[0][0], complete[-1][0]))
        out = a.out or (a.workbook + ".new.tmp")
        build(book, complete, out)
        ni, ns = prove(a.workbook, out, complete)
        print("Proof passed: old rows unchanged, %d new rows on each sheet, only the 2 data sheets changed, 20 charts untouched." % len(complete))
        print("Indices rows now %d, Spreads rows now %d" % (ni, ns))
        if a.in_place:
            bdir = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), "backups")
            os.makedirs(bdir, exist_ok=True)
            stem = os.path.splitext(os.path.basename(a.workbook))[0]
            bak = os.path.join(bdir, "%s_before_%s.xlsx" % (stem, complete[0][0].isoformat()))
            if os.path.exists(bak):
                stop("Backup %s already exists" % bak)
            shutil.copy2(a.workbook, bak)
            if open(bak, "rb").read() != open(a.workbook, "rb").read():
                stop("Backup check failed")
            os.replace(out, a.workbook)
            print("Backup written:", bak)
        for w in warnings:
            print("NOTE:", w)
        print("--- audit of the existing data (reported, not changed) ---")
        print("\n".join(audit(book)))
        return 0
    except Stop as e:
        print("STOP:", e)
        return 2
 
 
if __name__ == "__main__":
    sys.exit(main())