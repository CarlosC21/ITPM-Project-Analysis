#!/usr/bin/env python3
"""Update 'Bull_Bear_Markets.xlsx' (sheet 'S&P 500') with new daily index bars.
 
Usage:
  python bb_update_workbook.py WORKBOOK values.csv --out NEW.xlsx
  python bb_update_workbook.py WORKBOOK values.csv --in-place        (backup first)
  python bb_update_workbook.py WORKBOOK --audit                      (read-only checks)
  add --align-chart to set the 4 chart series to start at row 2 (default: keep the starts, extend the ends)
 
values.csv: header  date,open,high,low,close   (YYYY-MM-DD, one row per COMPLETED trading day).
It must include the workbook's newest date (and ideally a few earlier ones) as an overlap check.
 
Standard library only. Edits the sheet XML directly, so the chart and all formatting are kept.
New rows go in at row 2 (newest on top); every old row moves down by the number of new rows
(formulas, shared formulas, the Covid/GFC block in M:O and the dimension move with them).
Exit code 2 = STOP (nothing written).
"""
import argparse, csv, datetime as dt, html, os, re, shutil, sys, zipfile
 
SHEET = "xl/worksheets/sheet2.xml"      # 'S&P 500'
CHART = "xl/charts/chart1.xml"
EPOCH = dt.date(1899, 12, 30)
HEADERS = ["S&P 500", "Open", "High", "Low", "Close", "Rolling Index High", "Rolling Bear Market Level",
           "Bull Market Index", "Bear Market Index", "Labels"]
MAXMOVE = 0.12          # largest plausible close-to-close move (12%); 1987-10-19 was -20%, so only for new data
OHLC_TOL = 1e-6
 
 
class Stop(Exception):
    pass
 
 
def stop(msg):
    raise Stop(msg)
 
 
def serial(d):
    return (d - EPOCH).days
 
 
def from_serial(n):
    return EPOCH + dt.timedelta(days=int(n))
 
 
ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
 
 
def split_sheet(xml):
    a = xml.index("<sheetData>") + len("<sheetData>")
    b = xml.index("</sheetData>")
    return xml[:a], xml[a:b], xml[b:]
 
 
def rows_of(body):
    return [(int(m.group(1)), m.group(0)) for m in ROW_RE.finditer(body)]
 
 
def cells(rowxml):
    out = {}
    for m in CELL_RE.finditer(rowxml):
        col, attrs, inner = m.group(1), m.group(3), m.group(4) or ""
        v = re.search(r"<v>([^<]*)</v>", inner)
        f = re.search(r"<f[^>]*?(?:/>|>([^<]*)</f>)", inner)
        out[col] = {"attrs": attrs, "v": None if v is None else v.group(1), "f": None if f is None else (f.group(1) or ""),
                    "formula": f is not None}
    return out
 
 
REF_RE = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")
 
 
def shift_refs(text, n):
    return REF_RE.sub(lambda m: m.group(1) + str(int(m.group(2)) + n), text)
 
 
def shift_row(rowxml, n):
    rowxml = re.sub(r'<row r="(\d+)"', lambda m: '<row r="%d"' % (int(m.group(1)) + n), rowxml, count=1)
    rowxml = re.sub(r'<c r="([A-Z]+)(\d+)"', lambda m: '<c r="%s%d"' % (m.group(1), int(m.group(2)) + n), rowxml)
 
    def fix(m):
        attrs = re.sub(r'ref="([^"]+)"', lambda a: 'ref="%s"' % shift_refs(a.group(1), n), m.group(1))
        if m.group(2) is None:                       # <f .../>
            return "<f%s/>" % attrs
        return "<f%s>%s</f>" % (attrs, shift_refs(m.group(2), n))
    return re.sub(r"<f([^>]*?)(?:/>|>([^<]*)</f>)", fix, rowxml)
 
 
def shared_strings(z):
    s = z.read("xl/sharedStrings.xml").decode("utf-8")
    return [html.unescape(re.sub(r"<[^>]+>", "", x)) for x in re.findall(r"<si>(.*?)</si>", s, re.S)]
 
 
class Book:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.parts = {i.filename: self.z.read(i.filename) for i in self.z.infolist()}
        self.order = [i.filename for i in self.z.infolist()]
        wb = self.parts["xl/workbook.xml"].decode()
        names = re.findall(r'<sheet name="([^"]+)"', wb)
        if names != ["Definitions", "S&amp;P 500"]:
            stop("Unexpected sheets %s" % names)
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode()
        if not re.search(r'Id="rId2"[^>]*Target="worksheets/sheet2.xml"|Target="worksheets/sheet2.xml"[^>]*Id="rId2"', rels):
            stop("Sheet mapping changed")
        if not [n for n in self.parts if re.match(r"xl/charts/chart\d+\.xml$", n)] == [CHART]:
            stop("Expected exactly one chart (chart1.xml)")
        ss = shared_strings(self.z)
        self.head, body, self.tail = split_sheet(self.parts[SHEET].decode())
        self.rows = rows_of(body)
        self.row_cells = {r: cells(x) for r, x in self.rows}
        if 1 not in self.row_cells:
            stop("Header row missing")
        hdr = []
        for col in "ABCDEFGHIJ":
            c = self.row_cells[1].get(col)
            hdr.append(ss[int(c["v"])] if c and c["v"] is not None and 't="s"' in c["attrs"] else None)
        if hdr != HEADERS:
            stop("Row 1 headers changed: %s" % hdr)
        self.data = []                       # (row, date, o, h, l, c) newest first
        for r, _ in self.rows:
            if r < 2:
                continue
            cc = self.row_cells[r]
            if "A" not in cc or cc["A"]["v"] is None:
                continue
            vals = [cc[k]["v"] for k in "BCDE"]
            if any(v is None for v in vals):
                continue
            self.data.append((r, from_serial(cc["A"]["v"]), *[float(v) for v in vals]))
        if not self.data or self.data[0][0] != 2:
            stop("Data must start at row 2")
        f2 = self.row_cells[2]["F"]["f"] or ""
        m = re.fullmatch(r"MAX\(C2:\$C\$(\d+)\)", f2)
        if not m:
            stop("F2 formula changed: %r" % f2)
        self.anchor = int(m.group(1))        # row where the rolling-high formulas turn around (1931-09-03)
        last = self.data[-1][0]
        if last < self.anchor:
            stop("Anchor row beyond the data")
        if self.row_cells[2]["G"]["f"] != "F2*0.8":
            stop("G2 formula changed")
        self.series = re.findall(r"<c:ser>.*?</c:ser>", self.parts[CHART].decode("utf-8"), re.S)
        if len(self.series) != 3:
            stop("Chart should have 3 series")
        self.chart_refs = [re.findall(r"<c:f>([^<]*)</c:f>", s) for s in self.series]
 
 
def audit(book):
    d = book.data
    notes = ["Data rows %d (rows 2-%d), newest %s, oldest %s; rolling-high anchor row %d (%s)"
             % (len(d), d[-1][0], d[0][1], d[-1][1], book.anchor, [x[1] for x in d if x[0] == book.anchor][0])]
    order = [(x[0], str(x[1])) for x, y in zip(d, d[1:]) if not x[1] > y[1]]
    notes.append("Rows where the date is not older than the row below: %s" % (order[:6] or "none"))
    seen, dup = {}, []
    for x in d:
        seen.setdefault(x[1], []).append(x[0])
    dup = {k: v for k, v in seen.items() if len(v) > 1}
    if dup:
        ks = sorted(dup)
        notes.append("Duplicate dates: %d (%s to %s), rows %d-%d; values identical: %s"
                     % (len(dup), ks[0], ks[-1], min(min(v) for v in dup.values()), max(max(v) for v in dup.values()),
                        all(len({tuple(d[r - 2][2:]) for r in v}) == 1 for v in dup.values())))
    else:
        notes.append("Duplicate dates: none")
    bad = [x[0] for x in d[:3000] if not (x[4] <= min(x[2], x[5]) + OHLC_TOL and x[3] >= max(x[2], x[5]) - OHLC_TOL)]
    notes.append("OHLC inconsistent (low > open/close or high < open/close), newest 3000 rows: rows %s" % (bad or "none"))
    gaps = [(str(y[1]), str(x[1])) for x, y in zip(d, d[1:]) if (x[1] - y[1]).days > 7 and x[1].year >= 1990]
    notes.append("Calendar gaps over 7 days since 1990: %s" % (gaps or "none"))
    wk = [str(x[1]) for x in d if x[1].weekday() >= 5]
    notes.append("Weekend dates: %s" % (wk or "none"))
    # formula pattern
    pats = {}
    for r, cc in book.row_cells.items():
        if r >= 2 and "F" in cc and cc["F"]["f"]:
            k = "above" if r <= book.anchor else "below"
            pats[k] = pats.get(k, 0) + 1
    notes.append("Rolling-high formulas: %s rows use MAX(C<row>:$C$%d) (all-time high to date); rows below the anchor use "
                 "MAX(C$%d:$C<row>), which looks forward to 1931-09-03 (1928-1931 only)"
                 % (pats.get("above", 0), book.anchor, book.anchor))
    notes.append("Chart series ranges: %s" % " | ".join(", ".join(x.replace("&amp;", "&") for x in refs) for refs in book.chart_refs))
    return notes
 
 
def read_values(path):
    rows = []
    with open(path, newline="") as f:
        rd = csv.reader(f)
        hdr = [h.strip().lower() for h in next(rd)]
        if hdr != ["date", "open", "high", "low", "close"]:
            stop("values.csv header must be date,open,high,low,close (got %s)" % hdr)
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
            try:
                o, h, l, c = [float(x) for x in r[1:]]
            except ValueError:
                stop("values.csv line %d: a price is not a number: %s" % (ln, r[1:]))
            if not all(0 < x < 100000 for x in (o, h, l, c)):
                stop("values.csv line %d: price outside 0-100000" % ln)
            if not (l <= min(o, c) + OHLC_TOL and h >= max(o, c) - OHLC_TOL):
                stop("values.csv line %d (%s): low/high do not bracket open and close: %s" % (ln, d, r[1:]))
            rows.append((d, o, h, l, c))
    rows.sort()
    return rows
 
 
def plan(book, rows, warnings):
    latest = book.data[0][1]
    byd = {d: v for d, *v in rows}
    if latest not in byd:
        stop("values.csv must include the workbook's newest date %s (overlap check)" % latest)
    today = dt.date.today()
    for d, o, h, l, c in rows:
        if d.weekday() >= 5:
            stop("Weekend date %s in values.csv" % d)
        if d > today:
            stop("Future date %s in values.csv" % d)
    wb = {x[1]: x[2:] for x in book.data[:10]}
    for d in sorted(byd):
        if d in wb:
            if any(abs(a - b) > 1e-9 for a, b in zip(wb[d], byd[d])):
                stop("Overlap mismatch on %s: workbook %s, values.csv %s. History is never changed silently; ask the owner."
                     % (d, wb[d], byd[d]))
    new = [(d, *v) for d, *v in rows if d > latest]
    prev = book.data[0][5]
    for i, (d, o, h, l, c) in enumerate(new):
        gap = (d - (new[i - 1][0] if i else latest)).days
        if gap > 8:
            stop("Gap of %d days before %s: days are missing (do not add past a hole)" % (gap, d))
        if gap > 4:
            warnings.append("Gap of %d calendar days before %s (holiday or missing days?)" % (gap, d))
        if abs(c / prev - 1) > MAXMOVE:
            stop("Close moved %.1f%% on %s (limit %d%%). Check the source." % ((c / prev - 1) * 100, d, MAXMOVE * 100))
        prev = c
    return new
 
 
def chart_new(book, n, align):
    """New <c:f> text for the 3 series (each has tx, cat, val refs). Ends +n; starts kept (or row 2 with --align-chart)."""
    x = book.parts[CHART].decode("utf-8")
    pat = re.compile(r"(<c:f>'S&amp;P 500'!\$([A-Z])\$)(\d+):(\$[A-Z]\$)(\d+)(</c:f>)")
    if len(pat.findall(x)) != 6:
        stop("Chart range references changed (expected 6 ranges, found %d)" % len(pat.findall(x)))
 
    def fix(m):
        start = 2 if align else int(m.group(3))
        return "%s%d:%s%d%s" % (m.group(1), start, m.group(4), int(m.group(5)) + n, m.group(6))
    return pat.sub(fix, x)
 
 
def build(book, new, out, align):
    n = len(new)
    anchor_new = book.anchor + n
    base_max = max(x[3] for x in book.data if x[0] <= book.anchor)          # highest High, 1931-09-03 to the old newest day
    asc = sorted(new)
    fvals, run = {}, base_max
    for d, o, h, l, c in asc:
        run = max(run, h)
        fvals[d] = run
    # but F for a day = max over that day and everything older, so the day's value must include older new days only
    # (asc order gives exactly that)
    r2 = book.row_cells[2]
    sA = re.search(r's="(\d+)"', r2["A"]["attrs"]).group(1)
    sB = re.search(r's="(\d+)"', r2["B"]["attrs"]).group(1)
    row2xml = [x for r, x in book.rows if r == 2][0]
    rowattrs = re.match(r'<row r="2"([^>]*)>', row2xml).group(1)
    newrows = []
    for i, (d, o, h, l, c) in enumerate(sorted(new, reverse=True)):
        r = 2 + i
        F = fvals[d]
        G = F * 0.8
        bull = c > G
        cells_ = ['<c r="A%d" s="%s"><v>%d</v></c>' % (r, sA, serial(d))]
        for col, v in zip("BCDE", (o, h, l, c)):
            cells_.append('<c r="%s%d" s="%s"><v>%r</v></c>' % (col, r, sB, v))
        cells_.append('<c r="F%d" s="%s"><f>MAX(C%d:$C$%d)</f><v>%r</v></c>' % (r, sB, r, anchor_new, F))
        cells_.append('<c r="G%d" s="%s"><f>F%d*0.8</f><v>%r</v></c>' % (r, sB, r, G))
        if bull:
            cells_.append('<c r="H%d"><f>IF(E%d&gt;G%d,E%d,#N/A)</f><v>%r</v></c>' % (r, r, r, r, c))
            cells_.append('<c r="I%d" t="e"><f>IF(E%d&lt;=G%d,E%d,#N/A)</f><v>#N/A</v></c>' % (r, r, r, r))
        else:
            cells_.append('<c r="H%d" t="e"><f>IF(E%d&gt;G%d,E%d,#N/A)</f><v>#N/A</v></c>' % (r, r, r, r))
            cells_.append('<c r="I%d"><f>IF(E%d&lt;=G%d,E%d,#N/A)</f><v>%r</v></c>' % (r, r, r, r, c))
        cells_.append('<c r="J%d"/>' % r)
        newrows.append('<row r="%d"%s>%s</row>' % (r, rowattrs, "".join(cells_)))
    keep = "".join(x for r, x in book.rows if r < 2)
    shifted = "".join(shift_row(x, n) for r, x in book.rows if r >= 2)
    last = book.rows[-1][0] + n
    head = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: '<dimension ref="A1:%s%d"/>' % (m.group(1), last), book.head)
    parts = dict(book.parts)
    parts[SHEET] = (head + keep + "".join(newrows) + shifted + book.tail).encode("utf-8")
    parts[CHART] = chart_new(book, n, align).encode("utf-8")
    # drop calcChain, force a full recalculation when Excel opens the file
    order = [x for x in book.order if x != "xl/calcChain.xml"]
    parts.pop("xl/calcChain.xml", None)
    ct = parts["[Content_Types].xml"].decode()
    parts["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct).encode()
    rl = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship Id="[^"]*" Type="[^"]*calcChain" Target="calcChain.xml"/>', "", rl).encode()
    wbx = parts["xl/workbook.xml"].decode()
    if "fullCalcOnLoad" not in wbx:
        wbx = re.sub(r"<calcPr ([^/]*)/>", r'<calcPr \1 fullCalcOnLoad="1"/>', wbx)
    parts["xl/workbook.xml"] = wbx.encode()
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in order:
            zo.writestr(name, parts[name])
    os.replace(tmp, out)
 
 
def u_not_ok(za, zb):
    return set(za.namelist()) - set(zb.namelist()) - {"xl/calcChain.xml"} or set(zb.namelist()) - set(za.namelist())
 
 
def prove(old, new_path, new, align):
    n = len(new)
    a, b = Book(old), Book(new_path)
    if len(b.data) != len(a.data) + n:
        stop("PROOF FAILED: row count")
    if [(r + n, d, *v) for r, d, *v in a.data] != [tuple(x) for x in b.data[n:]]:
        stop("PROOF FAILED: old rows changed or moved wrongly")
    exp = sorted(new, reverse=True)
    for i, (d, o, h, l, c) in enumerate(exp):
        if tuple(b.data[i]) != (2 + i, d, o, h, l, c):
            stop("PROOF FAILED: new row %d" % (2 + i))
    if b.anchor != a.anchor + n:
        stop("PROOF FAILED: anchor row")
    # every old row: formulas equal the old ones shifted by n
    for r, cc in a.row_cells.items():
        if r < 2:
            continue
        for col, c in cc.items():
            if c["formula"]:
                nc = b.row_cells[r + n].get(col)
                if nc is None or not nc["formula"] or (nc["f"] or "") != shift_refs(c["f"] or "", n) or nc["v"] != c["v"]:
                    stop("PROOF FAILED: formula %s%d" % (col, r))
    # rolling high / level / bull / bear for every new row, recomputed from the stored data
    hi = max(x[3] for x in b.data if x[0] <= b.anchor and x[0] > n + 1)
    mx = hi
    for x in sorted(b.data[:n], key=lambda t: t[1]):
        mx = max(mx, x[3])
        cc = b.row_cells[x[0]]
        if abs(float(cc["F"]["v"]) - mx) > 1e-9 or abs(float(cc["G"]["v"]) - mx * 0.8) > 1e-9:
            stop("PROOF FAILED: F/G cache row %d" % x[0])
        if (x[5] > mx * 0.8) != (cc["H"]["v"] != "#N/A") or (x[5] <= mx * 0.8) != (cc["I"]["v"] != "#N/A"):
            stop("PROOF FAILED: bull/bear row %d" % x[0])
    za, zb = zipfile.ZipFile(old), zipfile.ZipFile(new_path)
    changed = sorted(x for x in zb.namelist() if x in za.namelist() and za.read(x) != zb.read(x))
    want = sorted([SHEET, CHART, "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/workbook.xml"])
    if not set(changed) <= set(want) or u_not_ok(za, zb):
        stop("PROOF FAILED: unexpected parts changed: %s" % changed)
    return len(b.data)
 
 
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("values", nargs="?")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--out")
    g.add_argument("--in-place", action="store_true")
    ap.add_argument("--audit", action="store_true")
    ap.add_argument("--align-chart", action="store_true")
    a = ap.parse_args()
    try:
        book = Book(a.workbook)
        if a.audit:
            print("\n".join(audit(book)))
            return 0
        if not a.values or not (a.out or a.in_place):
            stop("Give values.csv and --out FILE or --in-place")
        warnings = []
        new = plan(book, read_values(a.values), warnings)
        print("Workbook newest date:", book.data[0][1])
        if not new:
            print("NOTHING TO ADD. The workbook is already up to date with values.csv.")
            return 0
        print("New rows (%d): %s to %s" % (len(new), new[0][0], new[-1][0]))
        out = a.out or (a.workbook + ".new.tmp")
        build(book, new, out, a.align_chart)
        total = prove(a.workbook, out, new, a.align_chart)
        print("Proof passed: %d old rows unchanged (moved down %d), formulas shifted, new F/G/H/I checked, "
              "only sheet, chart, workbook.xml, rels, content types changed (calcChain removed)." % (total - len(new), len(new)))
        if a.in_place:
            bdir = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), "backups")
            os.makedirs(bdir, exist_ok=True)
            stem = os.path.splitext(os.path.basename(a.workbook))[0]
            bak = os.path.join(bdir, "%s_before_%s.xlsx" % (stem, new[0][0].isoformat()))
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