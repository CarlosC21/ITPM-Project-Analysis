#!/usr/bin/env python3
"""US Industrial Production workbook updater.
 
Edits the .xlsx at file (XML) level so the 42 charts and all formatting survive.
Standard library only.
 
  python ip_update_workbook.py US_Industrial_Production.xlsx values.csv --out updated.xlsx
  python ip_update_workbook.py US_Industrial_Production.xlsx values.csv --in-place
 
Any validation failure prints "STOP: ..." and writes nothing.
See ip_conventions.md for layout, rules and the values.csv format.
"""
import argparse
import csv
import datetime as dt
import hashlib
import os
import posixpath
import re
import shutil
import sys
import tempfile
import xml.dom.minidom
import zipfile
 
# sheet name, FRED series id (same order as the sheets in the workbook)
SHEETS = [
    ("Industrial_Production", "INDPRO"),
    ("Manufacturing", "IPMAN"),
    ("Food_Bev_Tob", "IPG311A2S"),
    ("Textiles", "IPG313A4S"),
    ("Chemicals", "IPG325S"),
    ("Plastics_Rubber", "IPG326S"),
    ("Apparel_Leather", "IPG315A6S"),
    ("Wood", "IPG321S"),
    ("Paper", "IPG322S"),
    ("Printing_Related", "IPG323S"),
    ("Petrol_Coal", "IPG324S"),
    ("Nonmetallic_Minerals", "IPG327S"),
    ("Primary_Metals", "IPG331S"),
    ("Fabricated_Metals", "IPG332S"),
    ("Machinery", "IPG333S"),
    ("Comp_Elec", "IPG334S"),
    ("ElecEquip_Appliances_Components", "IPG335S"),
    ("MotorVehicles_Parts", "IPG3361T3S"),
    ("Aero_MiscTransport", "IPG3364T9N"),
    ("Furniture_Related", "IPG337S"),
    ("Misc", "IPG339S"),
]
N_CHARTS = 42
WINDOW = 5            # months in values.csv: the new month and up to 4 prior months
SAME = 5e-5           # index values closer than this are "the same" (FRED shows 4 decimals)
MAX_REVISION = 0.05   # largest plausible revision of a prior month, as a fraction
MAX_MOM = 0.20        # largest plausible monthly move of the new month
EPS = 1e-6
 
 
class Stop(Exception):
    pass
 
 
# ---------------------------------------------------------------- helpers
def ym_add(ym, k):
    t = ym[0] * 12 + (ym[1] - 1) + k
    return (t // 12, t % 12 + 1)
 
 
def ym_str(ym):
    return "%d-%02d" % ym
 
 
def ym_label(ym):
    return dt.date(ym[0], ym[1], 1).strftime("%b-%y")
 
 
def parse_ym(s):
    m = re.fullmatch(r"\s*(\d{4})-(\d{2})\s*", s)
    if not m or not 1 <= int(m.group(2)) <= 12:
        raise Stop("bad month %r in values.csv (use YYYY-MM)" % s)
    return (int(m.group(1)), int(m.group(2)))
 
 
def serial(ym):
    return (dt.date(ym[0], ym[1], 1) - dt.date(1899, 12, 30)).days
 
 
def from_serial(v):
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(round(v)))
    return (d.year, d.month), d.day
 
 
def fmt(x):
    x = float(x)
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(x)
 
 
def attrs_of(s):
    return dict(re.findall(r'([\w:]+)="([^"]*)"', s))
 
 
def col_idx(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n
 
 
def col_name(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s
 
 
# ---------------------------------------------------------------- sheet model
ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
 
 
class Sheet:
    def __init__(self, name, xml_text):
        self.name = name
        m = re.match(r"(.*<sheetData>)(.*)(</sheetData>.*)", xml_text, re.S)
        if not m:
            raise Stop("%s: cannot read sheet XML" % name)
        self.head, body, self.tail = m.group(1), m.group(2), m.group(3)
        self.rows = {}
        for rm in ROW_RE.finditer(body):
            n = int(rm.group(1))
            cells = {}
            for cm in CELL_RE.finditer(rm.group(3) or ""):
                cells[cm.group(1)] = [cm.group(3), cm.group(4) or ""]
            self.rows[n] = {"attrs": rm.group(2), "cells": cells}
 
    def num(self, n, col):
        c = self.rows.get(n, {}).get("cells", {}).get(col)
        if not c or "t=" in c[0]:
            return None
        m = re.search(r"<v>(.*?)</v>", c[1])
        return float(m.group(1)) if m else None
 
    def last_row(self):
        return max(self.rows)
 
    def data_rows(self):
        return sorted(n for n in self.rows if n >= 2)
 
    def month(self, n):
        v = self.num(n, "A")
        if v is None:
            raise Stop("%s row %d: month cell is not a date" % (self.name, n))
        return from_serial(v)[0]
 
    def day(self, n):
        return from_serial(self.num(n, "A"))[1]
 
    def month_map(self):
        out = {}
        for n in self.data_rows():
            ym = self.month(n)
            if ym in out:
                raise Stop("%s: month %s appears twice" % (self.name, ym_label(ym)))
            out[ym] = n
        return out
 
    def set_num(self, n, col, val):
        c = self.rows[n]["cells"].setdefault(col, ["", ""])
        c[0] = re.sub(r'\s+t="[^"]*"', "", c[0])
        c[1] = "" if val is None else "<v>%s</v>" % fmt(val)
 
    def insert_top(self):
        """Insert an empty row at row 2 with the styles of the current row 2."""
        shifted = {}
        for n, r in self.rows.items():
            shifted[n + 1 if n >= 2 else n] = r
        tpl = shifted[3]
        cells = {}
        for col, (a, _) in tpl["cells"].items():
            if col in "ABCD":
                cells[col] = [re.sub(r'\s+t="[^"]*"', "", a), ""]
        shifted[2] = {"attrs": tpl["attrs"], "cells": cells}
        self.rows = shifted
 
    def serialize(self):
        out = []
        maxcol = 4
        for n in sorted(self.rows):
            r = self.rows[n]
            attrs = r["attrs"]
            if any(col_idx(c) > 4 for c in r["cells"]):
                maxcol = max(maxcol, max(col_idx(c) for c in r["cells"]))
            cs = []
            for col in sorted(r["cells"], key=col_idx):
                a, inner = r["cells"][col]
                cs.append('<c r="%s%d"%s>%s</c>' % (col, n, a, inner) if inner
                          else '<c r="%s%d"%s/>' % (col, n, a))
            out.append('<row r="%d"%s>%s</row>' % (n, attrs, "".join(cs)))
        head = re.sub(r'<dimension ref="[^"]*"/>',
                      '<dimension ref="A1:%s%d"/>' % (col_name(maxcol), self.last_row()), self.head)
        return head + "".join(out) + self.tail
 
    def expected(self, n, mm):
        ym = self.month(n)
        lv = self.num(n, "B")
        mom = yoy = None
        if lv is None:
            return None, None
        for off, which in ((-1, "m"), (-12, "y")):
            r = mm.get(ym_add(ym, off))
            p = self.num(r, "B") if r else None
            if p is None:
                continue
            if which == "m":
                mom = lv / p - 1
            else:
                yoy = lv / p - 1
        return mom, yoy
 
 
# ---------------------------------------------------------------- workbook (zip) model
class Book:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid = {}
        for m in re.finditer(r"<Relationship ([^>]*?)/>", rels):
            a = attrs_of(m.group(1))
            rid[a["Id"]] = posixpath.normpath(posixpath.join("xl", a["Target"]))
        self.sheets = {}
        self.order = []
        for m in re.finditer(r"<sheet ([^>]*?)/>", wb):
            a = attrs_of(m.group(1))
            p = rid[a["r:id"]]
            self.order.append((a["name"], p))
            self.sheets[a["name"]] = Sheet(a["name"], self.parts[p].decode("utf-8"))
        names = [n for n, _ in SHEETS]
        if [n for n, _ in self.order] != names:
            raise Stop("sheet names/order differ from the expected 21: %s" % [n for n, _ in self.order])
 
    def n_charts(self):
        return len([p for p in self.parts if re.fullmatch(r"xl/charts/chart\d+\.xml", p)])
 
    def sheet_paths(self):
        return {p for _, p in self.order}
 
    def render(self):
        parts = dict(self.parts)
        for name, path in self.order:
            parts[path] = self.sheets[name].serialize().encode("utf-8")
        for p, b in parts.items():
            if p.endswith(".xml") or p.endswith(".rels"):
                try:
                    xml.dom.minidom.parseString(b)
                except Exception as e:
                    raise Stop("internal error: %s is not well-formed XML (%s)" % (p, e))
        return parts
 
    def write(self, path, parts):
        tmp = tempfile.NamedTemporaryFile(delete=False, dir=os.path.dirname(os.path.abspath(path)),
                                          suffix=".tmp")
        tmp.close()
        with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
            for i in self.infos:
                zi = zipfile.ZipInfo(i.filename, i.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = i.external_attr
                z.writestr(zi, parts[i.filename])
        return tmp.name
 
 
# ---------------------------------------------------------------- checks on the existing workbook
def audit_workbook(bk):
    """Return (anomalies, notes, latest_month). Raises Stop on structural problems.
    anomalies and notes are reported only; nothing is changed."""
    anomalies, notes = [], []
    latest = {}
    for name, sid in SHEETS:
        sh = bk.sheets[name]
        rows = sh.data_rows()
        if not rows:
            raise Stop("%s has no data rows" % name)
        months = [sh.month(n) for n in rows]
        for a, b in zip(months, months[1:]):
            if ym_add(b, 1) != a:
                raise Stop("%s: months are not consecutive newest-first between %s and %s (gap or order)."
                           % (name, ym_label(a), ym_label(b)))
        latest[name] = months[0]
        for n in rows:
            if sh.num(n, "B") is None:
                raise Stop("%s row %d: index is empty" % (name, n))
        odd = [ym_label(sh.month(n)) for n in rows if sh.day(n) != 1]
        if odd:
            notes.append("%s: %d dates are not the 1st of the month (%s to %s)" % (name, len(odd), odd[-1], odd[0]))
        for n, r in sh.rows.items():
            extra = [c for c in r["cells"] if col_idx(c) > 4 and r["cells"][c][1]]
            if extra:
                notes.append("%s row %d: unexpected cell(s) %s" % (name, n, ",".join(extra + [""])[:-1]))
        mm = sh.month_map()
        for n in rows:
            em, ey = sh.expected(n, mm)
            for col, e in (("C", em), ("D", ey)):
                s = sh.num(n, col)
                if e is not None and (s is None or abs(s - e) > EPS):
                    anomalies.append("%s %s: %s is %s but the index gives %s" % (
                        name, ym_label(sh.month(n)), "MoM" if col == "C" else "YoY",
                        "empty" if s is None else "%.4f%%" % (s * 100), "%.4f%%" % (e * 100)))
    if len(set(latest.values())) != 1:
        raise Stop("sheets end on different months: %s" % {k: ym_label(v) for k, v in latest.items()})
    return anomalies, notes, next(iter(latest.values()))
 
 
def snapshot(bk):
    out = {}
    for name, sh in bk.sheets.items():
        out[name] = {sh.month(n): {c: sh.num(n, c) for c in "BCD"} for n in sh.data_rows()}
    return out
 
 
# ---------------------------------------------------------------- values.csv
def read_csv(path):
    vals = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for i, row in enumerate(csv.reader(f), 1):
            if not row or not row[0].strip() or row[0].strip().startswith("#"):
                continue
            if row[0].strip().lower() == "series":
                continue
            if len(row) < 3:
                raise Stop("values.csv line %d: need series,month,value" % i)
            sid, ym, v = row[0].strip().upper(), parse_ym(row[1]), row[2].strip()
            try:
                x = float(v.replace(",", ""))
            except ValueError:
                raise Stop("values.csv line %d: %r is not a number" % (i, v))
            if sid not in {s for _, s in SHEETS}:
                raise Stop("values.csv line %d: unknown series %r" % (i, sid))
            if (sid, ym) in vals:
                raise Stop("values.csv line %d: %s %s given twice" % (i, sid, ym_str(ym)))
            vals[(sid, ym)] = x
    return vals
 
 
# ---------------------------------------------------------------- the update
def update(bk, vals):
    anomalies, notes, L = audit_workbook(bk)
    M = ym_add(L, 1)
    before = snapshot(bk)
 
    # --- values.csv completeness and months
    by = {}
    for (sid, ym), x in vals.items():
        by.setdefault(sid, {})[ym] = x
    missing = [s for _, s in SHEETS if s not in by]
    if missing:
        raise Stop("values.csv has no rows for: %s" % ", ".join(missing))
    not_new = [s for s in by if M not in by[s]]
    if not_new:
        mx = {s: ym_label(max(by[s])) for s in not_new}
        raise Stop("%s is not in values.csv for %d series (latest given: %s). The workbook ends %s, so the new month "
                   "is %s. If that month is not published yet for these series, wait; nothing is written."
                   % (ym_label(M), len(not_new), mx, ym_label(L), ym_label(M)))
    for s, d in by.items():
        for ym in d:
            if ym > M:
                raise Stop("%s: %s is after the expected new month %s" % (s, ym_label(ym), ym_label(M)))
            if ym < ym_add(M, -(WINDOW - 1)):
                raise Stop("%s: %s is older than the %d-month revision window" % (s, ym_label(ym), WINDOW))
        if ym_add(M, -1) not in d:
            raise Stop("%s: the previous month %s must be in values.csv (to prove the series lines up)"
                       % (s, ym_label(ym_add(M, -1))))
    # --- compare prior months with the workbook
    plan = {}   # sheet -> {ym: new index} (revisions only)
    for name, sid in SHEETS:
        sh = bk.sheets[name]
        mm = sh.month_map()
        for ym, x in sorted(by[sid].items()):
            if ym == M:
                continue
            if ym not in mm:
                raise Stop("%s: %s is not in the workbook" % (name, ym_label(ym)))
            old = sh.num(mm[ym], "B")
            if abs(old - x) > SAME:
                if abs(x / old - 1) > MAX_REVISION:
                    raise Stop("%s %s: revision %s -> %s is %.1f%%, above the %.0f%% limit (annual revision, "
                               "wrong series or typo?)" % (name, ym_label(ym), fmt(old), fmt(x),
                                                           (x / old - 1) * 100, MAX_REVISION * 100))
                plan.setdefault(name, {})[ym] = x
    # --- new month checks
    for name, sid in SHEETS:
        x = by[sid][M]
        if not x > 0:
            raise Stop("%s: new value %s is not positive" % (name, x))
        prev = plan.get(name, {}).get(ym_add(M, -1))
        if prev is None:
            prev = bk.sheets[name].num(bk.sheets[name].month_map()[ym_add(M, -1)], "B")
        if abs(x / prev - 1) > MAX_MOM:
            raise Stop("%s: %s moves %.1f%% from %s (limit %.0f%%); check the value" % (
                name, ym_label(M), (x / prev - 1) * 100, ym_label(ym_add(M, -1)), MAX_MOM * 100))
 
    # --- apply
    out, revlog, recalc, warns = [], [], [], []
    for name, sid in SHEETS:
        sh = bk.sheets[name]
        mm = sh.month_map()
        for ym, x in sorted(plan.get(name, {}).items()):
            old = sh.num(mm[ym], "B")
            sh.set_num(mm[ym], "B", x)
            revlog.append((name, ym_label(ym), fmt(old), fmt(x), (x / old - 1) * 100))
        sh.insert_top()
        sh.set_num(2, "A", serial(M))
        sh.set_num(2, "B", by[sid][M])
        mm = sh.month_map()
        mom, yoy = sh.expected(2, mm)
        sh.set_num(2, "C", mom)
        sh.set_num(2, "D", yoy)
        if yoy is None:
            warns.append("%s: no %s in the workbook, so YoY for %s is blank" % (
                name, ym_label(ym_add(M, -12)), ym_label(M)))
        out.append((name, by[sid][M], mom, yoy))
        # recompute only the cells that depend on a revised index
        todo = {}
        for ym in plan.get(name, {}):
            todo.setdefault(ym, set()).update("CD")
            todo.setdefault(ym_add(ym, 1), set()).add("C")
            todo.setdefault(ym_add(ym, 12), set()).add("D")
        for ym, cols in sorted(todo.items(), reverse=True):
            if ym not in mm or ym == M:
                continue
            n = mm[ym]
            em, ey = sh.expected(n, mm)
            for col, e in (("C", em), ("D", ey)):
                if col in cols and e is not None:
                    old = sh.num(n, col)
                    if old is None or abs(old - e) > EPS:
                        recalc.append((name, ym_label(ym), "MoM" if col == "C" else "YoY", old, e))
                    sh.set_num(n, col, e)
 
    # --- proof: history unchanged except revised months and the cells that depend on them
    after = snapshot(bk)
    for name, sid in SHEETS:
        if set(after[name]) != set(before[name]) | {M}:
            raise Stop("internal error: %s months changed unexpectedly" % name)
        allowed = {}
        for ym in plan.get(name, {}):
            allowed.setdefault(ym, set()).update("BCD")
            allowed.setdefault(ym_add(ym, 1), set()).add("C")
            allowed.setdefault(ym_add(ym, 12), set()).add("D")
        for ym, cols in before[name].items():
            for c in "BCD":
                if cols[c] != after[name][ym][c] and c not in allowed.get(ym, ()):
                    raise Stop("internal error: %s %s column %s changed" % (name, ym_label(ym), c))
    return M, out, revlog, recalc, warns, anomalies, notes
 
 
# ---------------------------------------------------------------- main
def sha(b):
    return hashlib.sha256(b).hexdigest()
 
 
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("values")
    ap.add_argument("--in-place", action="store_true", help="make a backup in backups/, then replace the workbook")
    ap.add_argument("--out", help="write the result here and leave the workbook alone")
    a = ap.parse_args()
    try:
        if bool(a.in_place) == bool(a.out):
            raise Stop("give exactly one of --in-place or --out FILE")
        orig = open(a.workbook, "rb").read()
        bk = Book(a.workbook)
        if bk.n_charts() != N_CHARTS:
            raise Stop("expected %d charts, found %d" % (N_CHARTS, bk.n_charts()))
        vals = read_csv(a.values)
        M, rows, revlog, recalc, warns, anomalies, notes = update(bk, vals)
        parts = bk.render()
        # everything that is not a sheet must be byte-identical
        for p, b in bk.parts.items():
            if p not in bk.sheet_paths() and parts[p] != b:
                raise Stop("internal error: %s changed" % p)
        tmp = bk.write(a.out if a.out else a.workbook, parts)
        try:
            chk = Book(tmp)
            if chk.n_charts() != N_CHARTS or set(chk.parts) != set(bk.parts):
                raise Stop("internal error: output is missing charts or parts")
        except Exception:
            os.remove(tmp)
            raise
        if a.in_place:
            bdir = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), "backups")
            stem = os.path.splitext(os.path.basename(a.workbook))[0]
            bpath = os.path.join(bdir, "%s_before_%s.xlsx" % (stem, ym_str(M)))
            if os.path.exists(bpath):
                os.remove(tmp)
                raise Stop("backup %s already exists; not overwriting it" % bpath)
            os.makedirs(bdir, exist_ok=True)
            shutil.copyfile(a.workbook, bpath)
            if sha(open(bpath, "rb").read()) != sha(orig):
                os.remove(bpath)
                os.remove(tmp)
                raise Stop("backup could not be verified; workbook not changed")
            os.replace(tmp, a.workbook)
            print("Backup written: %s" % bpath)
            print("Workbook updated: %s" % a.workbook)
        else:
            os.replace(tmp, a.out)
            print("Written: %s (original untouched)" % a.out)
    except Stop as e:
        print("STOP: %s" % e)
        print("Nothing was written.")
        sys.exit(2)
 
    f = lambda x: "-" if x is None else "%.2f%%" % (x * 100)
    print("\nUPDATE OK: %s added to all 21 sheets" % ym_label(M))
    print("%-32s %10s %9s %9s" % ("sheet", "index", "MoM", "YoY"))
    for name, lv, mom, yoy in rows:
        print("%-32s %10s %9s %9s" % (name, fmt(lv), f(mom), f(yoy)))
    print("\nRevisions applied to prior months (old -> new): %d" % len(revlog))
    for r in revlog:
        print("  %-32s %-7s %s -> %s (%+.2f%%)" % r)
    print("\nMoM/YoY cells recomputed because an index was revised: %d" % len(recalc))
    for name, lab, what, old, new in recalc:
        print("  %-32s %-7s %s %s -> %s" % (name, lab, what, f(old), f(new)))
    if warns:
        print("\nWarnings:")
        for w in warns:
            print("  - " + w)
    print("\nExisting data, reported and NOT changed:")
    for x in notes:
        print("  - " + x)
    print("  - %d MoM/YoY cells differ from what the index gives (before this update). First 12:" % len(anomalies))
    for x in anomalies[:12]:
        print("      " + x)
    print("\nCharts: %d kept untouched (their ranges already run to row 1006)." % N_CHARTS)
 
 
if __name__ == "__main__":
    main()