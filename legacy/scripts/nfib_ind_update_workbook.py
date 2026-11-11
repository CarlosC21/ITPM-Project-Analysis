#!/usr/bin/env python3
"""NFIB Small Business Optimism INDUSTRIES workbook updater.
 
Edits the .xlsx at file (XML) level so the 60 charts and all formatting survive.
Standard library only.
 
  python nfib_ind_update_workbook.py "NFIB Small Business Optimism - INDUSTRIES.xlsx" values.csv --out test.xlsx
  python nfib_ind_update_workbook.py "NFIB Small Business Optimism - INDUSTRIES.xlsx" values.csv --in-place
  optional, repeatable:  --correct "Professions,2026-09,97.3"   (fixes one existing typed level, logged)
 
values.csv:  sheet,month,value   (one row per sheet for the NEW month; optional rows for the PRIOR month
are only compared with the workbook and reported, never written).  Lines starting with # are ignored.
 
Any validation failure prints "STOP: ..." and writes nothing.  See nfib_ind_conventions.md.
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
from decimal import Decimal, ROUND_HALF_UP
 
SHEETS = ["Total", "Construction", "Manufacturing", "Transportation", "Wholesale", "Retail",
          "Agriculture", "Finance", "Services", "Professions"]
N_CHARTS = 60
FIRST = 3                # first data row; newest month sits here
NEW_DAY = 26             # new rows are dated the 26th, like the typed rows now in the workbook
MAX_STEP = {"Total": 6.0}   # largest plausible move from the prior month (points); others use DEFAULT_STEP
DEFAULT_STEP = 30.0      # industry series are volatile (Transportation moved 21.8 in Sep-26)
RANGE = {"Total": (80, 120)}
DEFAULT_RANGE = (70, 140)
EPS = 1e-9
CHECK_TOL = 0.051        # prior-month comparison tolerance (1 decimal rounding)
 
 
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
        raise Stop("bad month %r (use YYYY-MM)" % s)
    return (int(m.group(1)), int(m.group(2)))
 
 
def serial(y, m, d):
    return (dt.date(y, m, d) - dt.date(1899, 12, 30)).days
 
 
def from_serial(v):
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(round(v)))
    return (d.year, d.month)
 
 
def fmt(x):
    x = float(x)
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(x)
 
 
def r1(x):
    return float(Decimal(repr(float(x))).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))
 
 
def attrs_of(s):
    return dict(re.findall(r'([\w:]+)=\"([^\"]*)\"', s))
 
 
def col_idx(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n
 
 
# ---------------------------------------------------------------- sheet model
ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
NA, VALUE_ERR, DIV0 = "#N/A", "#VALUE!", "#DIV/0!"
 
 
class Sheet:
    def __init__(self, name, xml):
        self.name = name
        m = re.match(r"(.*<sheetData>)(.*)(</sheetData>.*)", xml, re.S)
        if not m:
            raise Stop("%s: cannot read sheet XML" % name)
        self.head, body, self.tail = m.group(1), m.group(2), m.group(3)
        self.rows = {}
        for rm in ROW_RE.finditer(body):
            cells = {}
            for cm in CELL_RE.finditer(rm.group(3) or ""):
                cells[cm.group(1)] = [cm.group(3), cm.group(4) or ""]
            self.rows[int(rm.group(1))] = {"attrs": rm.group(2), "cells": cells}
        self.charts = []
 
    # -- reading.  kind: blank / num / text / err
    def get(self, n, col):
        c = self.rows.get(n, {}).get("cells", {}).get(col)
        if not c:
            return ("blank", None)
        a, inner = c
        t = attrs_of(a).get("t")
        vm = re.search(r"<v>(.*?)</v>", inner, re.S)
        if t == "e":
            return ("err", vm.group(1) if vm else NA)
        if t in ("s", "str", "inlineStr"):
            return ("text", vm.group(1) if vm else "")
        if vm is None:
            return ("blank", None)
        return ("num", float(vm.group(1)))
 
    def has_formula(self, n, col):
        c = self.rows.get(n, {}).get("cells", {}).get(col)
        return bool(c and "<f" in c[1])
 
    def month_of(self, n):
        k, v = self.get(n, "A")
        if k != "num":
            raise Stop("%s row %d: month cell is not a date" % (self.name, n))
        return from_serial(v)
 
    def data_rows(self):
        out, n = [], FIRST
        while self.get(n, "A")[0] != "blank":
            out.append(n)
            n += 1
        return out
 
    # -- writing
    def set_cell(self, n, col, attrs, inner):
        row = self.rows.setdefault(n, {"attrs": ' spans="1:5"', "cells": {}})
        row["cells"][col] = [attrs, inner]
 
    def serialize(self):
        out = []
        for n in sorted(self.rows):
            r = self.rows[n]
            cs = []
            for col in sorted(r["cells"], key=col_idx):
                a, inner = r["cells"][col]
                cs.append('<c r="%s%d"%s>%s</c>' % (col, n, a, inner) if inner else '<c r="%s%d"%s/>' % (col, n, a))
            out.append('<row r="%d"%s>%s</row>' % (n, r["attrs"], "".join(cs)))
        return self.head + "".join(out) + self.tail
 
    # -- calculation (what Excel would give for the three formula columns)
    def calc(self, last):
        """Return {(row, col): ('num', x) | ('err', code)} for C, D, E in rows FIRST..last."""
        def B(n):
            return self.get(n, "B")
 
        def arith(x):  # value used in a subtraction
            if x[0] == "blank":
                return ("num", 0.0)
            if x[0] == "text":
                return ("err", VALUE_ERR)
            return x
 
        res = {}
        for n in range(FIRST, last + 1):
            b2 = B(n + 2)
            if b2[0] == "blank":
                res[(n, "C")] = ("err", NA)
            else:
                vals = [B(k) for k in (n, n + 1, n + 2)]
                errs = [v for v in vals if v[0] == "err"]
                nums = [v[1] for v in vals if v[0] == "num"]
                if errs:
                    res[(n, "C")] = errs[0]
                elif not nums:
                    res[(n, "C")] = ("err", DIV0)
                else:
                    res[(n, "C")] = ("num", sum(nums) / len(nums))
            x, y = arith(B(n)), arith(B(n + 12))
            if x[0] == "err":
                res[(n, "D")] = x
            elif y[0] == "err":
                res[(n, "D")] = y
            else:
                res[(n, "D")] = ("num", x[1] - y[1])
        for n in range(FIRST, last + 1):
            x = res[(n, "C")]
            y = res.get((n + 12, "C"), ("num", 0.0))
            if x[0] == "err":
                res[(n, "E")] = x
            elif y[0] == "err":
                res[(n, "E")] = y
            else:
                res[(n, "E")] = ("num", x[1] - y[1])
        return res
 
 
def formula_xml(col, n):
    return {"C": 'IF(B%d="",#N/A,AVERAGE(B%d:B%d))' % (n + 2, n, n + 2),
            "D": "B%d-B%d" % (n, n + 12),
            "E": "C%d-C%d" % (n, n + 12)}[col]
 
 
def put_formula(sh, n, col, style, res):
    k, v = res
    if k == "err":
        sh.set_cell(n, col, '%s t="e"' % style, "<f>%s</f><v>%s</v>" % (formula_xml(col, n), v))
    else:
        sh.set_cell(n, col, style, "<f>%s</f><v>%s</v>" % (formula_xml(col, n), fmt(v)))
 
 
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
        self.sheets, self.order = {}, []
        for m in re.finditer(r"<sheet ([^>]*?)/>", wb):
            a = attrs_of(m.group(1))
            path = rid[a["r:id"]]
            self.order.append((a["name"], path))
            self.sheets[a["name"]] = Sheet(a["name"], self.parts[path].decode("utf-8"))
        if [n for n, _ in self.order] != SHEETS:
            raise Stop("sheet names/order differ from the expected %d: %s" % (len(SHEETS), [n for n, _ in self.order]))
        for name, path in self.order:
            relp = posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path) + ".rels")
            if relp not in self.parts:
                continue
            for m in re.finditer(r"<Relationship ([^>]*?)/>", self.parts[relp].decode()):
                a = attrs_of(m.group(1))
                if a["Type"].endswith("/drawing"):
                    dpath = posixpath.normpath(posixpath.join(posixpath.dirname(path), a["Target"]))
                    drel = posixpath.join(posixpath.dirname(dpath), "_rels", posixpath.basename(dpath) + ".rels")
                    for dm in re.finditer(r"<Relationship ([^>]*?)/>", self.parts[drel].decode()):
                        da = attrs_of(dm.group(1))
                        if da["Type"].endswith("/chart"):
                            self.sheets[name].charts.append(
                                posixpath.normpath(posixpath.join(posixpath.dirname(dpath), da["Target"])))
 
    def n_charts(self):
        return sum(len(s.charts) for s in self.sheets.values())
 
    def refresh_charts(self):
        """Rebuild every chart's cached points from the sheet cells (ranges are unchanged)."""
        for name, sh in self.sheets.items():
            for cp in sh.charts:
                xml = self.parts[cp].decode("utf-8")
 
                def rebuild(m):
                    f, fc, cnt = m.group(2), m.group(3), m.group(4)
                    mm = re.fullmatch(r"(?:'[^']+'|[^!]+)!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)", f)
                    if not mm:
                        raise Stop("chart %s: cannot read range %s" % (cp, f))
                    col, r1_, r2_ = mm.group(1), int(mm.group(2)), int(mm.group(3))
                    pts = []
                    for n in range(r1_, r2_ + 1):
                        k, v = sh.get(n, col)
                        if k == "num":
                            pts.append('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (n - r1_, fmt(v)))
                        elif n > max(sh.rows):
                            break
                    return "<c:%s><c:numRef><c:f>%s</c:f><c:numCache>%s%s%s</c:numCache>" % (
                        m.group(1), f, fc, cnt, "".join(pts))
 
                xml, k = re.subn(r"<c:(cat|val)><c:numRef><c:f>(.*?)</c:f><c:numCache>"
                                 r"(<c:formatCode>.*?</c:formatCode>)(<c:ptCount val=\"\d+\"/>).*?</c:numCache>",
                                 rebuild, xml, flags=re.S)
                if k == 0:
                    raise Stop("%s: chart %s has no series I can rebuild" % (name, cp))
                self.parts[cp] = xml.encode("utf-8")
 
    def render(self):
        parts = dict(self.parts)
        for name, path in self.order:
            parts[path] = self.sheets[name].serialize().encode("utf-8")
        # Excel rebuilds the calculation chain; a stale one can make the file fail to open
        parts.pop("xl/calcChain.xml", None)
        ct = parts["[Content_Types].xml"].decode("utf-8")
        parts["[Content_Types].xml"] = re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct).encode("utf-8")
        rl = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rl).encode("utf-8")
        wb = parts["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = wb.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
        parts["xl/workbook.xml"] = wb.encode("utf-8")
        for p, b in parts.items():
            if p.endswith(".xml") or p.endswith(".rels"):
                try:
                    xml.dom.minidom.parseString(b)
                except Exception as e:
                    raise Stop("internal error: %s is not well-formed XML (%s)" % (p, e))
        return parts
 
    def write(self, path, parts):
        tmp = tempfile.NamedTemporaryFile(delete=False, dir=os.path.dirname(os.path.abspath(path)), suffix=".tmp")
        tmp.close()
        with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as z:
            for i in self.infos:
                if i.filename not in parts:
                    continue
                zi = zipfile.ZipInfo(i.filename, i.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = i.external_attr
                z.writestr(zi, parts[i.filename])
        return tmp.name
 
 
# ---------------------------------------------------------------- checks on the existing workbook
def audit(bk):
    """Return (anomalies, latest_month, last_row). Raises Stop on structural problems."""
    anomalies, latest, last = [], {}, {}
    for name in SHEETS:
        sh = bk.sheets[name]
        rows = sh.data_rows()
        if not rows:
            raise Stop("%s has no data rows from row %d" % (name, FIRST))
        if rows != list(range(FIRST, FIRST + len(rows))):
            raise Stop("%s: data rows are not contiguous" % name)
        months = [sh.month_of(n) for n in rows]
        for a, b in zip(months, months[1:]):
            if ym_add(b, 1) != a:
                raise Stop("%s: months are not consecutive newest-first between %s and %s"
                           % (name, ym_label(a), ym_label(b)))
        latest[name], last[name] = months[0], rows[-1]
        if len(rows) < 14:
            raise Stop("%s has only %d months; need at least 14 (year-ago comparison)" % (name, len(rows)))
        nxt = rows[-1] + 1
        for col in "ABCDE":
            if sh.get(nxt, col)[0] != "blank":
                raise Stop("%s: row %d (below the data) is not empty in column %s" % (name, nxt, col))
        for n in rows:
            k, v = sh.get(n, "B")
            if k == "text":
                anomalies.append("%s %s: level is text, not a number (3-month average ignores it, YoY gives an error)"
                                 % (name, ym_label(sh.month_of(n))))
            elif k != "num":
                anomalies.append("%s %s: level is %s" % (name, ym_label(sh.month_of(n)), k))
        calc = sh.calc(rows[-1])
        for (n, col), (k, v) in sorted(calc.items()):
            sk, sv = sh.get(n, col)
            same = (sk == k) and (k == "err" and sv == v or k == "num" and abs(sv - v) <= EPS)
            if not same:
                anomalies.append("%s %s col %s: stored %s but the levels give %s" % (
                    name, ym_label(sh.month_of(n)), col,
                    sv if sk != "blank" else "empty", v if k == "err" else "%.6g" % v))
    if len(set(latest.values())) != 1:
        raise Stop("sheets end on different months: %s" % {k: ym_label(v) for k, v in latest.items()})
    if len(set(last.values())) != 1:
        raise Stop("sheets have different numbers of rows: %s" % last)
    return anomalies, next(iter(latest.values())), next(iter(last.values()))
 
 
def snapshot(bk):
    out = {}
    for name in SHEETS:
        sh = bk.sheets[name]
        out[name] = {sh.month_of(n): {c: sh.get(n, c) for c in "ABCDE"} for n in sh.data_rows()}
    return out
 
 
# ---------------------------------------------------------------- monthly update
def read_csv(path):
    vals = {}
    low = {s.lower(): s for s in SHEETS}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for i, row in enumerate(csv.reader(f), 1):
            if not row or not row[0].strip() or row[0].strip().startswith("#"):
                continue
            if row[0].strip().lower() == "sheet":
                continue
            if len(row) < 3:
                raise Stop("values.csv line %d: need sheet,month,value" % i)
            s = low.get(row[0].strip().lower())
            if s is None:
                raise Stop("values.csv line %d: unknown sheet %r" % (i, row[0]))
            ym = parse_ym(row[1])
            try:
                x = float(row[2].strip().replace(",", ""))
            except ValueError:
                raise Stop("values.csv line %d: %r is not a number" % (i, row[2]))
            if (s, ym) in vals:
                raise Stop("values.csv line %d: %s %s given twice" % (i, s, ym_str(ym)))
            vals[(s, ym)] = x
    return vals
 
 
def parse_corrections(items):
    out = []
    for it in items or []:
        p = [x.strip() for x in it.split(",")]
        if len(p) != 3 or p[0] not in SHEETS:
            raise Stop("--correct needs 'Sheet name,YYYY-MM,value' with an exact sheet name, got %r" % it)
        try:
            out.append((p[0], parse_ym(p[1]), float(p[2])))
        except ValueError:
            raise Stop("--correct: %r is not a number" % p[2])
    return out
 
 
def update(bk, vals, corrections):
    anomalies, latest, last = audit(bk)
    M, M1 = ym_add(latest, 1), latest
    warns = []
 
    need = {(s, M) for s in SHEETS}
    missing = sorted(need - set(vals))
    if missing:
        raise Stop("values.csv is missing: %s. Expected month %s for all %d sheets."
                   % (", ".join("%s %s" % (s, ym_str(y)) for s, y in missing), ym_str(M), len(SHEETS)))
    extra = sorted(k for k in set(vals) if k[1] not in (M, M1))
    if extra:
        raise Stop("values.csv has rows for the wrong month (workbook ends %s, so the new month is %s): %s"
                   % (ym_label(latest), ym_str(M), ", ".join("%s %s" % (s, ym_str(y)) for s, y in extra[:6])))
 
    # plausibility of the new values
    new = {}
    for s in SHEETS:
        v = r1(vals[(s, M)])
        lo, hi = RANGE.get(s, DEFAULT_RANGE)
        if not lo <= v <= hi:
            raise Stop("%s %s: %s is outside %s..%s; check the figure" % (s, ym_str(M), v, lo, hi))
        prev = bk.sheets[s].get(FIRST, "B")
        if prev[0] == "num" and abs(v - prev[1]) > MAX_STEP.get(s, DEFAULT_STEP):
            raise Stop("%s %s: moved %.1f from %s (limit %.0f); check the figure"
                       % (s, ym_str(M), v - prev[1], ym_label(M1), MAX_STEP.get(s, DEFAULT_STEP)))
        new[s] = v
 
    before = snapshot(bk)
    sheetmap = {s: bk.sheets[s] for s in SHEETS}
 
    # corrections to existing levels (explicit, logged)
    corlog = []
    for s, ym, v in corrections:
        sh = sheetmap[s]
        rows = {sh.month_of(n): n for n in sh.data_rows()}
        if ym not in rows:
            raise Stop("--correct: %s has no %s row" % (s, ym_str(ym)))
        n = rows[ym]
        if sh.has_formula(n, "B"):
            raise Stop("--correct: %s %s level is a formula (link to the other workbook); not changing it" % (s, ym_str(ym)))
        old = sh.get(n, "B")
        sty = re.sub(r'\s+t="[^"]*"', "", sh.rows[n]["cells"]["B"][0])
        sh.set_cell(n, "B", sty, "<v>%s</v>" % fmt(v))
        corlog.append((s, ym_label(ym), old[1] if old[0] != "blank" else "(empty)", fmt(v)))
 
    # prior-month comparison (report only)
    drift = []
    for s in SHEETS:
        if (s, M1) in vals:
            w = bk.sheets[s].get(FIRST, "B")
            if w[0] != "num":
                drift.append("%s %s: workbook has no number to compare with the source's %s" % (s, ym_label(M1), r1(vals[(s, M1)])))
            elif abs(w[1] - vals[(s, M1)]) > CHECK_TOL:
                drift.append("%s %s: workbook %s, source now shows %s" % (s, ym_label(M1), fmt(w[1]), fmt(r1(vals[(s, M1)]))))
    if not any(k[1] == M1 for k in vals):
        warns.append("no prior-month (%s) rows in values.csv, so the source's revisions were not checked" % ym_str(M1))
    warns.extend("prior month differs: " + d for d in drift)
 
    # shift the A:E block down one row, add the new row at FIRST
    for s in SHEETS:
        sh = sheetmap[s]
        L = last
        tpl = {c: sh.rows[FIRST]["cells"][c][0] for c in "ABCDE"}
        tpl = {c: re.sub(r'\s+t="[^"]*"', "", a) for c, a in tpl.items()}
        moved = {}
        for n in range(FIRST, L + 1):
            moved[n + 1] = {c: list(sh.rows[n]["cells"][c]) for c in "AB" if c in sh.rows[n]["cells"]}
        for n, cells in moved.items():
            for c, (a, inner) in cells.items():
                if c == "B" and "<f>" in inner:  # same-sheet reference to column A moves with the row
                    inner = inner.replace("$A%d" % (n - 1), "$A%d" % n)
                sh.set_cell(n, c, a, inner)
        sh.set_cell(FIRST, "A", tpl["A"], "<v>%d</v>" % serial(M[0], M[1], NEW_DAY))
        sh.set_cell(FIRST, "B", tpl["B"], "<v>%s</v>" % fmt(new[s]))
        res = sh.calc(L + 1)
        for n in range(FIRST, L + 2):
            for c in "CDE":
                put_formula(sh, n, c, tpl[c], res[(n, c)])
    bk.refresh_charts()
 
    # proof: history unchanged (A and B) except what --correct touched; computed columns may change only where
    # a corrected level feeds them or the old cache was already inconsistent
    after = snapshot(bk)
    touched = {(s, ym) for s, ym, _ in corrections}
    changed = []
    for s in SHEETS:
        if set(after[s]) != set(before[s]) | {M}:
            raise Stop("internal error: %s months changed unexpectedly" % s)
        for ym, cols in before[s].items():
            for c in "ABCDE":
                x, y = cols[c], after[s][ym][c]
                same = x[0] == y[0] and (x[0] == "blank" or (x[0] == "num" and abs(x[1] - y[1]) <= EPS)
                                         or (x[0] in ("err", "text") and x[1] == y[1]))
                if c in "AB" and not same and (s, ym) not in touched:
                    raise Stop("internal error: %s %s column %s changed" % (s, ym_label(ym), c))
                if c in "CDE" and not same:
                    changed.append((s, ym_label(ym), c, x, y))
    # no shared-formula stub may survive anywhere (they were all in C:E and have been rewritten)
    for s in SHEETS:
        for n, r in sheetmap[s].rows.items():
            for c, (a, inner) in r["cells"].items():
                if 't="shared"' in inner:
                    raise Stop("internal error: %s %s%d still has a shared formula" % (s, c, n))
    return M, new, after, corlog, warns, anomalies, changed
 
 
# ---------------------------------------------------------------- main
def sha(b):
    return hashlib.sha256(b).hexdigest()
 
 
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("values")
    ap.add_argument("--in-place", action="store_true", help="make a backup in backups/, then replace the workbook")
    ap.add_argument("--out", help="write the result here and leave the workbook alone")
    ap.add_argument("--correct", action="append", help='"Sheet,YYYY-MM,value" fix one existing typed level (repeatable)')
    a = ap.parse_args()
    try:
        if bool(a.in_place) == bool(a.out):
            raise Stop("give exactly one of --in-place or --out FILE")
        orig = open(a.workbook, "rb").read()
        bk = Book(a.workbook)
        if bk.n_charts() != N_CHARTS:
            raise Stop("expected %d charts, found %d" % (N_CHARTS, bk.n_charts()))
        vals = read_csv(a.values)
        M, new, after, corlog, warns, anomalies, changed = update(bk, vals, parse_corrections(a.correct))
        parts = bk.render()
        tmp = bk.write(a.out if a.out else a.workbook, parts)
        try:
            chk = Book(tmp)
            if chk.n_charts() != N_CHARTS or set(chk.parts) != set(parts):
                raise Stop("internal error: output is missing charts or parts")
            a2, latest2, _ = audit(chk)
            if latest2 != M:
                raise Stop("internal error: output ends on %s, not %s" % (ym_label(latest2), ym_label(M)))
            leftovers = [x for x in a2 if "level is" not in x]
            if leftovers:
                raise Stop("internal error: output is not self-consistent: %s" % leftovers[0])
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
 
    print("\nUPDATE OK: %s added to all %d sheets" % (ym_label(M), len(SHEETS)))
    print("%-16s %8s %8s %8s %8s" % ("sheet", "value", "3mma", "YoY", "3mma YoY"))
    for s in SHEETS:
        row = after[s][M]
        g = lambda c: ("%.2f" % row[c][1]) if row[c][0] == "num" else str(row[c][1])
        print("%-16s %8s %8s %8s %8s" % (s, g("B"), g("C"), g("D"), g("E")))
    if corlog:
        print("\nCorrections applied (old -> new):")
        for c in corlog:
            print("  %-16s %-7s %s -> %s" % c)
    if changed:
        print("\nComputed cells (3mma / YoY) whose stored value changed: %d" % len(changed))
        for s, lab, c, x, y in changed[:20]:
            print("  %-16s %-7s col %s: %s -> %s" % (s, lab, c, x[1], y[1] if y[0] != "blank" else "empty"))
    if warns:
        print("\nWarnings:")
        for w in warns:
            print("  - " + w)
    print("\nAnomaly audit of the workbook before the update (not changed unless listed above): %d" % len(anomalies))
    for x in anomalies[:25]:
        print("  - " + x)
    print("\nCharts: %d kept, cached points rebuilt. Excel will recalculate when the file is opened." % N_CHARTS)
 
 
if __name__ == "__main__":
    main()