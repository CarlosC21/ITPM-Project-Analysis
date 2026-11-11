#!/usr/bin/env python3
"""US Employment Situation workbook updater.
 
Edits the .xlsx at file (XML) level so the 22 charts and all formatting survive.
Standard library only.
 
  python employment_update_workbook.py US_Employment_Report.xlsx values.csv --in-place
  python employment_update_workbook.py US_Employment_Report.xlsx --restate --in-place   (one time)
 
Any validation failure prints "STOP: ..." and writes nothing.
See employment_conventions.md for layout, rules and the values.csv format.
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
 
# sheet name, csv key, kind  (kind: rate = percent, pp-difference rule; level = thousands; derived = TOTAL - GOODS)
SHEETS = [
    ("Unemployment_Rate", "UR", "rate"),
    ("Total_Nonfarm_Payrolls", "TOTAL", "level"),
    ("Total_Government", "GOVERNMENT", "level"),
    ("Total_Private", "PRIVATE", "level"),
    ("Goods_Producing", "GOODS", "level"),
    ("Service_Producing", "SERVICE", "derived"),
    ("Construction", "CONSTRUCTION", "level"),
    ("Manufacturing", "MANUFACTURING", "level"),
    ("Trade_Transport_Utilities", "TTU", "level"),
    ("Financials", "FINANCIAL", "level"),
    ("Professional_Business_Services", "PBS", "level"),
]
CSV_KEYS = ["UR", "TOTAL", "PRIVATE", "GOODS", "CONSTRUCTION", "MANUFACTURING",
            "PSP", "TTU", "FINANCIAL", "PBS", "GOVERNMENT"]
YEAR_AGO_KEYS = ["TOTAL", "PRIVATE", "GOODS", "CONSTRUCTION", "MANUFACTURING",
                 "PSP", "TTU", "FINANCIAL", "PBS", "GOVERNMENT"]
META_KEYS = ["HEADLINE_CHANGE", "REVISED_CHANGE", "ORIGINAL_CHANGE"]
UR_NOTE_OCT25 = "not published"
UR_NOTE_NOV25 = "not entered"
TOL_ID = 1          # identity tolerance, thousands (rounding)
MAX_REVISION = 300  # largest plausible level revision to a prior month, thousands
MAX_MOM = 0.01      # largest plausible monthly % move of a level series
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
 
 
def xml_escape(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
 
 
# ---------------------------------------------------------------- sheet model
ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
 
 
class Sheet:
    def __init__(self, name, xml):
        self.name = name
        m = re.match(r"(.*<sheetData>)(.*)(</sheetData>.*)", xml, re.S)
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
        self.charts = []  # chart part names
 
    # -- reading
    def num(self, n, col):
        c = self.rows.get(n, {}).get("cells", {}).get(col)
        if not c or "t=" in c[0]:
            return None
        m = re.search(r"<v>(.*?)</v>", c[1])
        return float(m.group(1)) if m else None
 
    def text(self, n, col):
        c = self.rows.get(n, {}).get("cells", {}).get(col)
        if not c:
            return None
        m = re.search(r"<t[^>]*>(.*?)</t>", c[1])
        return m.group(1) if m else None
 
    def last_row(self):
        return max(self.rows)
 
    def data_rows(self):
        return sorted(n for n in self.rows if n >= 2)
 
    def month(self, n):
        v = self.num(n, "A")
        if v is None:
            raise Stop("%s row %d: month cell is not a date" % (self.name, n))
        return from_serial(v)[0]
 
    def month_map(self):
        out = {}
        for n in self.data_rows():
            ym = self.month(n)
            if ym in out:
                raise Stop("%s: month %s appears twice" % (self.name, ym_label(ym)))
            out[ym] = n
        return out
 
    # -- writing
    def set_num(self, n, col, val):
        c = self.rows[n]["cells"].setdefault(col, ["", ""])
        c[0] = re.sub(r'\s+t="[^"]*"', "", c[0])
        c[1] = "" if val is None else "<v>%s</v>" % fmt(val)
 
    def set_note(self, n, text):
        self.rows[n]["cells"]["E"] = [' t="inlineStr"', "<is><t>%s</t></is>" % xml_escape(text)]
 
    def insert_rows(self, at, count, template_row):
        shifted = {}
        for n, r in self.rows.items():
            shifted[n + count if n >= at else n] = r
        tpl = shifted[template_row if template_row < at else template_row + count]
        for k in range(count):
            cells = {}
            for col, (a, _) in tpl["cells"].items():
                if col in "ABCD":
                    cells[col] = [re.sub(r'\s+t="[^"]*"', "", a), ""]
            shifted[at + k] = {"attrs": tpl["attrs"], "cells": cells}
        self.rows = shifted
 
    def serialize(self):
        out = []
        maxcol = 4
        for n in sorted(self.rows):
            r = self.rows[n]
            attrs = r["attrs"]
            if "E" in r["cells"]:
                maxcol = max(maxcol, 5)
                attrs = re.sub(r'spans="[^"]*"', 'spans="1:5"', attrs)
            cs = []
            for col in sorted(r["cells"], key=col_idx):
                a, inner = r["cells"][col]
                cs.append('<c r="%s%d"%s>%s</c>' % (col, n, a, inner) if inner
                          else '<c r="%s%d"%s/>' % (col, n, a))
            out.append('<row r="%d"%s>%s</row>' % (n, attrs, "".join(cs)))
        head = re.sub(r'<dimension ref="[^"]*"/>',
                      '<dimension ref="A1:%s%d"/>' % (col_name(maxcol), self.last_row()), self.head)
        return head + "".join(out) + self.tail
 
    # -- calculation
    def expected(self, kind, n, mm=None):
        mm = mm or self.month_map()
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
            val = (lv - p) / 100 if kind == "rate" else lv / p - 1
            if which == "m":
                mom = val
            else:
                yoy = val
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
            path = rid[a["r:id"]]
            self.order.append((a["name"], path))
            self.sheets[a["name"]] = Sheet(a["name"], self.parts[path].decode("utf-8"))
        names = [n for n, _, _ in SHEETS]
        if [n for n, _ in self.order] != names:
            raise Stop("sheet names/order differ from the expected 11: %s" % [n for n, _ in self.order])
        # sheet -> charts
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
        for s in self.sheets.values():
            s.charts.sort(key=lambda p: int(re.search(r"(\d+)", posixpath.basename(p)).group(1)))
 
    def n_charts(self):
        return sum(len(s.charts) for s in self.sheets.values())
 
    def refresh_charts(self):
        """Point every chart at its own sheet, rows 2..last, and rebuild the caches."""
        for name, sh in self.sheets.items():
            last = sh.last_row()
            for cp in sh.charts:
                xml = self.parts[cp].decode("utf-8")
                xml = re.sub(r"(<c:tx><c:strRef><c:f>)[^!<]+(!\$[A-Z]+\$1</c:f>)",
                             lambda m: m.group(1) + name + m.group(2), xml)
 
                def rebuild(m):
                    kind, f, cache = m.group(1), m.group(2), m.group(3)
                    col = re.search(r"\$([A-Z]+)\$\d+", f).group(1)
                    fc = re.search(r"<c:formatCode>.*?</c:formatCode>", cache, re.S).group(0)
                    pts = []
                    for n in range(2, last + 1):
                        v = sh.num(n, col)
                        if v is not None:
                            pts.append('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (n - 2, fmt(v)))
                    newf = "%s!$%s$2:$%s$%d" % (name, col, col, last)
                    return ("<c:%s><c:numRef><c:f>%s</c:f><c:numCache>%s<c:ptCount val=\"%d\"/>%s</c:numCache>"
                            % (kind, newf, fc, last - 1, "".join(pts)))
 
                xml, k = re.subn(r"<c:(cat|val)><c:numRef><c:f>(.*?)</c:f><c:numCache>(.*?)</c:numCache>",
                                 rebuild, xml, flags=re.S)
                if k == 0:
                    raise Stop("%s: chart %s has no series I can rebuild" % (name, cp))
                self.parts[cp] = xml.encode("utf-8")
 
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
    """Return (anomalies, latest_month). Raises Stop on structural problems."""
    anomalies = []
    latest = {}
    for name, key, kind in SHEETS:
        sh = bk.sheets[name]
        rows = sh.data_rows()
        if not rows:
            raise Stop("%s has no data rows" % name)
        months = [sh.month(n) for n in rows]
        for n, ym in zip(rows, months):
            d = from_serial(sh.num(n, "A"))[1]
            if d != 1:
                raise Stop("%s row %d: date is day %d, not the 1st. Run --restate once first." % (name, n, d))
        for a, b in zip(months, months[1:]):
            if ym_add(b, 1) != a:
                raise Stop("%s: months are not consecutive newest-first between %s and %s (gap or order). "
                           "Run --restate once first, or fix the sheet." % (name, ym_label(a), ym_label(b)))
        latest[name] = months[0]
        if kind != "rate":
            for n in rows:
                if sh.num(n, "B") is None:
                    raise Stop("%s row %d: level is empty" % (name, n))
        mm = sh.month_map()
        for n in rows:
            em, ey = sh.expected(kind, n, mm)
            for col, e in (("C", em), ("D", ey)):
                s = sh.num(n, col)
                if e is not None and (s is None or abs(s - e) > EPS):
                    anomalies.append("%s %s: %s is %s but the levels give %s" % (
                        name, ym_label(sh.month(n)), "MoM" if col == "C" else "YoY",
                        "empty" if s is None else "%.4f%%" % (s * 100), "%.4f%%" % (e * 100)))
    if len(set(latest.values())) != 1:
        raise Stop("sheets end on different months: %s" % {k: ym_label(v) for k, v in latest.items()})
    # level identities over all history
    T, G, P, V, GO = (bk.sheets[x] for x in ("Total_Nonfarm_Payrolls", "Total_Government", "Total_Private",
                                              "Service_Producing", "Goods_Producing"))
    mt = T.month_map()
    for ym, n in mt.items():
        vals = {}
        for nm, sh in (("T", T), ("G", G), ("P", P), ("V", V), ("GO", GO)):
            r = sh.month_map().get(ym)
            vals[nm] = sh.num(r, "B") if r else None
        if None in vals.values():
            anomalies.append("identity check skipped for %s: a sheet has no value" % ym_label(ym))
            continue
        if abs(vals["T"] - vals["G"] - vals["P"]) > TOL_ID:
            anomalies.append("%s: Total != Government + Private (diff %d)" % (
                ym_label(ym), vals["T"] - vals["G"] - vals["P"]))
        if abs(vals["T"] - vals["GO"] - vals["V"]) > TOL_ID:
            anomalies.append("%s: Total != Goods + Service-providing (diff %d)" % (
                ym_label(ym), vals["T"] - vals["GO"] - vals["V"]))
    return anomalies, next(iter(latest.values()))
 
 
def snapshot(bk):
    """{sheet: {ym: {col: value}}} for A-D, used to prove history is unchanged."""
    out = {}
    for name, sh in bk.sheets.items():
        d = {}
        for n in sh.data_rows():
            ym = sh.month(n)
            d[ym] = {c: sh.num(n, c) for c in "BCD"}
        out[name] = d
    return out
 
 
# ---------------------------------------------------------------- restatement (one time)
def restate(bk, nov25):
    log = []
    before = snapshot(bk)
    # 1. dates to the 1st
    for name, key, kind in SHEETS:
        sh = bk.sheets[name]
        for n in sh.data_rows():
            ym = sh.month(n)
            old = sh.num(n, "A")
            if from_serial(old)[1] != 1:
                sh.set_num(n, "A", serial(ym))
                log.append((name, ym_label(ym), "date", "day %d" % from_serial(old)[1], "day 1"))
    # 2. unemployment: add Oct-25 (not published) and Nov-25 rows
    ur = bk.sheets["Unemployment_Rate"]
    mm = ur.month_map()
    want = [(2025, 11), (2025, 10)]  # newest first
    if not all(w in mm for w in want):
        if any(w in mm for w in want):
            raise Stop("Unemployment_Rate has only one of Oct-25 / Nov-25; fix by hand")
        at = mm[(2025, 12)] + 1
        if mm.get((2025, 9)) != at:
            raise Stop("Unemployment_Rate: Sep-25 does not follow Dec-25 as expected; cannot insert rows")
        ur.insert_rows(at, 2, mm[(2025, 12)])
        for k, ym in enumerate(want):
            n = at + k
            ur.set_num(n, "A", serial(ym))
        ur.set_note(at, UR_NOTE_NOV25 if nov25 is None else "Nov-25: rate entered by owner")
        ur.set_note(at + 1, UR_NOTE_OCT25)
        if nov25 is not None:
            ur.set_num(at, "B", nov25)
        log.append(("Unemployment_Rate", "Oct-25", "row added", "(missing)", "blank, note in column E"))
        log.append(("Unemployment_Rate", "Nov-25", "row added", "(missing)",
                    "blank, note in column E" if nov25 is None else "%s" % nov25))
    # 3. recompute MoM/YoY where the sheet holds the inputs and the stored value is wrong
    for name, key, kind in SHEETS:
        sh = bk.sheets[name]
        mm = sh.month_map()
        for n in sh.data_rows():
            em, ey = sh.expected(kind, n, mm)
            for col, e, lab in (("C", em, "MoM"), ("D", ey, "YoY")):
                s = sh.num(n, col)
                if e is not None and (s is None or abs(s - e) > EPS):
                    sh.set_num(n, col, e)
                    log.append((name, ym_label(sh.month(n)), lab,
                                "empty" if s is None else "%.4f%%" % (s * 100), "%.4f%%" % (e * 100)))
    # 4. charts
    bk.refresh_charts()
    # proof: levels untouched
    after = snapshot(bk)
    for name in before:
        for ym, cols in before[name].items():
            if after[name][ym]["B"] != cols["B"]:
                raise Stop("internal error: restatement changed a level (%s %s)" % (name, ym_label(ym)))
    return log
 
 
# ---------------------------------------------------------------- monthly update
def read_csv(path):
    vals, meta = {}, {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for i, row in enumerate(csv.reader(f), 1):
            if not row or not row[0].strip() or row[0].strip().startswith("#"):
                continue
            if row[0].strip().lower() == "key":
                continue
            if len(row) < 3:
                raise Stop("values.csv line %d: need key,month,value" % i)
            key, ym, v = row[0].strip().upper(), parse_ym(row[1]), row[2].strip()
            try:
                x = float(v.replace(",", ""))
            except ValueError:
                raise Stop("values.csv line %d: %r is not a number" % (i, v))
            if key in CSV_KEYS:
                if (key, ym) in vals:
                    raise Stop("values.csv line %d: %s %s given twice" % (i, key, ym_str(ym)))
                vals[(key, ym)] = x
            elif key in META_KEYS:
                if (key, ym) in meta:
                    raise Stop("values.csv line %d: %s %s given twice" % (i, key, ym_str(ym)))
                meta[(key, ym)] = x
            else:
                raise Stop("values.csv line %d: unknown key %r" % (i, key))
    return vals, meta
 
 
def update(bk, vals, meta, accept_prev=False):
    anomalies, latest = audit_workbook(bk)
    M = ym_add(latest, 1)
    M1, M2, M12 = ym_add(M, -1), ym_add(M, -2), ym_add(M, -12)
    notes, warns = [], []
 
    # --- completeness
    need = set()
    for k in CSV_KEYS:
        for ym in (M, M1, M2):
            need.add((k, ym))
    for k in YEAR_AGO_KEYS:
        need.add((k, M12))
    missing = sorted(need - set(vals), key=lambda t: (t[0], t[1]))
    if missing:
        raise Stop("values.csv is missing: " + ", ".join("%s %s" % (k, ym_str(y)) for k, y in missing[:12]) +
                   (" ..." if len(missing) > 12 else "") +
                   ". Expected month %s (+ the two prior months and the year-ago level)." % ym_str(M))
    extra = sorted(set(vals) - need)
    if extra:
        raise Stop("values.csv has unexpected rows (wrong month? workbook ends %s, so the new month is %s): %s"
                   % (ym_label(latest), ym_str(M), ", ".join("%s %s" % (k, ym_str(y)) for k, y in extra[:8])))
    for key in ("HEADLINE_CHANGE",):
        if (key, M) not in meta:
            raise Stop("values.csv needs HEADLINE_CHANGE,%s,<published change in thousands>" % ym_str(M))
    for ym in (M1, M2):
        for key in ("REVISED_CHANGE", "ORIGINAL_CHANGE"):
            if (key, ym) not in meta:
                raise Stop("values.csv needs %s,%s,<value>" % (key, ym_str(ym)))
    if set(meta) - {("HEADLINE_CHANGE", M)} - {(k, y) for k in ("REVISED_CHANGE", "ORIGINAL_CHANGE") for y in (M1, M2)}:
        raise Stop("values.csv has unexpected HEADLINE/REVISED/ORIGINAL rows")
 
    def v(k, ym):
        return vals[(k, ym)]
 
    # --- derived series and identities (for every month supplied)
    months_all = [M, M1, M2]
    for ym in months_all:
        vals[("SERVICE", ym)] = v("TOTAL", ym) - v("GOODS", ym)
    vals[("SERVICE", M12)] = v("TOTAL", M12) - v("GOODS", M12)
    for ym in months_all + [M12]:
        t, p, g, psp, gov = (v(k, ym) for k in ("TOTAL", "PRIVATE", "GOODS", "PSP", "GOVERNMENT"))
        if abs(t - p - gov) > TOL_ID:
            raise Stop("%s: Total %d != Private %d + Government %d" % (ym_label(ym), t, p, gov))
        if abs(p - g - psp) > TOL_ID:
            raise Stop("%s: Private %d != Goods %d + Private service-providing %d" % (ym_label(ym), p, g, psp))
        if abs(v("SERVICE", ym) - psp - gov) > TOL_ID:
            raise Stop("%s: Total - Goods != Private service-providing + Government" % ym_label(ym))
    # components that identities do not cover
    prevrem_g = prevrem_s = None
    for ym in [M2, M1, M]:
        rem_g = v("GOODS", ym) - v("CONSTRUCTION", ym) - v("MANUFACTURING", ym)
        rem_s = v("PSP", ym) - v("TTU", ym) - v("FINANCIAL", ym) - v("PBS", ym)
        if not (400 < rem_g < 900):
            raise Stop("%s: Goods - Construction - Manufacturing = %d, outside 400..900 (mining and logging)"
                       % (ym_label(ym), rem_g))
        if prevrem_g is not None:
            if abs(rem_g - prevrem_g) > 15:
                raise Stop("%s: Goods - Construction - Manufacturing moved %d (limit 15); check Goods, Construction, Manufacturing"
                           % (ym_label(ym), rem_g - prevrem_g))
            if abs(rem_s - prevrem_s) > 250:
                raise Stop("%s: Private service-providing - TTU - Financial - PBS moved %d (limit 250); check TTU, Financial, PBS"
                           % (ym_label(ym), rem_s - prevrem_s))
        prevrem_g, prevrem_s = rem_g, rem_s
 
    # --- headline and revisions tie to the published text
    if v("TOTAL", M) - v("TOTAL", M1) != meta[("HEADLINE_CHANGE", M)]:
        raise Stop("Total nonfarm %s - %s = %d but the published headline change is %d" % (
            ym_str(M), ym_str(M1), v("TOTAL", M) - v("TOTAL", M1), meta[("HEADLINE_CHANGE", M)]))
    if v("TOTAL", M1) - v("TOTAL", M2) != meta[("REVISED_CHANGE", M1)]:
        raise Stop("Total nonfarm revised change for %s: levels give %d, published %d" % (
            ym_str(M1), v("TOTAL", M1) - v("TOTAL", M2), meta[("REVISED_CHANGE", M1)]))
    T = bk.sheets["Total_Nonfarm_Payrolls"]
    tm = T.month_map()
    M3 = ym_add(M, -3)
    wb_t = {ym: T.num(tm[ym], "B") for ym in (M1, M2, M3) if ym in tm}
    if wb_t[M1] - wb_t[M2] != meta[("ORIGINAL_CHANGE", M1)]:
        if accept_prev and wb_t[M1] - wb_t[M2] == meta[("REVISED_CHANGE", M1)]:
            warns.append("--accept-previous-mismatch used: the workbook's %s change (%d) is the REVISED figure, not the "
                         "originally published %d. Accepted for this run only." % (
                             ym_str(M1), wb_t[M1] - wb_t[M2], meta[("ORIGINAL_CHANGE", M1)]))
        else:
            raise Stop("workbook %s change is %d but the release says it was originally %d: the workbook does not hold "
                   "the previous release" % (ym_str(M1), wb_t[M1] - wb_t[M2], meta[("ORIGINAL_CHANGE", M1)]))
    if M3 in wb_t and v("TOTAL", M2) - wb_t[M3] != meta[("REVISED_CHANGE", M2)]:
        warns.append("%s: the release says the revised change is %d, but %s minus the workbook's %s gives %d. "
                     "The workbook's %s total (%d) differs from the one the release implies (%d). Not changed."
                     % (ym_str(M2), meta[("REVISED_CHANGE", M2)], ym_str(M2), ym_str(M3),
                        v("TOTAL", M2) - wb_t[M3], ym_str(M3), wb_t[M3],
                        v("TOTAL", M2) - meta[("REVISED_CHANGE", M2)]))
 
    # --- year-ago ties
    for name, key, kind in SHEETS:
        if kind == "rate":
            continue
        sh = bk.sheets[name]
        r = sh.month_map().get(M12)
        if r is None:
            raise Stop("%s has no %s row to compare the year-ago level with" % (name, ym_label(M12)))
        if sh.num(r, "B") != v(key, M12):
            raise Stop("%s %s: workbook %s but the release's year-ago column shows %s (annual benchmark? "
                       "refresh history before continuing)" % (name, ym_label(M12), fmt(sh.num(r, "B")), fmt(v(key, M12))))
 
    # --- unemployment rate
    ur = bk.sheets["Unemployment_Rate"]
    um = ur.month_map()
    for ym in (M1, M2):
        if abs(ur.num(um[ym], "B") - v("UR", ym)) > 1e-9:
            raise Stop("Unemployment rate %s: workbook %.1f vs release %.1f (it is only revised once a year)" % (
                ym_label(ym), ur.num(um[ym], "B"), v("UR", ym)))
    if not 0 < v("UR", M) < 20 or abs(v("UR", M) - ur.num(um[latest], "B")) > 0.5:
        raise Stop("Unemployment rate %.1f looks implausible after %.1f" % (v("UR", M), ur.num(um[latest], "B")))
 
    # --- revisions to prior levels (overwrite the two prior months)
    revlog = []
    for name, key, kind in SHEETS:
        if kind == "rate":
            continue
        sh = bk.sheets[name]
        mm = sh.month_map()
        for ym in (M2, M1):
            old, new = sh.num(mm[ym], "B"), v(key, ym)
            if abs(new - old) > MAX_REVISION:
                raise Stop("%s %s: revision %d is larger than %d; check the figure" % (name, ym_label(ym), new - old, MAX_REVISION))
 
    # --- apply
    before = snapshot(bk)
    out = []
    for name, key, kind in SHEETS:
        sh = bk.sheets[name]
        mm = sh.month_map()
        if kind != "rate":
            for ym in (M2, M1):
                old, new = sh.num(mm[ym], "B"), v(key, ym)
                if old != new:
                    sh.set_num(mm[ym], "B", new)
                    revlog.append((name, ym_label(ym), int(old), int(new), int(new - old)))
        sh.insert_rows(2, 1, 2)
        sh.set_num(2, "A", serial(M))
        sh.set_num(2, "B", v(key, M))
        mm = sh.month_map()
        for ym in (M, M1, M2):
            em, ey = sh.expected(kind, mm[ym], mm)
            for col, e in (("C", em), ("D", ey)):
                s = sh.num(mm[ym], col)
                if ym == M or kind != "rate" or s is None or e is None or abs(s - e) > EPS:
                    sh.set_num(mm[ym], col, e)
        em, ey = sh.expected(kind, 2, mm)
        if em is not None and kind != "rate" and abs(em) > MAX_MOM:
            raise Stop("%s %s: monthly move %.2f%% is larger than %.0f%%; check the figure" % (
                name, ym_label(M), em * 100, MAX_MOM * 100))
        out.append((name, v(key, M), em, ey))
    bk.refresh_charts()
 
    # --- proof: history unchanged except the listed revisions
    after = snapshot(bk)
    allowed = {M1, M2}
    for name, key, kind in SHEETS:
        if set(after[name]) != set(before[name]) | {M}:
            raise Stop("internal error: %s months changed unexpectedly" % name)
        for ym, cols in before[name].items():
            if ym in allowed:
                if kind == "rate" and cols != after[name][ym]:
                    raise Stop("internal error: unemployment history changed")
                continue
            if cols != after[name][ym]:
                raise Stop("internal error: history changed in %s %s" % (name, ym_label(ym)))
    return M, out, revlog, warns, anomalies
 
 
# ---------------------------------------------------------------- main
def sha(b):
    return hashlib.sha256(b).hexdigest()
 
 
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("values", nargs="?")
    ap.add_argument("--restate", action="store_true", help="one-time clean-up of the existing workbook")
    ap.add_argument("--nov25-ur", type=float, help="Nov-25 unemployment rate, with --restate")
    ap.add_argument("--in-place", action="store_true", help="make a backup in backups/, then replace the workbook")
    ap.add_argument("--accept-previous-mismatch", action="store_true",
                    help="one run only: accept a workbook that already holds the revised prior-month change")
    ap.add_argument("--out", help="write the result here and leave the workbook alone")
    a = ap.parse_args()
    try:
        if bool(a.restate) == bool(a.values):
            raise Stop("give either values.csv (monthly update) or --restate (one time), not both/neither")
        if bool(a.in_place) == bool(a.out):
            raise Stop("give exactly one of --in-place or --out FILE")
        orig = open(a.workbook, "rb").read()
        bk = Book(a.workbook)
        n_charts = bk.n_charts()
        if n_charts != 22:
            raise Stop("expected 22 charts, found %d" % n_charts)
        if a.restate:
            log = restate(bk, a.nov25_ur)
            anomalies, latest = audit_workbook(bk)
            tag = "restate"
        else:
            vals, meta = read_csv(a.values)
            M, rows, revlog, warns, anomalies = update(bk, vals, meta, a.accept_previous_mismatch)
            tag = ym_str(M)
        parts = bk.render()
        # final proof on the rendered file (written to a temp file next to the target)
        tmp = bk.write(a.out if a.out else a.workbook, parts)
        try:
            chk = Book(tmp)
            if chk.n_charts() != 22 or set(chk.parts) != set(bk.parts):
                raise Stop("internal error: output is missing charts or parts")
        except Exception:
            os.remove(tmp)
            raise
        if a.in_place:
            bdir = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), "backups")
            stem = os.path.splitext(os.path.basename(a.workbook))[0]
            bpath = os.path.join(bdir, "%s_before_%s.xlsx" % (stem, tag))
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
 
    if a.restate:
        print("\nRESTATEMENT DONE: %d changes (levels untouched)" % len(log))
        for l in log:
            print("  %-32s %-8s %-10s %s -> %s" % l)
    else:
        print("\nUPDATE OK: %s added to all 11 sheets" % ym_label(M))
        print("%-32s %10s %9s %9s" % ("sheet", "value", "MoM", "YoY"))
        for name, lv, mom, yoy in rows:
            f = lambda x: "-" if x is None else "%.2f%%" % (x * 100)
            print("%-32s %10s %9s %9s" % (name, fmt(lv), f(mom), f(yoy)))
        print("\nRevisions applied to prior months (old -> new):")
        for r in revlog:
            print("  %-32s %-7s %s -> %s (%+d)" % r)
        if not revlog:
            print("  none")
        if warns:
            print("\nWarnings:")
            for w in warns:
                print("  - " + w)
    print("\nAnomaly audit of existing data (not changed): %d" % len(anomalies))
    for x in anomalies[:25]:
        print("  - " + x)
    print("\nCharts: 22 kept, ranges and caches rebuilt.")
 
 
if __name__ == "__main__":
    main()