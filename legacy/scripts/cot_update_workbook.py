#!/usr/bin/env python3
"""
Commitment of Traders - Metals & Energy: weekly workbook update.
 
Usage:
  python cot_update_workbook.py <workbook.xlsx> <f_year.xls> [<older_year.xls> ...]
         (--in-place | --out NEW.xlsx) [--extract-csv values.csv]
 
What it does (see cot_conventions.md):
  1. Reads the CFTC Disaggregated Futures-Only year file (.xls/.xlsx/.csv).
  2. Keeps the 11 contracts (by CFTC_Contract_Market_Code) and 7 columns.
  3. Checks them against the workbook (overlapping weeks must be identical).
  4. Appends every week after the workbook's last date to 'Data - Energy & Metals'
     and extends columns A:E of the 11 commodity sheets by the same rows.
  5. Refreshes the cached values on the Summary sheet; Excel recalculates on open.
Anything unexpected prints STOP and nothing is written (exit code 2).
The workbook is edited at XML level, so charts, formatting and conditional formats are kept.
"""
import argparse, csv, datetime as dt, os, re, shutil, sys, zipfile
 
# ---------------------------------------------------------------- configuration
# Order = column order of the data sheet (3 columns per commodity from column B).
CONTRACTS = [
    ("Crude Oil WTI",   "067411", "CRUDE OIL, LIGHT SWEET-WTI"),
    ("Crude Oil Brent", "06765T", "BRENT LAST DAY"),
    ("Heating Oil",     "022651", "NY HARBOR ULSD"),
    ("Palladium",       "075651", "PALLADIUM"),
    ("Platinum",        "076651", "PLATINUM"),
    ("Silver",          "084691", "SILVER"),
    ("Gold",            "088691", "GOLD"),
    ("Copper",          "085692", "COPPER"),
    ("Aluminium",       "191693", "ALUMINUM"),
    ("Steel",           "192651", "STEEL-HRC"),
    ("Natural Gas",     "023651", "NAT GAS NYME"),
]
SRC_COLS = ["Market_and_Exchange_Names", "Report_Date_as_MM_DD_YYYY", "CFTC_Contract_Market_Code",
            "Open_Interest_All", "M_Money_Positions_Long_ALL", "M_Money_Positions_Short_ALL",
            "Open_Interest_Old"]
DATA_SHEET = "Data - Energy & Metals"
SUMMARY_SHEET = "Energy & Metals Summary"
FIRST_DATA_ROW = 3
SUMMARY_FIRST, SUMMARY_LAST = 4, 28
OI_JUMP_WARN = 0.30          # warn (not stop) when open interest moves more than this in a week
 
 
class Stop(Exception):
    pass
 
 
def stop(msg):
    raise Stop(msg)
 
 
def col_letter(n):               # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s
 
 
def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n
 
 
def serial(d):
    return (d - dt.date(1899, 12, 30)).days
 
 
def from_serial(n):
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(n))
 
 
# ---------------------------------------------------------------- reading the CFTC file
def read_source_rows(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
    else:
        try:
            from python_calamine import CalamineWorkbook
        except ImportError:
            stop("python_calamine is not installed, so the .xls file cannot be read. "
                 "Save the file as .xlsx or .csv and run again.")
        wb = CalamineWorkbook.from_path(path)
        rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python()
    if not rows:
        stop(f"{path} is empty.")
    return rows
 
 
def to_date(v, where):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    stop(f"{where}: cannot read the report date {v!r}.")
 
 
def to_num(v, where):
    try:
        f = float(str(v).replace(",", "").strip())
    except ValueError:
        stop(f"{where}: {v!r} is not a number.")
    if f != int(f):
        stop(f"{where}: {v!r} is not a whole number.")
    return int(f)
 
 
def extract(paths):
    """Return {code: {date: dict(name, oi, long, short, oi_old)}} for the 11 contracts."""
    codes = {c[1]: c for c in CONTRACTS}
    out = {c[1]: {} for c in CONTRACTS}
    for p in paths:
        rows = read_source_rows(p)
        head = [str(h).strip() for h in rows[0]]
        missing = [c for c in SRC_COLS if c not in head]
        if missing:
            stop(f"{os.path.basename(p)}: column(s) not found: {', '.join(missing)}. "
                 "Is this the Disaggregated Futures-Only report?")
        ix = {h: i for i, h in enumerate(head)}
        fo = ix.get("FutOnly_or_Combined")
        for n, r in enumerate(rows[1:], start=2):
            code = str(r[ix["CFTC_Contract_Market_Code"]]).strip()
            if code not in codes:
                continue
            where = f"{os.path.basename(p)} row {n} ({codes[code][0]})"
            if fo is not None and str(r[fo]).strip() != "FutOnly":
                stop(f"{where}: FutOnly_or_Combined is {r[fo]!r}, expected 'FutOnly'. "
                     "Download the Futures Only file, not Futures-and-Options Combined.")
            name = str(r[ix["Market_and_Exchange_Names"]]).strip()
            if codes[code][2] not in name.upper():
                stop(f"{where}: contract name {name!r} does not contain {codes[code][2]!r}. "
                     "The code may have been reassigned; check before continuing.")
            d = to_date(r[ix["Report_Date_as_MM_DD_YYYY"]], where)
            rec = dict(name=name,
                       oi=to_num(r[ix["Open_Interest_All"]], where),
                       long=to_num(r[ix["M_Money_Positions_Long_ALL"]], where),
                       short=to_num(r[ix["M_Money_Positions_Short_ALL"]], where),
                       oi_old=to_num(r[ix["Open_Interest_Old"]], where))
            if d in out[code]:
                if out[code][d] != rec:
                    stop(f"{where}: {d} appears twice with different values.")
                continue                      # same week in two year files: fine
            out[code][d] = rec
    return out
 
 
def validate_source(src):
    dates = None
    for name, code, _ in CONTRACTS:
        if not src[code]:
            stop(f"{name} ({code}) not found in the source file.")
        ds = set(src[code])
        if dates is None:
            dates = ds
        elif ds != dates:
            diff = sorted(ds ^ dates)
            stop(f"{name} ({code}) has a different set of report dates from the other contracts "
                 f"(difference: {', '.join(map(str, diff[:6]))}).")
    for name, code, _ in CONTRACTS:
        for d, r in sorted(src[code].items()):
            w = f"{name} {d}"
            if r["oi"] <= 0:
                stop(f"{w}: open interest {r['oi']} is not positive.")
            if min(r["long"], r["short"], r["oi_old"]) < 0:
                stop(f"{w}: negative value in the source.")
            if r["long"] > r["oi"] or r["short"] > r["oi"]:
                stop(f"{w}: money-manager long/short ({r['long']}/{r['short']}) exceeds open interest {r['oi']}.")
            if r["oi_old"] > r["oi"]:
                stop(f"{w}: Open_Interest_Old {r['oi_old']} exceeds Open_Interest_All {r['oi']}.")
    return sorted(dates)
 
 
# ---------------------------------------------------------------- reading the workbook
def read_zip(path):
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        return infos, {i.filename: z.read(i.filename) for i in infos}
 
 
def sheet_files(parts):
    wbx = parts["xl/workbook.xml"].decode("utf8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        t = m.group(0)
        rid2t[re.search(r'Id="([^"]+)"', t).group(1)] = re.search(r'Target="([^"]+)"', t).group(1)
    out = {}
    for m in re.finditer(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        name = m.group(1).replace("&amp;", "&")
        out[name] = "xl/" + rid2t[m.group(2)].lstrip("/").replace("xl/", "")
    return out
 
 
ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
 
 
def parse_rows(xml):
    """{row: {col_letters: (attrs, inner)}}"""
    rows = {}
    for m in ROW_RE.finditer(xml):
        cells = {}
        for c in CELL_RE.finditer(m.group(0)):
            cells[c.group(1)] = (c.group(3), c.group(4) or "")
        rows[int(m.group(1))] = cells
    return rows
 
 
def cell_number(cells, col):
    if col not in cells:
        return None
    attrs, inner = cells[col]
    if "t=" in attrs:
        return None
    m = re.search(r"<v>([^<]*)</v>", inner)
    return float(m.group(1)) if m else None
 
 
def load_sst(parts):
    shared = parts["xl/sharedStrings.xml"].decode("utf8")
    return [re.sub(r"<[^>]+>", "", x) for x in re.findall(r"<si>.*?</si>", shared, re.S)]
 
 
def read_master(parts, files):
    xml = parts[files[DATA_SHEET]].decode("utf8")
    rows = parse_rows(xml)
    # header check
    hdr1 = rows.get(1, {})
    shared = parts["xl/sharedStrings.xml"].decode("utf8")
    sst = [re.sub(r"<[^>]+>", "", s) for s in re.findall(r"<si>.*?</si>", shared, re.S)]
 
    def text(cells, col):
        if col not in cells:
            return None
        attrs, inner = cells[col]
        m = re.search(r"<v>([^<]*)</v>", inner)
        if not m:
            return None
        return sst[int(m.group(1))].replace("&amp;", "&") if 't="s"' in attrs else m.group(1)
 
    for k, (name, _, _) in enumerate(CONTRACTS):
        col = col_letter(2 + 3 * k)
        if text(hdr1, col) != name:
            stop(f"Data sheet cell {col}1 is {text(hdr1, col)!r}, expected {name!r}. "
                 "The workbook layout has changed; the script was not told about it.")
    hdr2 = rows.get(2, {})
    if text(hdr2, "A") != "Date":
        stop("Data sheet A2 is not 'Date'.")
    data = {}
    last = FIRST_DATA_ROW - 1
    for r in sorted(rows):
        if r < FIRST_DATA_ROW:
            continue
        v = cell_number(rows[r], "A")
        if v is None:
            if any(cell_number(rows[r], col_letter(c)) is not None for c in range(2, 35)):
                stop(f"Data sheet row {r} has values but no date.")
            continue
        d = from_serial(v)
        if d in data:
            stop(f"Data sheet: date {d} appears twice (row {data[d][0]} and row {r}).")
        vals = []
        for c in range(2, 35):
            n = cell_number(rows[r], col_letter(c))
            vals.append(None if n is None else int(n))
        data[d] = (r, vals)
        last = max(last, r)
    if not data:
        stop("Data sheet has no data rows.")
    ds = [d for d, _ in sorted(data.items(), key=lambda kv: kv[1][0])]
    if ds != sorted(ds):
        stop("Data sheet dates are not in ascending order.")
    # nothing may sit below the last dated row
    for r in rows:
        if r > last and any(cell_number(rows[r], col_letter(c)) is not None for c in range(1, 35)):
            stop(f"Data sheet row {r} holds values below the last dated row {last}.")
    return data, last, rows
 
 
# ---------------------------------------------------------------- compare / plan
def plan(src, src_dates, master):
    data, last_row, _ = master
    last_date = max(data)
    problems = []
    for d in src_dates:
        if d > last_date:
            continue
        if d not in data:
            # a source week that falls inside the workbook's range but is not in it
            if d >= min(data):
                problems.append(f"{d}: in the source but not in the workbook (workbook ends {last_date}).")
            continue
        row, vals = data[d]
        for k, (name, code, _) in enumerate(CONTRACTS):
            s = src[code][d]
            want = [s["oi"], s["long"], s["short"]]
            have = vals[3 * k:3 * k + 3]
            if have != want:
                problems.append(f"{d} {name}: workbook row {row} has {have}, source has {want}.")
    if problems:
        stop("Weeks in both files do not match (OI, longs, shorts):\n  " + "\n  ".join(problems[:20]) +
             (f"\n  ... and {len(problems) - 20} more" if len(problems) > 20 else "") +
             "\nNo history is changed silently. Tell the owner; do not hand-edit.")
    if last_date not in src_dates:
        stop(f"The workbook's last week ({last_date}) is not in the source file, so continuity cannot be "
             "checked. Add the previous year's file as a further argument.")
    new = [d for d in src_dates if d > last_date]
    prev = last_date
    for d in new:
        gap = (d - prev).days
        if not 5 <= gap <= 9:
            stop(f"Gap of {gap} days between {prev} and {d}: a week is missing or the dates are odd.")
        prev = d
    return new, last_date
 
 
def warnings_for(src, new, last_date):
    w = []
    for name, code, _ in CONTRACTS:
        prev = src[code][last_date]
        for d in new:
            cur = src[code][d]
            if prev["oi"] and abs(cur["oi"] / prev["oi"] - 1) > OI_JUMP_WARN:
                w.append(f"{name} {d}: open interest {prev['oi']:,} -> {cur['oi']:,} "
                         f"({cur['oi'] / prev['oi'] - 1:+.0%}); likely a contract roll, check.")
            prev = cur
    return w
 
 
# ---------------------------------------------------------------- writing XML
def num(v):
    return repr(float(v)) if isinstance(v, float) else str(v)
 
 
def set_row(xml, r, new_cells, template_row_attrs, replace_cols):
    """Insert row r, or fill the empty styled row that already exists."""
    m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % r, xml, re.S)
    if m:
        old = parse_rows(m.group(0))[r]
        for col in replace_cols:
            if col in old and re.search(r"<v>|<f", old[col][1]):
                stop(f"Row {r} already has content in {col}{r}; refusing to overwrite.")
        row_xml = m.group(0)
        for col, cell_xml in new_cells.items():
            pat = re.compile(r'<c r="%s%d"[^>]*?(?:/>|>.*?</c>)' % (col, r), re.S)
            if not pat.search(row_xml):
                stop(f"Row {r} exists but has no {col}{r} cell to fill.")
            row_xml = pat.sub(lambda _m, c=cell_xml: c, row_xml, count=1)
        return xml[:m.start()] + row_xml + xml[m.end():]
    cells = "".join(new_cells[c] for c in sorted(new_cells, key=col_index))
    row_xml = f'<row r="{r}"{template_row_attrs}>{cells}</row>'
    after = [int(x.group(1)) for x in ROW_RE.finditer(xml) if int(x.group(1)) > r]
    if after:
        pos = xml.index('<row r="%d"' % min(after))
        return xml[:pos] + row_xml + xml[pos:]
    pos = xml.index("</sheetData>")
    return xml[:pos] + row_xml + xml[pos:]
 
 
def row_attrs(xml, r):
    m = re.search(r'<row r="%d"([^>]*?)/?>' % r, xml)
    return m.group(1) if m else ""
 
 
def style_of(rows, r, col):
    m = re.search(r's="(\d+)"', rows[r][col][0]) if r in rows and col in rows[r] else None
    return m.group(1) if m else None
 
 
def set_dimension(xml, last_row):
    def rep(m):
        a, b = m.group(1), m.group(2)
        col = re.match(r"[A-Z]+", b).group(0)
        cur = int(re.search(r"\d+", b).group(0))
        return f'<dimension ref="{a}:{col}{max(cur, last_row)}"/>'
    return re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+\d+)"/>', rep, xml, count=1)
 
 
def write_data_sheet(xml, master, last_row, new, src):
    rows = parse_rows(xml)
    tmpl = last_row
    sa = style_of(rows, tmpl, "A")
    attrs = row_attrs(xml, tmpl)
    r = last_row
    for d in new:
        r += 1
        cells = {"A": f'<c r="A{r}" s="{sa}"><v>{serial(d)}</v></c>'}
        for k, (_, code, _) in enumerate(CONTRACTS):
            s = src[code][d]
            for j, key in enumerate(("oi", "long", "short")):
                col = col_letter(2 + 3 * k + j)
                st = style_of(rows, tmpl, col)
                st_attr = f' s="{st}"' if st else ""
                cells[col] = f'<c r="{col}{r}"{st_attr}><v>{s[key]}</v></c>'
        xml = set_row(xml, r, cells, attrs, list(cells))
    return set_dimension(xml, r)
 
 
def write_commodity_sheet(xml, name, code, last_row, new, src, data_cols):
    rows = parse_rows(xml)
    if last_row not in rows or "A" not in rows[last_row] or "<f>" not in rows[last_row]["A"][1]:
        stop(f"Sheet {name!r}: row {last_row} has no link formula in column A to copy.")
    tmpl = rows[last_row]
    for col in "BCD":
        if col not in tmpl or "<f>" not in tmpl[col][1]:
            stop(f"Sheet {name!r}: {col}{last_row} has no formula to copy.")
    attrs = row_attrs(xml, last_row)
    r = last_row
    for d in new:
        r += 1
        s = src[code][d]
        cells = {}
        for col, val in (("A", serial(d)), ("B", s["oi"]), ("C", s["long"]), ("D", s["short"])):
            a, inner = tmpl[col]
            f = re.search(r"<f>.*?</f>", inner, re.S).group(0)
            f = re.sub(r"(?<![0-9])%d(?![0-9])" % last_row, str(r), f)
            st = re.search(r's="(\d+)"', a)
            st_attr = f' s="{st.group(1)}"' if st else ""
            cells[col] = f'<c r="{col}{r}"{st_attr}>{f}<v>{val}</v></c>'
        a, _ = tmpl["E"]
        st = re.search(r's="(\d+)"', a)
        st_attr = f' s="{st.group(1)}"' if st else ""
        e = (s["long"] - s["short"]) / s["oi"]
        cells["E"] = f'<c r="E{r}"{st_attr}><f>(C{r}-D{r})/B{r}</f><v>{e!r}</v></c>'
        xml = set_row(xml, r, cells, attrs, list(cells))
    return set_dimension(xml, r)
 
 
def refresh_summary(xml, data_after, sst):
    """Update cached values on the Summary sheet (Excel recalculates anyway).
    The commodity order on the Summary sheet is read from its row-3 headers."""
    hdr = parse_rows(xml).get(3, {})
    pairs = []                                   # (first column letter, commodity name)
    for col, (attrs, inner) in hdr.items():
        m = re.search(r"<v>([^<]*)</v>", inner)
        if m and 't="s"' in attrs:
            nm = sst[int(m.group(1))].replace("&amp;", "&")
            if nm in {c[0] for c in CONTRACTS}:
                pairs.append((col, nm))
    if sorted(n for _, n in pairs) != sorted(c[0] for c in CONTRACTS):
        stop("Summary sheet row 3 does not list the 11 commodities as expected.")
    dates = sorted(data_after, reverse=True)
    out = xml
    for i, r in enumerate(range(SUMMARY_FIRST, SUMMARY_LAST + 1)):
        if i >= len(dates):
            break
        d = dates[i]
        pat = re.compile(r'(<c r="A%d"[^>]*>.*?<v>)[^<]*(</v>)' % r, re.S)
        out = pat.sub(lambda m: m.group(1) + str(serial(d)) + m.group(2), out, count=1)
        for first, nm in pairs:
            oi, lo, sh = data_after[d][nm]
            if not oi or lo is None or sh is None:
                continue
            for col in (first, col_letter(col_index(first) + 1)):
                pat = re.compile(r'(<c r="%s%d"[^>]*>)(.*?)(</c>)' % (col, r), re.S)
                m = pat.search(out)
                if not m:
                    continue
                new_inner = re.sub(r"<v>[^<]*</v>", "<v>%r</v>" % ((lo - sh) / oi), m.group(2))
                head = re.sub(r'\s+t="[a-z]+"', "", m.group(1))
                out = out[:m.start()] + head + new_inner + m.group(3) + out[m.end():]
    return out
 
 
def drop_calc_chain(parts):
    parts.pop("xl/calcChain.xml", None)
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf8")
    rels = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", rels)
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
    ct = parts["[Content_Types].xml"].decode("utf8")
    ct = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', "", ct)
    parts["[Content_Types].xml"] = ct.encode("utf8")
    wb = parts["xl/workbook.xml"].decode("utf8")
    if "fullCalcOnLoad" not in wb:
        wb = re.sub(r"<calcPr\b([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
    parts["xl/workbook.xml"] = wb.encode("utf8")
 
 
def update_charts(parts, last_row):
    """Point every chart series of the 11 commodity sheets at rows 3..last_row (date column A, net % column E).
    Returns the number of chart references changed."""
    names = {c[0] for c in CONTRACTS}
    count = 0
 
    def fix(m):
        nonlocal count
        sheet = m.group(1).strip("'").replace("&amp;", "&")
        if sheet not in names:
            return m.group(0)
        if int(m.group(4)) != last_row:
            count += 1
        return f"<c:f>{m.group(1)}!${m.group(2)}$3:${m.group(3)}${last_row}</c:f>"
 
    for fn in sorted(p for p in parts if re.match(r"xl/charts/chart\d+\.xml$", p)):
        x = parts[fn].decode("utf8")
        y = re.sub(r"<c:f>('[^'!<]+'|[^'!<]+)!\$([AE])\$3:\$([AE])\$(\d+)</c:f>", fix, x)
        if y != x:
            parts[fn] = y.encode("utf8")
    return count
 
 
# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("source", nargs="+")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--in-place", action="store_true")
    g.add_argument("--out")
    ap.add_argument("--extract-csv")
    a = ap.parse_args()
    try:
        src = extract(a.source)
        src_dates = validate_source(src)
        infos, parts = read_zip(a.workbook)
        files = sheet_files(parts)
        for need in [DATA_SHEET, SUMMARY_SHEET] + [c[0] for c in CONTRACTS]:
            if need not in files:
                stop(f"Sheet {need!r} not found in the workbook.")
        master = read_master(parts, files)
        data, last_row, _ = master
        new, last_date = plan(src, src_dates, master)
 
        if a.extract_csv:                                    # the 'filtered' table, all rows of the file
            with open(a.extract_csv, "w", newline="") as f:
                w = csv.writer(f)
                w.writerow(SRC_COLS)
                for name, code, _ in CONTRACTS:
                    for d in sorted(src[code]):
                        s = src[code][d]
                        w.writerow([s["name"], d.strftime("%m/%d/%Y"), code, s["oi"], s["long"], s["short"], s["oi_old"]])
        print(f"Workbook last week: {last_date} (data row {last_row}). Source covers {src_dates[0]} to {src_dates[-1]}.")
        print(f"Overlap check: {sum(1 for d in src_dates if d in data)} weeks x 11 contracts x 3 values equal to the workbook.")
        if not new:
            print("Nothing to add: the workbook already has the latest week. Nothing written.")
            return 0
        print(f"New weeks: {', '.join(map(str, new))}")
        for wmsg in warnings_for(src, new, last_date):
            print("WARNING:", wmsg)
 
        # ---- build the new file in memory
        ds = files[DATA_SHEET]
        parts[ds] = write_data_sheet(parts[ds].decode("utf8"), master, last_row, new, src).encode("utf8")
        for name, code, _ in CONTRACTS:
            fn = files[name]
            parts[fn] = write_commodity_sheet(parts[fn].decode("utf8"), name, code, last_row, new, src, None).encode("utf8")
        # cached summary values
        after = {}
        for d, (_, vals) in data.items():
            after[d] = {n: tuple(vals[3 * k:3 * k + 3]) for k, (n, _, _) in enumerate(CONTRACTS)}
        for d in new:
            after[d] = {n: (src[c][d]["oi"], src[c][d]["long"], src[c][d]["short"]) for n, c, _ in CONTRACTS}
        sf = files[SUMMARY_SHEET]
        parts[sf] = refresh_summary(parts[sf].decode("utf8"), after, load_sst(parts)).encode("utf8")
        drop_calc_chain(parts)
        new_last = last_row + len(new)
        n_ch = update_charts(parts, new_last)
        print(f"Charts: {n_ch} series references set to rows 3 to {new_last} (all 22 charts end at the last data row).")
 
        # ---- write
        outpath = a.workbook if a.in_place else a.out
        if a.in_place:
            os.makedirs("backups", exist_ok=True)
            stem = os.path.splitext(os.path.basename(a.workbook))[0]
            bak = os.path.join("backups", f"{stem}_before_{new[0].isoformat()}.xlsx")
            shutil.copy2(a.workbook, bak)
            print("Backup written:", bak)
        tmp = outpath + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for i in infos:
                if i.filename in parts:
                    z.writestr(i, parts[i.filename], compress_type=zipfile.ZIP_DEFLATED)
        os.replace(tmp, outpath)
        print("Written:", outpath)
 
        print("\nNew rows (Open interest | Mgr long | Mgr short | Net % OI):")
        for name, code, _ in CONTRACTS:
            for d in new:
                s = src[code][d]
                print(f"  {d}  {name:16s} {s['oi']:>10,} {s['long']:>9,} {s['short']:>9,}  {(s['long'] - s['short']) / s['oi']:+.2%}")
        return 0
    except Stop as e:
        print("STOP:", e)
        return 2
 
 
if __name__ == "__main__":
    sys.exit(main())