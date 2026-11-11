"""Add the next month's row to the US CPI workbook (10 sheets, newest month at the top).
 
Usage:
    python cpi_update_workbook.py US_CPI.xlsx values.csv [--in-place | --out new.xlsx]
                                  [--accept-mismatch SHEET[,SHEET]]
 
values.csv has one line per sheet (see cpi_conventions.md):
    sheet,month,prev_index,index,year_ago_index,bls_12m_pct
  month          YYYY-MM being added (must be the month after the sheet's newest row)
  prev_index     the source's index for the PREVIOUS month (used for MoM and as a check)
  index          the source's index for the new month
  year_ago_index the source's index for the same month a year earlier (used for YoY)
  bls_12m_pct    optional: the published 12-month % change, as a cross-check
 
The workbook is edited at file level (new row inserted at row 2 of every sheet, chart ranges
extended by one row), so the 20 charts, formatting and everything else are kept as they are.
MoM = index / prev_index - 1 and YoY = index / year_ago_index - 1 (not seasonally adjusted).
 
Checks (any failure = STOP, nothing written): all sheets end on the same month; the new month is
the next one; prev_index equals the workbook's newest index (otherwise the series changed, e.g.
a different index was used earlier; override one sheet at a time with --accept-mismatch);
computed YoY rounds to the published 12-month % when given. Existing history is never changed.
"""
import argparse
import csv
import datetime
import os
import re
import shutil
import sys
import zipfile
 
EPOCH = datetime.datetime(1899, 12, 30)
SHEET_ORDER = ["CPI_All", "CPI_All_Less_Food_Energy", "Food_Beverages", "Energy", "Housing",
               "Apparel", "Transportation", "Medical_Care", "Recreation", "Education"]
 
 
def serial(d):
    return (d - EPOCH).days
 
 
def from_serial(n):
    return EPOCH + datetime.timedelta(days=int(float(n)))
 
 
def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        i, t = re.search(r'Id="([^"]+)"', tag), re.search(r'Target="([^"]+)"', tag)
        if i and t:
            target[i.group(1)] = t.group(1)
    out = {}
    for m in re.finditer(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        out[m.group(1)] = "xl/" + target[m.group(2)].lstrip("/").replace("xl/", "")
    return out
 
 
def top_row(xml):
    """Return (row 2 xml, date, index) of the newest data row."""
    m = re.search(r'<row r="2"[^>]*>.*?</row>', xml, re.S)
    if not m:
        raise ValueError("no data in row 2")
    row = m.group(0)
    a = re.search(r'<c r="A2"[^>]*><v>([^<]+)</v>', row)
    b = re.search(r'<c r="B2"[^>]*><v>([^<]+)</v>', row)
    if not a or not b:
        raise ValueError("row 2 does not hold a date in A and an index in B")
    return row, from_serial(a.group(1)), float(b.group(1))
 
 
def next_month(d):
    return datetime.datetime(d.year + (d.month == 12), d.month % 12 + 1, 1)
 
 
def new_row_xml(row2, date, idx, mom, yoy):
    def style(cell):
        m = re.search(r'<c r="' + cell + r'2"([^>]*)>', row2)
        s = re.search(r'\bs="(\d+)"', m.group(1)) if m else None
        return f' s="{s.group(1)}"' if s else ""
    head = re.match(r'<row r="2"([^>]*)>', row2).group(1)
    return (f'<row r="2"{head}><c r="A2"{style("A")}><v>{serial(date)}</v></c>'
            f'<c r="B2"{style("B")}><v>{idx!r}</v></c><c r="C2"{style("C")}><v>{mom!r}</v></c>'
            f'<c r="D2"{style("D")}><v>{yoy!r}</v></c></row>')
 
 
def shift_rows(xml, row2_new):
    def bump_row(m):
        n = int(m.group(1))
        return f'<row r="{n + 1}"' if n >= 2 else m.group(0)
 
    def bump_cell(m):
        n = int(m.group(2))
        return f'<c r="{m.group(1)}{n + 1}"' if n >= 2 else m.group(0)
    head, rest = xml.split("<sheetData>", 1)
    data, tail = rest.split("</sheetData>", 1)
    data = re.sub(r'<row r="(\d+)"', bump_row, data)
    data = re.sub(r'<c r="([A-Z]+)(\d+)"', bump_cell, data)
    data = data.replace('<row r="3"', row2_new + '<row r="3"', 1)
    head = re.sub(r'<dimension ref="A1:([A-Z]+)(\d+)"/>',
                  lambda m: f'<dimension ref="A1:{m.group(1)}{int(m.group(2)) + 1}"/>', head)
    return head + "<sheetData>" + data + "</sheetData>" + tail
 
 
def extend_chart(xml, sheet):
    n = 0
 
    def f(m):
        nonlocal n
        n += 1
        return f"{m.group(1)}${m.group(2)}$2:${m.group(2)}${int(m.group(3)) + 1}"
    new = re.sub(r"(?<![\w])(" + re.escape(sheet) + r"!)\$([A-D])\$2:\$\2\$(\d+)", f, xml)
    return new, n
 
 
def read_values(path):
    rows = {}
    with open(path, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            rows[r["sheet"].strip()] = r
    return rows
 
 
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("values")
    dest = ap.add_mutually_exclusive_group()
    dest.add_argument("--in-place", action="store_true")
    dest.add_argument("--out")
    ap.add_argument("--accept-mismatch", default="", help="comma list of sheets whose prev_index may differ from the workbook")
    a = ap.parse_args()
    accept = {s.strip() for s in a.accept_mismatch.split(",") if s.strip()}
 
    z = zipfile.ZipFile(a.workbook)
    paths = sheet_paths(z)
    vals = read_values(a.values)
    if set(paths) != set(SHEET_ORDER):
        sys.exit(f"STOP: workbook sheets are not the expected ten: {sorted(paths)}")
    if set(vals) != set(SHEET_ORDER):
        sys.exit("STOP: values file must have exactly one line for each of: " + ", ".join(
            s for s in SHEET_ORDER if s not in vals) + (" (extra: " + ", ".join(set(vals) - set(SHEET_ORDER)) + ")" if set(vals) - set(SHEET_ORDER) else ""))
 
    plan, notes, warns = {}, [], []
    new_months = set()
    for sheet in SHEET_ORDER:
        xml = z.read(paths[sheet]).decode("utf-8")
        try:
            row2, date, cur = top_row(xml)
        except ValueError as e:
            sys.exit(f"STOP: {sheet}: {e}")
        v = vals[sheet]
        try:
            month = datetime.datetime.strptime(v["month"].strip(), "%Y-%m")
            prev, idx, ago = (float(v[k]) for k in ("prev_index", "index", "year_ago_index"))
        except (ValueError, KeyError, TypeError) as e:
            sys.exit(f"STOP: {sheet}: bad line in values file ({e})")
        newest = datetime.datetime(date.year, date.month, 1)
        if month != next_month(newest):
            sys.exit(f"STOP: {sheet} ends at {newest:%b-%y}, so the month to add is "
                     f"{next_month(newest):%Y-%m}, not {v['month']}")
        new_months.add(month)
        if abs(prev - cur) > 0.0005:
            msg = (f"{sheet}: the source's {newest:%b-%y} index is {prev}, but the workbook has {cur}")
            if sheet in accept:
                warns.append(msg + " (accepted by you; MoM uses the source's figure)")
            else:
                sys.exit("STOP: " + msg + ". Either the source was read wrongly or the series in the "
                         "workbook differs; nothing was written.")
        mom, yoy = idx / prev - 1, idx / ago - 1
        pub = (v.get("bls_12m_pct") or "").strip()
        if pub and abs(yoy * 100 - float(pub)) > 0.051:
            sys.exit(f"STOP: {sheet}: computed YoY {yoy*100:.2f}% differs from the published {pub}%")
        plan[sheet] = (xml, row2, month, idx, mom, yoy)
        notes.append((sheet, idx, mom, yoy))
    month = new_months.pop() if len(new_months) == 1 else sys.exit("STOP: sheets end on different months")
 
    out_parts = {}
    for sheet, (xml, row2, month, idx, mom, yoy) in plan.items():
        out_parts[paths[sheet]] = shift_rows(xml, new_row_xml(row2, month, idx, mom, yoy)).encode("utf-8")
    charts = 0
    for name in z.namelist():
        if re.match(r"xl/charts/chart\d+\.xml$", name):
            xml = z.read(name).decode("utf-8")
            for sheet in SHEET_ORDER:
                xml, n = extend_chart(xml, sheet)
                charts += n
            out_parts[name] = xml.encode("utf-8")
 
    if a.in_place:
        dst = a.workbook + ".tmp"
    else:
        dst = a.out or os.path.splitext(os.path.basename(a.workbook))[0] + f"_{month:%Y-%m}.xlsx"
        if os.path.abspath(dst) == os.path.abspath(a.workbook):
            sys.exit("STOP: refusing to overwrite the input workbook (use --in-place)")
    with zipfile.ZipFile(dst, "w") as zo:
        for info in z.infolist():
            data = out_parts.get(info.filename, z.read(info.filename))
            zo.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    z.close()
    if a.in_place:
        folder = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), "backups")
        os.makedirs(folder, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(a.workbook))
        backup, n = os.path.join(folder, f"{stem}_before_{month:%Y-%m}{ext}"), 2
        while os.path.exists(backup):
            backup = os.path.join(folder, f"{stem}_before_{month:%Y-%m}_{n}{ext}")
            n += 1
        shutil.copy2(a.workbook, backup)
        os.replace(dst, a.workbook)
        print("Backup of the original:", backup)
        print("Updated in place:", a.workbook)
    else:
        print("Saved:", dst)
    print(f"Added {month:%b-%y} as row 2 of all 10 sheets; {charts} chart ranges extended by one row.\n")
    print(f"{'sheet':26}{'index':>10}{'MoM':>9}{'YoY':>9}")
    for sheet, idx, mom, yoy in notes:
        print(f"{sheet:26}{idx:>10.3f}{mom*100:>8.2f}%{yoy*100:>8.2f}%")
    for w in warns:
        print("WARNING:", w)
 
 
if __name__ == "__main__":
    main()