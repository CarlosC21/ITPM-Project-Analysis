#!/usr/bin/env python3
"""One-off: remove the second copy of a duplicated block of trading days from Bull_Bear_Markets.xlsx.
Usage: python bb_repair_duplicates.py WORKBOOK --out NEW.xlsx
Finds dates that appear twice with identical values in one contiguous block (rows a..b duplicating the block just above),
deletes the lower copy, moves everything below up. Anything else (different values, not contiguous) = STOP.
Same safety approach as bb_update_workbook.py (imports it)."""
import argparse, os, re, sys, zipfile
import bb_update_workbook as u
 
def main():
    ap = argparse.ArgumentParser(); ap.add_argument("workbook"); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    try:
        book = u.Book(a.workbook)
        by = {}
        for r, d, *v in book.data: by.setdefault(d, []).append((r, tuple(v)))
        dup = {d: x for d, x in by.items() if len(x) > 1}
        if not dup: u.stop("No duplicate dates")
        if any(len(x) != 2 or x[0][1] != x[1][1] for x in dup.values()): u.stop("Duplicates with different values or 3 copies")
        low = sorted(x[1][0] for x in dup.values()); hi = sorted(x[0][0] for x in dup.values())
        a1, b1 = low[0], low[-1]
        if low != list(range(a1, b1 + 1)): u.stop("Lower copies are not one contiguous block")
        k = len(low)
        if hi != list(range(hi[0], hi[0] + k)) or hi[-1] + 1 != a1: u.stop("Upper copies are not the block just above the lower copies")
        print("Duplicate block: %d days, kept rows %d-%d, deleting rows %d-%d (%s to %s)" % (
            k, hi[0], hi[-1], a1, b1, book.data[a1 - 2][1], book.data[b1 - 2][1]))
        mp = lambda r: r if r < a1 else r - k
        def remap(text):
            return u.REF_RE.sub(lambda m: m.group(1) + str(mp(int(m.group(2)))), text)
        def fixrow(x):
            x = re.sub(r'<row r="(\d+)"', lambda m: '<row r="%d"' % mp(int(m.group(1))), x, count=1)
            x = re.sub(r'<c r="([A-Z]+)(\d+)"', lambda m: '<c r="%s%d"' % (m.group(1), mp(int(m.group(2)))), x)
            def f(m):
                at = re.sub(r'ref="([^"]+)"', lambda q: 'ref="%s"' % remap(q.group(1)), m.group(1))
                return "<f%s/>" % at if m.group(2) is None else "<f%s>%s</f>" % (at, remap(m.group(2)))
            return re.sub(r"<f([^>]*?)(?:/>|>([^<]*)</f>)", f, x)
        # a deleted row must not hold a shared-formula master that survivors use
        for r, x in book.rows:
            if a1 <= r <= b1 and 't="shared" ref=' in x: u.stop("Deleted row %d holds a shared-formula master" % r)
        body = "".join(fixrow(x) for r, x in book.rows if not (a1 <= r <= b1))
        last = book.rows[-1][0] - k
        head = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: '<dimension ref="A1:%s%d"/>' % (m.group(1), last), book.head)
        parts = dict(book.parts)
        parts[u.SHEET] = (head + body + book.tail).encode("utf-8")
        ch = book.parts[u.CHART].decode("utf-8")
        ch = re.sub(r"(<c:f>'S&amp;P 500'!\$[A-Z]\$)(\d+):(\$[A-Z]\$)(\d+)(</c:f>)",
                    lambda m: "%s%d:%s%d%s" % (m.group(1), mp(int(m.group(2))), m.group(3), mp(int(m.group(4))), m.group(5)), ch)
        parts[u.CHART] = ch.encode("utf-8")
        order = [x for x in book.order if x != "xl/calcChain.xml"]; parts.pop("xl/calcChain.xml", None)
        parts["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", parts["[Content_Types].xml"].decode()).encode()
        parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship Id="[^"]*" Type="[^"]*calcChain" Target="calcChain.xml"/>', "", parts["xl/_rels/workbook.xml.rels"].decode()).encode()
        w = parts["xl/workbook.xml"].decode()
        if "fullCalcOnLoad" not in w: w = re.sub(r"<calcPr ([^/]*)/>", r'<calcPr \1 fullCalcOnLoad="1"/>', w)
        parts["xl/workbook.xml"] = w.encode()
        with zipfile.ZipFile(a.out + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
            for n in order: zo.writestr(n, parts[n])
        os.replace(a.out + ".tmp", a.out)
        # proof
        b = u.Book(a.out)
        exp = [(mp(r), d, *v) for r, d, *v in book.data if not (a1 <= r <= b1)]
        if [tuple(x) for x in b.data] != exp: u.stop("PROOF FAILED: data rows")
        if len({x[1] for x in b.data}) != len(b.data): u.stop("PROOF FAILED: duplicates remain")
        if b.anchor != mp(book.anchor): u.stop("PROOF FAILED: anchor")
        for r, cc in book.row_cells.items():
            if a1 <= r <= b1: continue
            for col, c in cc.items():
                if c["formula"]:
                    nc = b.row_cells[mp(r)].get(col)
                    if nc is None or (nc["f"] or "") != remap(c["f"] or "") or nc["v"] != c["v"]: u.stop("PROOF FAILED: formula %s%d" % (col, r))
                elif col in "JMNOQ" and c["v"] != b.row_cells[mp(r)][col]["v"]: u.stop("PROOF FAILED: value %s%d" % (col, r))
        za, zb = zipfile.ZipFile(a.workbook), zipfile.ZipFile(a.out)
        ch_ = sorted(x for x in zb.namelist() if za.read(x) != zb.read(x))
        if ch_ != sorted([u.SHEET, u.CHART, "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/workbook.xml"]): u.stop("PROOF FAILED: parts %s" % ch_)
        print("Proof passed: %d rows deleted, %d rows remain, all formulas remapped, cached values equal, anchor row now %d." % (k, len(b.data), b.anchor))
        print("\n".join(u.audit(b)[:3]))
        return 0
    except u.Stop as e:
        print("STOP:", e); return 2
if __name__ == "__main__": sys.exit(main())