"""Parse ISM Manufacturing industry lists for all seven indices from report text.
 
Input: a text file with one heading line per index ("New Orders -", "Production -", ...)
followed by that index's paragraph(s).
 
Usage:
    python ism_industry_parser.py report.txt 2026-10
Writes ism_industries_<month>.csv (long format) and prints validation checks plus
a quick industry-by-index matrix (+ up, - down, 0 neutral) for eyeballing.
"""
import csv
import re
import sys
 
INDUSTRIES = [
    "Apparel, Leather & Allied Products", "Chemical Products",
    "Computer & Electronic Products",
    "Electrical Equipment, Appliances & Components", "Fabricated Metal Products",
    "Food, Beverage & Tobacco Products", "Furniture & Related Products",
    "Machinery", "Miscellaneous Manufacturing", "Nonmetallic Mineral Products",
    "Paper Products", "Petroleum & Coal Products", "Plastics & Rubber Products",
    "Primary Metals", "Printing & Related Support Activities", "Textile Mills",
    "Transportation Equipment", "Wood Products",
]
CANON = {i.lower(): i for i in INDUSTRIES}
 
# The 18 ISM Services (Non-Manufacturing) industries.
SERVICES_INDUSTRIES = [
    "Accommodation & Food Services", "Agriculture, Forestry, Fishing & Hunting",
    "Arts, Entertainment & Recreation", "Construction", "Educational Services",
    "Finance & Insurance", "Health Care & Social Assistance", "Information",
    "Management of Companies & Support Services", "Mining", "Other Services",
    "Professional, Scientific & Technical Services", "Public Administration",
    "Real Estate, Rental & Leasing", "Retail Trade", "Transportation & Warehousing",
    "Utilities", "Wholesale Trade",
]
 
INDEX_ORDER = ["Manufacturing PMI", "New Orders", "Production", "Employment",
               "Supplier Deliveries", "Inventories", "Customer Inventories"]
ABBREV = {"Manufacturing PMI": "PMI", "New Orders": "NO", "Production": "PR",
          "Employment": "EM", "Supplier Deliveries": "SD", "Inventories": "INV",
          "Customer Inventories": "CI"}
HEADINGS = {
    "ism manufacturing": "Manufacturing PMI", "manufacturing": "Manufacturing PMI",
    "manufacturing pmi": "Manufacturing PMI", "new orders": "New Orders",
    "production": "Production", "employment": "Employment",
    "supplier deliveries": "Supplier Deliveries", "deliveries": "Supplier Deliveries",
    "inventories": "Inventories",
    "customer inventories": "Customer Inventories",
    # Services (Non-Manufacturing)
    "ism non-manufacturing": "Services PMI", "non-manufacturing": "Services PMI",
    "ism services": "Services PMI", "services pmi": "Services PMI",
    "business activity": "Business Activity",
}
 
# Native ISM label, direction, regex. Order matters (first match wins).
CATS = [
    ("no change", "neutral", r"\bno change\b"),
    ("too high", "up", r"\btoo high\b"),
    ("too low", "down", r"\btoo low\b"),
    ("growth", "up", r"\bgrowth\b"),
    ("contraction", "down", r"\b(?:contraction|decline)\b"),
    ("increase", "up", r"\b(?:increase|increasing)\b"),
    ("decrease", "down", r"\b(?:decrease|decreasing)\b"),
    ("slower", "up", r"\bslower\b"),
    ("faster", "down", r"\bfaster\b"),
    ("higher", "up", r"\bhigher\b"),
    ("lower", "down", r"\blower\b"),
]
 
NUMS = {w: n for n, w in enumerate(
    "zero one two three four five six seven eight nine ten eleven twelve thirteen "
    "fourteen fifteen sixteen seventeen eighteen".split())}
NUMS.update({"no": 0, "only": 1})
NUMTOKEN = re.compile(r"\b(\d+|" + "|".join(NUMS) + r")\b", re.I)
LEAD = re.compile(r"^Of\s+(?:the\s+)?\d+\s+(?:manufacturing\s+|services\s+)?industries,\s*", re.I)
ONLY_IS = re.compile(r"\bonly industry\b.*?\bis\s+(.+?)\.?$", re.I)
MONTHS = ["january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"]
 
 
def split_items(raw):
    raw = " ".join(raw.split())
    if ";" in raw:  # names contain commas, so split on semicolons only
        parts = re.split(r";\s*(?:and\s+)?", raw)
    else:
        parts = re.split(r"\s+and\s+", raw)
    return [re.sub(r"^and\s+", "", p.strip(), flags=re.I) for p in parts if p.strip()]
 
 
def parse_sentence(s):
    s = " ".join(s.split())
    cat = next(((l, d) for l, d, rx in CATS if re.search(rx, s, re.I)), None)
    if cat is None:
        return None
    if ":" in s:
        prefix, raw = s.rsplit(":", 1)
    elif "(" in s and ")" in s:
        prefix, raw = s[: s.index("(")], s[s.index("(") + 1: s.rindex(")")]
    else:
        prefix, raw = s, ""
        mo = ONLY_IS.search(s)  # "the only industry reporting a decrease ... is Construction."
        if mo:
            prefix, raw = "only", mo.group(1)
    m = NUMTOKEN.search(LEAD.sub("", prefix))
    stated = None
    if m:
        tok = m.group(1).lower()
        stated = int(tok) if tok.isdigit() else NUMS[tok]
    items = split_items(raw.strip().rstrip(".").strip()) if raw.strip() else []
    return cat, stated, items
 
 
def parse_section(index, text, month_num, industries=None):
    industries = industries or INDUSTRIES
    canon_map = {i.lower(): i for i in industries}
    rows, problems, seen_cats, stated_neutral = [], [], set(), None
    mm = re.search(r"\b(" + "|".join(MONTHS) + r")\b", text, re.I)
    if mm and MONTHS.index(mm.group(1).lower()) + 1 != month_num:
        problems.append(f"text mentions {mm.group(1)} but run month is {month_num}")
    recognized = 0
    for s in re.split(r"(?<=\.)\s+(?=[A-Z])", " ".join(text.split())):
        parsed = parse_sentence(s)
        if parsed is None:
            continue
        recognized += 1
        (label, direction), stated, items = parsed
        if label == "no change":
            stated_neutral = stated
            continue
        if label in seen_cats:
            problems.append(f"'{label}' sentence appears twice")
        seen_cats.add(label)
        if stated is not None and stated != len(items):
            problems.append(f"{label}: sentence says {stated}, parsed {len(items)}")
        for rank, name in enumerate(items, 1):
            canon = canon_map.get(name.lower())
            if canon is None:
                problems.append(f"{label}: unrecognized industry '{name}'")
            rows.append({"index": index, "category": label, "direction": direction,
                         "rank": rank, "industry": canon or name})
    if not recognized:
        problems.append("no industry sentences recognized; wording may have changed")
    names = [r["industry"] for r in rows]
    for dup in sorted({n for n in names if names.count(n) > 1}):
        problems.append(f"industry in more than one list: {dup}")
    neutral = [i for i in industries if i not in set(names)]
    if stated_neutral is not None and stated_neutral != len(neutral):
        problems.append(f"report says {stated_neutral} unchanged, inferred {len(neutral)}")
    for name in neutral:
        rows.append({"index": index, "category": "neutral (inferred)",
                     "direction": "neutral", "rank": "", "industry": name})
    return rows, problems
 
 
def split_sections(text):
    sections, current = {}, None
    for line in text.splitlines():
        key = line.strip().rstrip("-: ").strip().lower()
        if key in HEADINGS:
            current = HEADINGS[key]
            sections.setdefault(current, [])
        elif current and line.strip():
            sections[current].append(line)
    return {k: "\n".join(v) for k, v in sections.items()}
 
 
if __name__ == "__main__":
    path, month = sys.argv[1], sys.argv[2]
    month_num = int(month.split("-")[1])
    sections = split_sections(open(path, encoding="utf-8").read())
    all_rows, any_problem = [], False
    for index in INDEX_ORDER:
        if index not in sections:
            print(f"[MISSING] {index}: heading not found")
            any_problem = True
            continue
        rows, problems = parse_section(index, sections[index], month_num)
        counts = {}
        for r in rows:
            counts[r["category"]] = counts.get(r["category"], 0) + 1
        status = "OK " if not problems else "CHECK"
        print(f"[{status}] {index}: {counts}")
        for p in problems:
            print("        -", p)
            any_problem = True
        all_rows += rows
    out = f"ism_industries_{month}.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["month", "index", "category", "direction", "rank", "industry"])
        w.writeheader()
        for r in all_rows:
            w.writerow({"month": month, **r})
    print("\nWrote", out, f"({len(all_rows)} rows)")
    sign = {"up": "+", "down": "-", "neutral": "0"}
    lookup = {(r["index"], r["industry"]): sign[r["direction"]] for r in all_rows}
    print("\n" + "Industry".ljust(46) + " ".join(ABBREV[i].rjust(3) for i in INDEX_ORDER))
    for ind in INDUSTRIES:
        print(ind.ljust(46) + " ".join(lookup.get((i, ind), "?").rjust(3) for i in INDEX_ORDER))
    sys.exit(1 if any_problem else 0)