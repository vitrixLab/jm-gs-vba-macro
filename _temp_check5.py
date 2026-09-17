import re, os
BASE = r'C:\citrixlabph\globalsmile\graphify'
xml = open(os.path.join(BASE, 'xl', 'worksheets/sheet10.xml'), encoding='utf-8', errors='replace').read()
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)
cols = {}
for m in CELL_RE.finditer(xml):
    col, rown, attrs, _, inner = m.groups()
    t = re.search(r' t="(\w+)"', attrs)
    t_attr = t.group(1) if t else None
    vm = re.search(r'<v>(.*?)</v>', inner or '', re.S)
    raw = vm.group(1) if vm else ''
    if raw == '' and t_attr != 'str':
        continue
    if t_attr == 's':
        kind, shown = 'text', raw
    elif t_attr in (None, 'n'):
        x = float(raw)
        shown = f'{raw[:20]}'
    elif t_attr == 'str':
        # Extract the <t> element content
        shown = re.search(r'<t[^>]*>(.*?)</t>', inner or '', re.S)
        shown = shown.group(1).strip() if shown else raw[:40]
    elif t_attr == 'b':
        kind, shown = 'bool', raw
    else:
        kind, shown = t_attr or '?', raw[:40]
    cols.setdefault(col, []).append((int(rown), kind, shown))

# Check column G in PJ2 (PJ_COL_COA = 7 = G) - show first 15 text entries
print('=== Column G (COA descriptions) in PJ2 ===')
if 'G' in cols:
    items = [(r, k, s) for r, k, s in cols['G']]
    text_items = [(r, s) for r, k, s in items if k == 'text']
    print(f'Total text cells: {len(text_items)}')
    for r, s in text_items[:15]:
        print(f'  Row {r}: {s}')