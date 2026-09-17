import re, os
BASE = r'C:\citrixlabph\globalsmile\graphify'

def read(p):
    with open(p, encoding='utf-8', errors='replace') as f:
        return f.read()

# Read CDJ sheet (sheet2)
xml = read(os.path.join(BASE, 'xl', 'worksheets/sheet2.xml'))
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
        kind, shown = 'fx-text', raw[:40]
    elif t_attr == 'b':
        kind, shown = 'bool', raw
    else:
        kind, shown = t_attr or '?', raw[:40]
    cols.setdefault(col, []).append((int(rown), kind, shown))

# Check column S in CDJ
if 'S' in cols:
    print('=== Column S in CDJ ===')
    for r, k, s in cols['S'][:15]:
        print(f'  Row {r}: {k} = {s}')
    print(f'Total S cells: {len(cols["S"])}')
else:
    print('Column S not found in CDJ')

# Also check expense-related columns (N-T per the code)
print()
print('=== Expense columns (N-T) in CDJ ===')
for c in ['N', 'O', 'P', 'Q', 'R', 'S', 'T']:
    if c in cols:
        text_cells = [(r, s) for r, k, s in cols[c] if k == 'text']
        print(f'Column {c}: {len(text_cells)} text cells')
        for r, s in text_cells[:3]:
            print(f'  Row {r}: {s[:60]}')