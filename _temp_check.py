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
        kind, shown = 'fx-text', raw[:40]
    elif t_attr == 'b':
        kind, shown = 'bool', raw
    else:
        kind, shown = t_attr or '?', raw[:40]
    cols.setdefault(col, []).append((int(rown), kind, shown))

# Print column S data
if 'S' in cols:
    print('=== Column S in PJ2 ===')
    for r, k, s in cols['S'][:15]:
        print(f'  Row {r}: {k} = {s}')
    print(f'Total S cells: {len(cols["S"])}')
else:
    print('Column S not found in PJ2')
    # Show nearby columns
    for c in ['R', 'T', 'U']:
        if c in cols:
            print(f'Column {c}: {len(cols[c])} cells, first 3:')
            for r, k, s in cols[c][:3]:
                print(f'  Row {r}: {k} = {s}')