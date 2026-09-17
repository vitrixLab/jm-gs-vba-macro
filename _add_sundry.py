import re, os
BASE = r'C:\citrixlabph\globalsmile\graphify'
xml_path = os.path.join(BASE, 'xl', 'worksheets/sheet2.xml')
with open(xml_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Find all column S cells
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)
cells = []
for m in CELL_RE.finditer(xml):
    col, rown, attrs, _, inner = m.groups()
    if col == 'S':
        cells.append((int(rown), m.group(0)))

print(f'Found {len(cells)} column S cells')
# Show first 10
for r, full_cell in cells[:15]:
    print(f'  Row {r}: {full_cell[:80]}')

# Also check what's in rows 13-162 (CDJ_DATA_START range)
print()
print('Column S in CDJ_DATA_START range (rows 15-162):')
for r, full_cell in cells:
    if 15 <= r <= 162:
        print(f'  Row {r}: {full_cell[:80]}')