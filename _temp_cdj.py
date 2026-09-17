import re
with open(r'C:\citrixlabph\globalsmile\_vba_extract\v78_unzip\xl\worksheets\sheet2.xml', encoding='utf-8', errors='replace') as f:
    xml = f.read()
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
cells = list(re.finditer(CELL_RE, xml))
s_cells = [(m.group(2), m.group(1)) for m in cells if m.group(1) == 'S']
# Show rows 10-20 S cells
for row, col in s_cells:
    row_num = int(row)
    if 10 <= row_num <= 25:
        print(f'Row {row_num}: col={col}')

# Also check row 17 specifically
for row, col in s_cells:
    row_num = int(row)
    if row_num == 17:
        # Get full cell content
        print(f'\nRow 17 S cell full: ')
        # Find the actual cell in original xml
        m = next(c for c in cells if c.group(2) == '17' and c.group(1) == 'S')
        print(repr(m.group(0)))
        # Get the content after >
        start = m.start() + m.group(0).rfind('>') + 1
        cell_xml = m.group(0)
        v_match = re.search(r'<v>(.*?)</v>', cell_xml, re.S)
        if v_match:
            text = v_match.group(1).strip()
            print(f'  Value: {text}')
        else:
            print('  No <v> element, likely empty')