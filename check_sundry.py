import zipfile, os, re

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\graphify_xl"

with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

# Check sheet2 (CDJ worksheet)
xl_worksheets = os.path.join(extract_dir, "xl/worksheets/sheet2.xml")
with open(xl_worksheets, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Find S cells with text content (t="s" or str type)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
cells = list(re.finditer(CELL_RE, xml))

# Find S column cells with actual text/v values
s_cells_with_text = []
for m in cells:
    col = m.group(1)
    row = int(m.group(2))
    if col == 'S':
        # Get the full cell content
        start = m.start() + m.group(0).rfind('>') + 1
        # Find closing tags
        cell_xml = m.group(0)
        # Check if there's a <v> element with text
        v_match = re.search(r'<v>(.*?)</v>', cell_xml, re.S)
        if v_match:
            text = v_match.group(1).strip()
            if text and text not in ('0', ''):
                s_cells_with_text.append((row, text[:50]))

print(f"S column cells with text in sheet2: {len(s_cells_with_text)}")
print()
print("First 20 S cells with text:")
for r, text in s_cells_with_text[:20]:
    print(f"  Row {r}: {text}")

# Also check rows 15-162 (CDJ_DATA_START range)
print(f"\nS column cells with text in rows 15-162:")
s_in_range = [(r, t) for r, t in s_cells_with_text if 15 <= r <= 162]
print(f"Count: {len(s_in_range)}")
for r, text in s_in_range[:20]:
    print(f"  Row {r}: {text}")