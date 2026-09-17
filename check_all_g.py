import re, os

xml_path = r'C:\citrixlabph\globalsmile\graphify_xl_pj\xl\worksheets\sheet10.xml'
with open(xml_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Check ALL G column cells and their values
print("=== All G column cells in PJ2 ===")
CELL_RE = re.compile(r'<c r="G(\d+)"[^>]*>(.*?)</c>', re.S)
for m in CELL_RE.finditer(xml):
    row_str = m.group(1)
    row = int(row_str)
    content = m.group(2).strip()
    # Clean up HTML entities or extra spaces
    text = content.replace('&', '&').replace('<', '<').replace('>', '>')
    # Remove any v tags if present to get just the text
    inner_match = re.search(r'<v>(.*?)</v>', content, re.S)
    if inner_match:
        text = inner_match.group(1).strip()
    print(f'Row {row}: {text[:50]}')

# Specifically check rows 6-12 where COA data exists
print()
print("=== G column rows 6-12 ===")
for row in range(6, 13):
    pattern = f'<c r="G{row}"[^>]*>.*?</c>'
    matches = re.findall(pattern, xml)
    for m in matches:
        v_match = re.search(r'<v>(.*?)</v>', m, re.S)
        text = v_match.group(1).strip() if v_match else m
        print(f'  Row {row}: {text[:50]}')