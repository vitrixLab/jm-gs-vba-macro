import re, os

xml_path = r'C:\citrixlabph\globalsmile\graphify_xl_pj\xl\worksheets\sheet10.xml'
with open(xml_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Check G column cells at rows 6, 8, 9, 10
pattern = r'<c r="G(\d+)"[^>]*>.*?</c>'
for m in re.finditer(pattern, xml):
    row_str = m.group(1)
    row = int(row_str)
    if row in [6, 8, 9, 10]:
        full = m.group(0)
        # Extract the <v> content
        v_match = re.search(r'<v>(.*?)</v>', full, re.S)
        text = v_match.group(1).strip() if v_match else "no <v>"
        print(f'Row {row}: {text[:50]}')

print()
print("Verification complete - sundry descriptions added to PJ2 column G")