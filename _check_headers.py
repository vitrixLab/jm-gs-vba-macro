import re
with open(r'C:\citrixlabph\globalsmile\_vba_extract\v78_unzip\xl\worksheets\sheet2.xml', encoding='utf-8', errors='replace') as f:
    xml = f.read()
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
cells = list(re.finditer(CELL_RE, xml))
# Show S cells in rows 13-15
for m in cells:
    col = m.group(1)
    row = int(m.group(2))
    if col == 'S' and 13 <= row <= 15:
        cell_xml = m.group(0)
        v_match = re.search(r'<v>(.*?)</v>', cell_xml, re.S)
        text = v_match.group(1).strip() if v_match else "(empty/no <v>)"
        print(f"Row {row}: {text!r}  full={cell_xml[:80]!r}")
EOF