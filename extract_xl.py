import zipfile, os, re

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\graphify_xl"

# Extract the workbook
with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

# Check worksheet XML files
xl_worksheets = os.path.join(extract_dir, "xl/worksheets")
if os.path.exists(xl_worksheets):
    for f in os.listdir(xl_worksheets):
        if f.startswith("sheet") and f.endswith(".xml"):
            xml_path = os.path.join(xl_worksheets, f)
            with open(xml_path, encoding='utf-8', errors='replace') as fh:
                xml = fh.read()
            
            # Find column S cells
            CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
            s_cells = []
            for m in re.finditer(CELL_RE, xml):
                col = m.group(1)
                row = int(m.group(2))
                if col == 'S':
                    # Get the full cell
                    start = m.start()
                    # Find closing >
                    end = xml.find('>', start)
                    if end > start:
                        cell_text = xml[start:end+1]
                        s_cells.append((row, cell_text[:80]))
            
            print(f"=== {f} ===")
            print(f"Column S cells: {len(s_cells)}")
            for r, s in s_cells[:10]:
                print(f"  Row {r}: {s}")
            print()
else:
    print("No worksheets folder found")