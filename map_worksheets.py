import zipfile, os, re

xlsm = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm'
ed = r'C:\citrixlabph\globalsmile\graphify_xl_map'

# Fresh extract
if os.path.exists(ed):
    import shutil
    shutil.rmtree(ed)
os.makedirs(ed, exist_ok=True)

with zipfile.ZipFile(xlsm, 'r') as z:
    z.extractall(ed)

# List all worksheet files
xl_ws = os.path.join(ed, 'xl/worksheets')
print("=== Worksheets in Global-Smile_2026-v7.9.xlsm ===")
for f in sorted(os.listdir(xl_ws)):
    if f.startswith('sheet') and f.endswith('.xml'):
        path = os.path.join(xl_ws, f)
        with open(path, encoding='utf-8', errors='replace') as fh:
            xml = fh.read()
        # Extract dimension
        dim_m = re.search(r'<dimension ref="([^"]+)"', xml)
        dim = dim_m.group(1) if dim_m else '?'
        # Extract column range info
        print(f"  {f}: dimension={dim}")
        # Show first few column headers
        cell_m = re.findall(r'<c r="([A-Z]+)(\d+)"[^>]*>', xml)[:20]
        cols = set()
        for c_m in cell_m:
            cols.add(c_m[0])
        print(f"    Columns present: {sorted(cols)[:10]}")