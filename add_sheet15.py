import zipfile, os, re, shutil

xlsm_from = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.1.xlsm'
xlsm_to = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm'
extract_from = r'C:\citrixlabph\globalsmile\extract_v791'
extract_to = r'C:\citrixlabph\globalsmile\extract_v79'

# Fresh extracts
for ed in [extract_from, extract_to]:
    if os.path.exists(ed):
        shutil.rmtree(ed)
    os.makedirs(ed, exist_ok=True)
    with zipfile.ZipFile(xlsm_from if ed == extract_from else xlsm_to, 'r') as z:
        z.extractall(ed)

# Copy sheet15 from v7.9.1 to v7.9
src_sheet = os.path.join(extract_from, 'xl/worksheets/sheet15.xml')
dst_sheet = os.path.join(extract_to, 'xl/worksheets/sheet15.xml')
shutil.copy2(src_sheet, dst_sheet)

# Also copy the workbook.xml to preserve sheet references
src_wb = os.path.join(extract_from, 'xl/workbook.xml')
dst_wb = os.path.join(extract_to, 'xl/workbook.xml')
shutil.copy2(src_wb, dst_wb)

# Verify the sheet was added
print("=== Checking sheet15 in v7.9 ===")
with open(dst_sheet, encoding='utf-8', errors='replace') as f:
    xml = f.read()
print("Sheet15 found: " + str('sheet15' in xml))
dim_m = re.search(r'<dimension ref="([^"]+)"', xml)
if dim_m:
    print("Sheet15 dimension: " + dim_m.group(1))

# Also check workbook.xml for sheet references
with open(dst_wb, encoding='utf-8', errors='replace') as f:
    wb_xml = f.read()
# Check if sheet15 is referenced
if 'sheet15' in wb_xml.lower():
    print("sheet15 referenced in workbook.xml")
else:
    print("sheet15 NOT referenced in workbook.xml (may need to add reference)")

# List all worksheets in the new file
print()
print("=== All worksheets in new v7.9 file ===")
xl_to = os.path.join(extract_to, 'xl/worksheets')
for f in sorted(os.listdir(xl_to)):
    if f.startswith('sheet') and f.endswith('.xml'):
        path = os.path.join(xl_to, f)
        with open(path, encoding='utf-8', errors='replace') as fh:
            xm = fh.read()
        dim_m = re.search(r'<dimension ref="([^"]+)"', xm)
        print("  " + f + ": dimension=" + (dim_m.group(1) if dim_m else "?"))
PYEOF