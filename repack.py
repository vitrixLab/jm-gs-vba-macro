import zipfile, os, shutil

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

# Re-pack the Excel file
output_path = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9-updated.xlsm'
with zipfile.ZipFile(xlsm_to, 'r') as zin:
    with zipfile.ZipFile(output_path, 'w') as zout:
        for item in zin.namelist():
            if 'sheet15' in item:
                # Read from the extracted v7.9 folder (which now has sheet15)
                source_path = os.path.join(extract_to, item)
                if os.path.exists(source_path):
                    with open(source_path, 'rb') as f:
                        zout.writestr(item, f.read())
                else:
                    zout.writestr(item, b'')
            elif item.startswith('xl/'):
                zout.writestr(item, zin.read(item))
            else:
                zout.writestr(item, zin.read(item))

print("New file created: " + output_path)
print("Sheet15 (Unmapped Detail) added to v7.9 file")

# Verify the new file
import zipfile as zf
with zipfile.ZipFile(output_path, 'r') as zcheck:
    names = zcheck.namelist()
    print("sheet15 in new file: " + str('sheet15.xml' in names))
    for n in names:
        if 'sheet15' in n:
            print("  " + n)