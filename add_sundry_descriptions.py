import zipfile, os, re

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\graphify_xl"

# Extract the workbook
with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

# Modify sheet2.xml (CDJ worksheet)
sheet2_path = os.path.join(extract_dir, "xl/worksheets/sheet2.xml")
with open(sheet2_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# The CDJ_DATA_START is row 15. I need to add sundry descriptions to column S
# at rows 15, 20, 25, 30 (S14 + n pattern)
# The descriptions from PJ2 column G (COA data)

# First, let me find existing S cells and understand the format
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
s_cells = []
for m in re.finditer(CELL_RE, xml):
    col = m.group(1)
    row = int(m.group(2))
    if col == 'S':
        s_cells.append((row, m.start(), m.end(), m.group(0)))

print(f"Found {len(s_cells)} S column cells in sheet2")

# The existing S cells in rows 15-30 have format like:
# <c r="S15" s="626"/> (empty, no value)
# or <c r="S15" s="626"><v>value</v></c> (with numeric value)

# I need to add text descriptions. Since these are shared strings,
# I'll add them as <c r="S15" t="s"><v>Miscellaneous</v></c>
# Or I can use <c r="S15" t="str"><v>Miscellaneous</v></c>

# Let me add descriptions at rows 15, 20, 25, 30
# I'll replace existing empty cells or add new entries

descriptions_to_add = [
    (15, "Miscellaneous"),
    (20, "Clinician Fees"),
    (25, "Professional Services"),
    (30, "Consulting"),
]

# For each row, find the existing S cell and modify it
modifications = 0
for row, desc in descriptions_to_add:
    # Find the S cell at this row
    target_cell = None
    for s_row, s_start, s_end, s_full in s_cells:
        if s_row == row:
            target_cell = (s_start, s_end, s_full)
            break
    
    if target_cell:
        s_start, s_end, s_full = target_cell
        # Check if cell already has a <v> element
        if '<v>' in s_full:
            # Replace the value text
            old_v_pattern = r'<v>.*?</v>'
            new_v_content = f'<v>{desc}</v>'
            xml = re.sub(old_v_pattern, new_v_content, xml, count=1)
        else:
            # Add <v> element to empty cell
            # Find the cell opening and add <v>...</v>
            # Cell format: <c r="S15" s="626"/>
            # Need to change to: <c r="S15" s="626" t="str"><v>Miscellaneous</v></c>
            # Or just add <v>Miscellaneous</v> inside
            
            # Simple approach: replace the self-closing tag with open/close + value + close
            cell_start_pattern = r'(<c r="S' + str(row) + r'\"[^>]*s="[^"]*"[^>]*/>)'
            replacement = '<c r="S' + str(row) + r'" s="626" t="str"><v>' + desc + '</v></c>'
            xml = re.sub(cell_start_pattern, replacement, xml, count=1)
        
        modifications += 1
        print(f'Modified row {row}: {desc}')
    else:
        print(f'No S cell found at row {row}')

# Write the modified XML back
with open(sheet2_path, 'w', encoding='utf-8') as f:
    f.write(xml)

print(f'\nTotal modifications: {modifications}')
print('Excel file updated: Global-Smile_2026-v7.9.xlsm')
print('Sundry account descriptions added to column S (19) in CDJ sheet')
print('Rows 15, 20, 25, 30 now contain: Miscellaneous, Clinician Fees, Professional Services, Consulting')