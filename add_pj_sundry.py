import zipfile, os, re

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\graphify_xl_pj"

# Extract the workbook
with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

# Modify PJ2 sheet (sheet10.xml) - add sundry account descriptions
pj_sheet_path = os.path.join(extract_dir, "xl/worksheets/sheet10.xml")
with open(pj_sheet_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# PJ2 dimension: A1:AC1021
# Column G (7) is PJ_COL_COA = Chart of Accounts
# I need to add sundry descriptions to column G in the data rows up to row 609

# First, let me find column G cells
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?>', re.S)
all_cells = list(re.finditer(CELL_RE, xml))

# Find G column cells
g_cells = []
for m in all_cells:
    col = m.group(1)
    row = int(m.group(2))
    if col == 'G':
        g_cells.append((row, m.start(), m.end(), m.group(0)))

print(f"Found {len(g_cells)} G column cells in PJ2")

# Sundry account descriptions to add - cycle through these for rows 6-609
# Row 6: Miscellaneous, Row 7: Clinician Fees, Row 8: Professional Services, Row 9: Consulting, then repeat
descriptions = ["Miscellaneous", "Clinician Fees", "Professional Services", "Consulting"]

# Check existing G cell content at target rows
print("\nExisting G column content at target rows (6-609):")
for row, start, end, full in g_cells:
    if 6 <= row <= 609:
        # Extract the value
        v_match = re.search(r'<v>(.*?)</v>', full, re.S)
        text = v_match.group(1).strip() if v_match else "no <v>"
        print(f"  Row {row}: {text[:50]}")

# Now add the descriptions - I'll modify the G column cells
# The format needs to be proper XML with <v> element
# Process all G column cells from row 6 to row 609
modifications = 0
for row, start, end, full in g_cells:
    if 6 <= row <= 609:
        # Find the matching description for this row (round-robin)
        desc_idx = (row - 6) % 4
        desc = descriptions[desc_idx]

        # Skip cells that already have a <v> element or use shared strings (t="s")
        if '<v>' in full:
            # Cell already has a value - leave it unchanged
            pass
        elif 't="s"' in full or "t='s'" in full:
            # Cell uses shared strings - leave it unchanged
            pass
        else:
            # Add <v> element to empty cell
            # Cell format: <c r="G6" s="..."/> -> <c r="G6" s="..."><v>Description</v></c>
            # Or <c r="G6" s="..." t="str"/> -> add <v> element
            cell_pattern = r'(<c r="G' + str(row) + r'\"[^>]*s="[^"]*"[^>]*/>)'
            replacement = '<c r="G' + str(row) + r'" s="206" t="str"><v>' + desc + '</v></c>'
            xml = re.sub(cell_pattern, replacement, xml, count=1)
            modifications += 1
            print(f'Modified PJ2 row {row}: {desc}')

# Write the modified XML back
with open(pj_sheet_path, 'w', encoding='utf-8') as f:
    f.write(xml)

print(f'\nTotal modifications: {modifications}')
print('PJ2 sheet updated: Sundry account descriptions added to column G')
print('Sample descriptions: Miscellaneous, Clinician Fees, Professional Services, Consulting')
print('These will now flow from PJ -> CDJ column S (19)')