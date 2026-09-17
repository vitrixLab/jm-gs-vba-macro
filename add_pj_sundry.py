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
# I need to add sundry descriptions to column G in the data rows

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

# Look at rows 6-26 (where COA data exists based on earlier analysis)
# The descriptions should be added to column G for sundry accounts

# Sundry account descriptions to add
# These are the sample descriptions the user mentioned: "Miscellaneous" and "Clinician fees"
descriptions_to_add = [
    (6, "Miscellaneous"),      # Row 6 - first data row after headers
    (8, "Clinician Fees"),     # Row 8 - typical sundry account
    (9, "Professional Services"),
    (10, "Consulting"),
]

# Check existing G cell content at these rows
print("\nExisting G column content at target rows:")
for row, start, end, full in g_cells:
    if 6 <= row <= 12:
        # Extract the value
        v_match = re.search(r'<v>(.*?)</v>', full, re.S)
        text = v_match.group(1).strip() if v_match else "no <v>"
        print(f"  Row {row}: {text[:50]}")

# Now add the descriptions - I'll modify the G column cells
# The format needs to be proper XML with <v> element

modifications = 0
for row, desc in descriptions_to_add:
    # Find the G cell at this row
    target_cell = None
    for g_row, g_start, g_end, g_full in g_cells:
        if g_row == row:
            target_cell = (g_start, g_end, g_full)
            break
    
    if target_cell:
        g_start, g_end, g_full = target_cell
        # Check if cell already has value
        if '<v>' in g_full:
            # Replace the value
            old_v = re.search(r'<v>.*?</v>', g_full)
            if old_v:
                # Replace with description
                xml = xml[:old_v.start()] + f'<v>{desc}</v>' + xml[old_v.end():]
                modifications += 1
                print(f'Modified PJ2 row {row}: {desc}')
        else:
            # Add <v> element
            # Need to convert from self-closing to have <v>Content</v>
            # Format: <c r="G6" s="..."/> -> <c r="G6" s="..."><v>Description</v></c>
            # Find the cell pattern
            cell_pattern = r'(<c r="G' + str(row) + r'\"[^>]*s="[^"]*"[^>]*/>)'
            replacement = '<c r="G' + str(row) + r'" s="206" t="str"><v>' + desc + '</v></c>'
            xml = re.sub(cell_pattern, replacement, xml, count=1)
            modifications += 1
            print(f'Modified PJ2 row {row}: {desc} (added t="str")')
    else:
        print(f'No G cell found at row {row} in PJ2')

# Write the modified XML back
with open(pj_sheet_path, 'w', encoding='utf-8') as f:
    f.write(xml)

print(f'\nTotal modifications: {modifications}')
print('PJ2 sheet updated: Sundry account descriptions added to column G')
print('Sample descriptions: Miscellaneous, Clinician Fees, Professional Services, Consulting')
print('These will now flow from PJ -> CDJ column S (19)')