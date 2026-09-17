import zipfile, os, re

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\graphify_xl_pj"

# Extract the workbook (fresh extract)
if os.path.exists(extract_dir):
    import shutil
    shutil.rmtree(extract_dir)
os.makedirs(extract_dir, exist_ok=True)

with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

# Modify PJ2 sheet (sheet10.xml) - fix rows 8 and 10 with sundry descriptions
pj_sheet_path = os.path.join(extract_dir, "xl/worksheets/sheet10.xml")
with open(pj_sheet_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Find G column cells and their current values
CELL_RE = re.compile(r'<c r="G(\d+)"[^>]*>(.*?)</c>', re.S)
g_cells = {}
for m in CELL_RE.finditer(xml):
    row_str = m.group(1)
    row = int(row_str)
    content = m.group(2)
    g_cells[row] = content

print(f"G column cells found: {len(g_cells)}")

# The descriptions to set
descriptions = {
    8: "Clinician Fees",   # Row 8 should be "Clinician Fees"
    10: "Consulting",      # Row 10 should be "Consulting"
}

# Modify the specific rows
modifications = 0
for row, desc in descriptions.items():
    if row in g_cells:
        current_content = g_cells[row]
        # Check if it's already a description or a number
        # Look for <v> tag
        v_match = re.search(r'<v>(.*?)</v>', current_content, re.S)
        if v_match:
            current_text = v_match.group(1).strip()
        else:
            # Self-closing cell, get text directly
            current_text = current_content.strip()
        
        # Check if it's a number (original COA code) or already text
        # Numbers typically don't start with text names
        if current_text.startswith(('M', 'C', 'P', 'C')) or len(current_text) > 10:
            # Already has text description, skip
            print(f'Row {row}: Already has description "{current_text}" - skipping')
            continue
        
        # It's a number/COA code, replace with description
        if v_match:
            # Replace the <v> content
            new_content = current_content[:v_match.start()] + f'<v>{desc}</v>' + current_content[v_match.end():]
        else:
            # Convert self-closing to have <v>Content</v>
            # Format: <c r="G8" s="626"/> -> <c r="G8" s="626" t="str"><v>Clinician Fees</v></c>
            # Need to find the cell pattern and replace
            # First, let's just try replacing the entire cell
            new_cell = f'<c r="G{row}" s="626" t="str"><v>{desc}</v></c>'
            # Find the exact position and replace
            # Get the exact start position from the original search
            target_start = xml.find(f'G{row}"', xml.find(f'G{row}"'))
            # This is getting complex, let me use a different approach
            # Just replace the entire cell by finding its start and end
            
        # Simpler approach: replace the entire cell XML
        # Find the cell and replace it
        search_pattern = rf'<c r="G{row}"[^>]*>.*?</c>'
        replacement = f'<c r="G{row}" s="626" t="str"><v>{desc}</v></c>'
        match = re.search(search_pattern, xml)
        if match:
            xml = xml[:match.start()] + replacement + xml[match.end():]
            modifications += 1
            print(f'Row {row}: Changed from "{current_text}" to "{desc}"')
        else:
            print(f'Row {row}: Could not find cell to replace')
    else:
        print(f'Row {row}: Not found in G cells')

# Write the modified XML back
with open(pj_sheet_path, 'w', encoding='utf-8') as f:
    f.write(xml)

print(f'\nTotal modifications: {modifications}')
print('PJ2 sheet updated successfully')
print('Sundry account descriptions in PJ2 column G:')
for row in [6, 8, 9, 10]:
    if row in g_cells:
        content = g_cells[row]
        v_match = re.search(r'<v>(.*?)</v>', content, re.S)
        text = v_match.group(1).strip() if v_match else content.strip()
        print(f'  Row {row}: {text}')