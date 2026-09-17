import re, os

BASE = r'C:\citrixlabph\globalsmile\graphify'
xml_path = os.path.join(BASE, 'xl', 'worksheets/sheet2.xml')
with open(xml_path, encoding='utf-8', errors='replace') as f:
    xml = f.read()

# Do the actual replacements
for row, desc in [(15, 'Miscellaneous'), (20, 'Clinician Fees'), (25, 'Professional Services'), (30, 'Consulting')]:
    old_pattern = r'<c r="S' + str(row) + r'\"[^>]*?>.*?</c>'
    new_cell = '<c r="S' + str(row) + r'" s="575" t="s"><v>' + desc + '</v></c>'
    xml = re.sub(old_pattern, new_cell, xml, count=1)
    print(f'Replaced row {row} with: {desc}')

# Write back the modified XML
with open(xml_path, 'w', encoding='utf-8') as f:
    f.write(xml)

print()
print('Excel file modified successfully!')
print('Added sundry account descriptions to column S (19) in CDJ sheet')