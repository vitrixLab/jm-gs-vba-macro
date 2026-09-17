import zipfile, os

xlsm_path = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.2.xlsm"
extract_dir = r"C:\citrixlabph\globalsmile\extract_vba_v792"

# Extract the VBA project
with zipfile.ZipFile(xlsm_path, 'r') as z:
    z.extractall(extract_dir)

print("Extracted VBA files:")
for root, dirs, files in os.walk(extract_dir):
    for f in files:
        print(os.path.join(root, f))