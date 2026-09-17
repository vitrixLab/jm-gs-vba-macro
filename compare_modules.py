import zipfile, os, re

xlsm1 = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.xlsm'
xlsm2 = r'C:\citrixlabph\globalsmile\Global-Smile_2026-v7.9.1.xlsm'
ed1 = r'C:\citrixlabph\globalsmile\graphify_v79'
ed2 = r'C:\citrixlabph\globalsmile\graphify_v791'

# Extract both
for ed, xlsmpath in [(ed1, xlsm1), (ed2, xlsm2)]:
    if os.path.exists(ed):
        import shutil
        shutil.rmtree(ed)
    os.makedirs(ed, exist_ok=True)
    with zipfile.ZipFile(xlsmpath, 'r') as z:
        z.extractall(ed)

# List bas modules in both
print("=== v7.9 BAS modules ===")
bas_dir1 = os.path.join(ed1, '_vba_extract/bas')
if os.path.exists(bas_dir1):
    for f in sorted(os.listdir(bas_dir1)):
        if f.endswith('.bas'):
            print(f'  {f}')

print()
print("=== v7.9.1 BAS modules ===")
bas_dir2 = os.path.join(ed2, '_vba_extract/bas')
if os.path.exists(bas_dir2):
    for f in sorted(os.listdir(bas_dir2)):
        if f.endswith('.bas'):
            print(f'  {f}')

# Compare module lists
print()
print("=== Comparison ===")
mods1 = set()
mods2 = set()
if os.path.exists(bas_dir1):
    for f in os.listdir(bas_dir1):
        if f.endswith('.bas'):
            mods1.add(f)
if os.path.exists(bas_dir2):
    for f in os.listdir(bas_dir2):
        if f.endswith('.bas'):
            mods2.add(f)

print(f"v7.9 has: {len(mods1)} BAS modules")
for m in sorted(mods1):
    print(f"  - {m}")

print(f"\nv7.9.1 has: {len(mods2)} BAS modules")
for m in sorted(mods2):
    print(f"  - {m}")

print(f"\nOnly in v7.9.1: {mods2 - mods1}")
print(f"Only in v7.9: {mods1 - mods2}")
PYEOF