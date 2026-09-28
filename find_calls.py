import zipfile
import sys

XLsm_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.1.xlsm"
try:
    with zipfile.ZipFile(XLsm_PATH, 'r') as z:
        vba_data = z.read("xl/vbaProject.bin")
    vba_text = vba_data.decode("utf-8", errors="replace")
    
    lines = vba_text.split("\n")
    for i, line in enumerate(lines):
        if "CalculateTotals" in line:
            # Write to file instead of print to avoid unicode issues
            with open("calculated_totals_lines.txt", "a", encoding="utf-8") as f:
                f.write(f"Line {i+1}: {line}\n")
    
    print(f"Found {sum(1 for l in lines if 'CalculateTotals' in l)} lines with CalculateTotals")
    
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)