import zipfile
import sys

XLsm_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.1.xlsm"
try:
    with zipfile.ZipFile(XLsm_PATH, 'r') as z:
        vba_data = z.read("xl/vbaProject.bin")
    vba_text = vba_data.decode("utf-8", errors="replace")
    
    ct_count = vba_text.count("CalculateTotals")
    fbs_count = vba_text.count("FindBlockStart")
    
    print(f"CalculateTotals occurrences: {ct_count}")
    print(f"FindBlockStart occurrences: {fbs_count}")
    
    # Also check for the specific call pattern
    call_pattern = "CalculateTotals(cell.row"
    if call_pattern in vba_text:
        print(f"Found call pattern: {call_pattern}")
    else:
        print(f"Call pattern '{call_pattern}' NOT found")
    
    # Look for CalculateTotals with just (cellRow or similar
    # Check what the actual calls look like
    for line in vba_text.split("\n"):
        if "CalculateTotals" in line and "cell" in line.lower():
            # Truncate for display
            print(f"  Call line: {line[:200]}")
    
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)