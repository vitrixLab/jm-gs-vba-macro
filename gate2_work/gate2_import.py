"""Gate 2 — import patched VBA modules into Global-Smile_2026-v8.4-fix.xlsm.

Standard modules are replaced with Remove + Import; the Sheet9 document class is
replaced through CodeModule (the same mechanism apply_and_test_v833.py uses for
Sheet11/Sheet12, since document classes cannot be re-imported by name).

Ends with a VBE compile of the project so a compile break fails the run.
"""
import io
import os
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
PATCHED = os.path.join(ROOT, "gate2_work", "vba_patched")


def strip_attributes(text):
    out = []
    for line in text.splitlines():
        if line.strip().upper().startswith("ATTRIBUTE VB_"):
            continue
        out.append(line)
    return "\n".join(out)


def main():
    print("Opening %s" % CANDIDATE)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1  # msoAutomationSecurityLow: allow VBProject access + macros for the test pass
    wb = xl.Workbooks.Open(CANDIDATE, 0, False)
    vbp = wb.VBProject

    # 1. Standard modules: Remove then Import.
    for name in ("modEngine", "modPJSync", "modGLWorkbookMap"):
        path = os.path.join(PATCHED, name + ".bas")
        for comp in list(vbp.VBComponents):
            if comp.Name == name:
                vbp.VBComponents.Remove(comp)
                print("  removed  %s" % name)
                break
        vbp.VBComponents.Import(path)
        print("  imported %s" % name)

    # 2. Document class: swap the CodeModule body.
    with io.open(os.path.join(PATCHED, "Sheet9.cls"), "r", encoding="utf-8", newline="") as fh:
        body = strip_attributes(fh.read())
    comp = vbp.VBComponents("Sheet9")
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(body)
    print("  updated Sheet9 CodeModule (%d lines)" % cm.CountOfLines)

    wb.Save()
    print("saved")

    # 3. Compile the project.
    print("compiling...")
    try:
        xl.VBE.CommandBars.Item("Standard").Controls.Item("Compile VBAProject").Execute()
        print("COMPILE: PASS")
        result = 0
    except Exception as exc:
        print("COMPILE: FAIL -> %s" % exc)
        result = 1

    wb.Close(True)
    xl.Quit()
    return result


if __name__ == "__main__":
    sys.exit(main())
