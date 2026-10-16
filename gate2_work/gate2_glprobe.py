"""Gate 2 helper: inspect the GL sheet layout and CDJ/GL structures."""
import sys

import win32com.client

CANDIDATE = 'D:/citrixlabph/globalsmile/Global-Smile_2026-v8.4-fix.xlsm'
SOURCE = 'D:/citrixlabph/globalsmile/Global-Smile_2026-v8.4.xlsm'


def open_wb(path):
    xl = win32com.client.DispatchEx('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    wb = xl.Workbooks.Open(path, 0, True)
    return xl, wb


def main():
    xl, wb = open_wb(CANDIDATE)
    try:
        wsheets = wb.Worksheets
        print('--- sheets ---')
        for i in range(1, wsheets.Count + 1):
            ws = wsheets.Item(i)
            print('  %-2d %-16s dims=%s' % (i, ws.Name, ws.UsedRange.Address))

        gl = wb.Worksheets('GL')
        ur = gl.UsedRange
        print()
        print('GL UsedRange: %s  rows=%d  cols=%d'
              % (ur.Address, ur.Rows.Count, ur.Columns.Count))
        print('--- GL rows 1..34 (A..H) ---')
        for r in range(1, 35):
            vals = [gl.Cells(r, c).Value for c in range(1, 9)]
            if any(v is not None and v != '' for v in vals):
                print('  r%-4d %s' % (r, vals))

        print('--- GL rows 200..230 (A..H) ---')
        for r in range(200, 231):
            vals = [gl.Cells(r, c).Value for c in range(1, 9)]
            if any(v is not None and v != '' for v in vals):
                print('  r%-4d %s' % (r, vals))
    finally:
        wb.Close(False)
        xl.Quit()
    return 0


if __name__ == '__main__':
    sys.exit(main())
