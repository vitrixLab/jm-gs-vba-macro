P = r'C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas'
L = open(P, encoding='utf-8').read().split('\n')
for i, l in enumerate(L, 1):
    s = l.strip()
    if 'CDJ_COL_DAY' in s and 'Private Const' not in s:
        print(f'{i}: DAY-WRITE :: {s}')
    if 'netAmount + whtax' in s:
        print(f'{i}: NET+WHTAX :: {s}')
    if '.Cells(targetRow, expenseCol).value' in s:
        print(f'{i}: EXPENSE-WRITE :: {s}')
    if 'CDJ_COL_DEBIT).value' in s and 'Private Const' not in s:
        print(f'{i}: DEBIT-WRITE :: {s}')
