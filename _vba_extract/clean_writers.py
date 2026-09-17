P = r'C:\citrixlabph\globalsmile\_vba_extract\bas\modEngine.bas'
s = open(P, encoding='utf-8').read()
orig = s
old = """            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = ""

            End If

            If False Then ' (removed Sundry-description branch)

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

            End If

            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated from PJ Row"""
new = """            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

            End If

            .Cells(targetRow, CDJ_COL_NOTE).value = "Updated from PJ Row"""
c = s.count(old)
s = s.replace(old, new)
old2 = """            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = ""

            End If

            If False Then ' (removed Sundry-description branch)

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

            End If

            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted from PJ Row"""
new2 = """            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

                .Cells(targetRow, expenseCol).value = netAmount

                .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

            End If

            .Cells(targetRow, CDJ_COL_NOTE).value = "Inserted from PJ Row"""
c2 = s.count(old2)
s = s.replace(old2, new2)
open(P, "w", encoding="utf-8").write(s)
print(f"UPDATE cleaned x{c}; INSERT cleaned x{c2}; changed={s != orig}")
