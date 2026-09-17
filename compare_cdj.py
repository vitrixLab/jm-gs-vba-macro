# Patch PJ join logic on every projection sheet (excludes the pilot AC sheet)
PATCH_FILE = r"C:\citrixlabph\globalsmile\patch_pj_join.vbs"

sheets_txt = {
    "PJ":       True,
    "PJ2":       True,
    "PJ Non-Vat": True,
    "PJ Non-Vat2": True,
    "AC":        False,  # pilot table intentionally excluded
}
END = r"""                If Not isFirst Then
                    rng.Range("B" & L & 21).Value2 = tol
                    If Not rs.BOF Then nsp = rng.Range("B" & L & 21).Offset(0, -18).Text
                    If Not IsDBNull(nsp) And nsp <> "" Then
                        If Not IsDBNull(tol) And tol <> "" Then
                            Set rn = rng.Range("B" & L & 21).Find(What:=tol, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        Else
                            Set rn = Nothing
                        End If
                        If Not rn Is Nothing Then
                            rn.Offset(0, 1).Value2 = nsp
                            If Not IsDBNull(ctr) And ctr <> "" Then rn.Offset(0, 2).Value2 = ctr
                            inc = inc + 1
                        End If
                    End If
                End If
            End If
        Len = Len - 1
    End If
End Sub"""
