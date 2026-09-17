P = r'C:\citrixlabph\globalsmile\_vba_extract\inject_and_test.ps1'
s = open(P, encoding='utf-8-sig').read()
old = None
for line in s.split('\n'):
    if 'T-write count' in line:
        old = line
        break
print(repr(old))
new = '    L "modEngine T-write count: $([regex]::Matches($meCode, \'netAmount\').Count) (expect >=4)"'
s = s.replace(old, new)
open(P, 'w', encoding='utf-8').write(s)
print("patched")
