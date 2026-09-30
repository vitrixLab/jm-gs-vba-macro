import json
p = 'v8.2_review/review_v82_structure.json'
d = json.load(open(p, encoding='utf-8-sig'))
print('keys=', list(d.keys()))
for k, v in d.items():
    print('\n##', k)
    s = json.dumps(v, ensure_ascii=False)
    print(s[:4000])
