"""Split the olevba extract into per-module .bas/.cls files for editing/re-injection."""
import os
import re

SRC = r"C:\citrixlabph\globalsmile\_vba_extract\Global-Smile_2026-v7.8_vba.txt"
OUTDIR = r"C:\citrixlabph\globalsmile\_vba_extract\bas"

with open(SRC, encoding="utf-8") as f:
    text = f.read()

blocks = re.split(r"={78}\r?\n== MODULE: (.+?)  \(OLE stream: (.+?)\)\r?\n={78}\r?\n", text)
# blocks: [pre, name1, stream1, code1, name2, stream2, code2, ...]
os.makedirs(OUTDIR, exist_ok=True)
count = 0
for i in range(1, len(blocks), 3):
    name, stream, code = blocks[i], blocks[i + 1], blocks[i + 2]
    ext = os.path.splitext(name)[1].lower()
    out = os.path.join(OUTDIR, name.replace("/", "_"))
    with open(out, "w", encoding="utf-8", newline="") as f:
        f.write(code)
    print(f"{name}  <- {stream}  ({len(code)} chars) -> {out}")
    count += 1
print(f"total modules: {count}")
