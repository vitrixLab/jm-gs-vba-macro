"""Dump VBA source code from .xlsm/.zip workbooks into text files.

Usage: python extract_vba.py <input1> [input2 ...]
Writes one .txt per input next to this script (input stem + '_vba.txt').
"""
import os
import sys
import zipfile

from oletools.olevba import VBA_Parser


def extract(path: str) -> str:
    parts = []
    vp = VBA_Parser(path)
    try:
        for (_, stream_path, vba_filename, vba_code) in vp.extract_macros():
            if vba_code.strip():
                parts.append(
                    "=" * 78
                    + f"\n== MODULE: {vba_filename}  (OLE stream: {stream_path})\n"
                    + "=" * 78
                    + "\n" + vba_code
                )
    finally:
        vp.close()
    return "\n\n".join(parts) + ("\n" if parts else "")


def main() -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    for arg in sys.argv[1:]:
        # .zip archives: unpack to temp and scan every member that looks like xlsm
        targets = [arg]
        if arg.lower().endswith(".zip"):
            import tempfile
            with tempfile.TemporaryDirectory() as td:
                with zipfile.ZipFile(arg) as zf:
                    zf.extractall(td)
                targets = [
                    os.path.join(root, f)
                    for root, _, files in os.walk(td)
                    for f in files if f == "vbaProject.bin"
                ]
            out = os.path.join(here, os.path.splitext(os.path.basename(arg))[0] + "_vba.txt")
            chunks = []
            for t in targets:
                chunks.append(extract(t))
            text = "\n\n".join(chunks)
        else:
            out = os.path.join(here, os.path.splitext(os.path.basename(arg))[0] + "_vba.txt")
            text = extract(arg)
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"{arg} -> {out} ({len(text)} chars)")


if __name__ == "__main__":
    main()
