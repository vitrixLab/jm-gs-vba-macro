"""Read-only VBA extraction for the Gate 2 remediation pass.

Extracts every VBA module from Global-Smile_2026-v8.4.xlsm into gate2_work/vba_baseline/
and writes a hash manifest. The workbook binary is never modified.
"""
import hashlib
import io
import json
import logging
import os
import re
import sys
import zipfile

from oletools.olevba import VBA_Parser

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(HERE, "Global-Smile_2026-v8.4.xlsm")
OUTDIR = os.path.join(HERE, "gate2_work", "vba_baseline")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> None:
    os.makedirs(OUTDIR, exist_ok=True)
    warnings = []
    log_capture = io.StringIO()
    handler = logging.StreamHandler(log_capture)
    handler.setLevel(logging.WARNING)
    logging.getLogger("olevba").addHandler(handler)

    wb_hash_before = sha256_file(WORKBOOK)

    vp_bin_hash = None
    project_name = None
    with zipfile.ZipFile(WORKBOOK) as z:
        if "xl/vbaProject.bin" in z.namelist():
            data = z.read("xl/vbaProject.bin")
            vp_bin_hash = sha256_bytes(data)
            import olefile
            ole = olefile.OleFileIO(io.BytesIO(data))
            if ole.exists("PROJECT"):
                txt = ole.openstream("PROJECT").read().decode("latin-1", "replace")
                m = re.search(r'^Name="(.*?)"', txt, re.M)
                if m:
                    project_name = m.group(1)
            ole.close()

    modules = []
    vp = VBA_Parser(WORKBOOK)
    try:
        for (_, stream_path, vba_filename, vba_code) in vp.extract_macros():
            if not vba_code.strip():
                continue
            modules.append({
                "name": vba_filename,
                "stream": stream_path,
                "chars": len(vba_code),
                "lines": vba_code.count("\n") + (0 if vba_code.endswith("\n") else 1),
                "sha256": sha256_bytes(vba_code.encode("utf-8", "replace")),
            })
            out = os.path.join(OUTDIR, vba_filename.replace("/", "_"))
            with open(out, "w", encoding="utf-8", newline="") as fh:
                fh.write(vba_code)
    finally:
        vp.close()

    warnings.extend([l for l in log_capture.getvalue().splitlines() if l.strip()])

    wb_hash_after = sha256_file(WORKBOOK)
    manifest = {
        "workbook": os.path.basename(WORKBOOK),
        "workbook_sha256": wb_hash_before,
        "workbook_sha256_after_extraction": wb_hash_after,
        "workbook_unmodified": wb_hash_before == wb_hash_after,
        "vbaProject_bin_sha256": vp_bin_hash,
        "vba_project_name": project_name,
        "module_count": len(modules),
        "modules": modules,
        "extraction_warnings": warnings,
    }
    with open(os.path.join(OUTDIR, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2)

    print(f"workbook sha256      : {wb_hash_before}")
    print(f"vbaProject.bin sha256: {vp_bin_hash}")
    print(f"vba project name     : {project_name}")
    print(f"modules extracted    : {len(modules)}")
    for m in sorted(modules, key=lambda x: x["name"]):
        print(f"  {m['name']:<40} lines={m['lines']:>5}  sha={m['sha256'][:16]}")
    print(f"workbook unmodified  : {manifest['workbook_unmodified']}")
    if warnings:
        print("warnings:")
        for w in warnings:
            print(f"  {w}")


if __name__ == "__main__":
    sys.exit(main())
