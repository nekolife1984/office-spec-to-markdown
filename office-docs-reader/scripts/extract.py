#!/usr/bin/env python3
"""Dispatcher for the office-docs-reader extraction pipeline.

Routes by file extension to the per-format modules:
    .docx       -> extract_docx.convert_docx
    .xlsx/.xlsm -> extract_xlsx.convert_xlsx
    .pdf        -> extract_pdf.extract_pdf

All formats share one output contract (see extract_common.py):
extracted.md + inventory.json + media/. Each per-format module can also
be run on its own.

Usage:
    python3 extract.py <input.docx|input.xlsx|input.pdf> -o <output_dir>
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from extract_common import finish


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()
    src = Path(args.input)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    suffix = src.suffix.lower()
    if suffix == ".docx":
        from extract_docx import convert_docx
        inv = convert_docx(src, out)
    elif suffix in (".xlsx", ".xlsm"):
        from extract_xlsx import convert_xlsx
        inv = convert_xlsx(src, out)
    elif suffix == ".pdf":
        from extract_pdf import extract_pdf
        inv = extract_pdf(src, out)
    elif suffix == ".xls":
        sys.exit("Legacy .xls is not supported. Re-save as .xlsx first.")
    elif suffix == ".doc":
        sys.exit("Legacy .doc is not supported. Re-save as .docx first.")
    else:
        sys.exit(f"Unsupported file type: {suffix}")
    finish(inv, out)


if __name__ == "__main__":
    main()
