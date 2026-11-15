#!/usr/bin/env python3
"""Shared helpers for the per-format extractors (extract_docx / extract_xlsx /
extract_pdf). The output contract is common to all formats:

    extracted.md     faithful extraction
    inventory.json   counts and warnings (the verification baseline)
    media/           extracted images
"""
import argparse
import json
import sys
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}


def q(ns, tag):
    return f"{{{NS[ns]}}}{tag}"


def cell_ref(col, row):
    """0-based col/row -> A1 style reference."""
    letters = ""
    c = col
    while True:
        letters = chr(ord("A") + c % 26) + letters
        c = c // 26 - 1
        if c < 0:
            break
    return f"{letters}{row + 1}"


def esc(text):
    return str(text).replace("|", "\\|").replace("\r\n", "<br>").replace("\n", "<br>")


def run_extraction(convert, suffixes):
    """Common CLI entry point used by each per-format module."""
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()
    src = Path(args.input)
    if src.suffix.lower() not in suffixes:
        sys.exit(f"This module handles {sorted(suffixes)} only: {src.suffix}")
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    inv = convert(src, out)
    finish(inv, out)


def finish(inv, out):
    (out / "inventory.json").write_text(
        json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(inv, ensure_ascii=False, indent=2))
    print(f"\nExtracted text: {out / 'extracted.md'}")
    print(f"Inventory: {out / 'inventory.json'}")
