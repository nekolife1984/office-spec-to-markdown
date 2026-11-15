#!/usr/bin/env python3
"""Verify a converted Markdown against its inventory.json baseline.

Checks:
  - every image link in the .md points to a file that exists
  - every extracted media file is referenced at least once
  - prints the count checklist (headings / tables / images / shapes / comments)
    to compare with the source document by eye

Usage:
    python3 verify.py <output_dir>   # dir containing <stem>.md + inventory.json

Exit code 1 when a link is broken or a media file is unreferenced.
"""
import json
import re
import sys
from pathlib import Path


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
    inv_path = out / "inventory.json"
    if not inv_path.exists():
        sys.exit(f"inventory.json not found in {out}")
    inv = json.loads(inv_path.read_text(encoding="utf-8"))
    md_files = [p for p in out.glob("*.md")]
    if not md_files:
        sys.exit(f"no .md file in {out}")
    md_path = md_files[0]
    md = md_path.read_text(encoding="utf-8")

    problems = []
    links = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", md)
    for link in links:
        if link.startswith("http"):
            continue
        if not (out / link).exists():
            problems.append(f"broken image link: {link}")
    media_dir = out / "media"
    media_files = [p.name for p in media_dir.iterdir()] if media_dir.exists() else []
    for fn in media_files:
        if fn not in md:
            problems.append(f"media file not referenced in markdown: {fn}")

    print(f"== checklist for {inv.get('source')} ({inv.get('type')}, engine: {inv.get('engine', 'n/a')}) ==")
    if inv["type"] == "docx":
        print(f"headings: {inv.get('headings', 'see pandoc output')} / tables: {inv.get('tables', 'count in md')} / "
              f"media files: {inv.get('media_files')}")
    else:
        for s in inv.get("sheets", []):
            print(f"sheet '{s['name']}': rows={s['table_rows']} merged={len(s['merged_ranges'])} "
                  f"shapes={s['shapes_with_text']} comments={s['comments']} images={s['images']}")
        print(f"media files: {inv.get('media_files')} / charts (not converted): {inv.get('charts')}")
    print(f"image links in md: {len(links)}")
    for w in inv.get("warnings", []):
        print(f"WARNING: {w}")
    if problems:
        for p in problems:
            print(f"PROBLEM: {p}")
        sys.exit(1)
    print("OK: all image links resolve and all media files are referenced.")


if __name__ == "__main__":
    main()
