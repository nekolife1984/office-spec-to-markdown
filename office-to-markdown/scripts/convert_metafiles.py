#!/usr/bin/env python3
"""Convert EMF/WMF images in an extraction media/ directory to SVG/PNG.

Browsers and Markdown renderers cannot display EMF/WMF, but Office
documents frequently embed figures in those formats. EMF/WMF are vector
formats, so SVG keeps them scalable and faithful; PNG is the fallback
when the SVG conversion fails.

Conversion uses LibreOffice headless (`soffice`), which must be
installed. The original metafiles are kept untouched; converted files
are written next to them as <stem>.svg (or <stem>.png).

Usage:
    python3 convert_metafiles.py <extract-dir or media-dir>

Exit code 1 when a metafile could be converted to neither SVG nor PNG.
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

METAFILE_SUFFIXES = {".emf", ".wmf"}


def convert_one(soffice, src, media_dir, fmt):
    """Convert src to fmt with LibreOffice; return the output path or None."""
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp) / "profile"
        work = Path(tmp) / "out"
        work.mkdir()
        cmd = [soffice, f"-env:UserInstallation=file://{profile}",
               "--headless", "--convert-to", fmt, "--outdir", str(work), str(src)]
        try:
            subprocess.run(cmd, capture_output=True, timeout=180, check=False)
        except (OSError, subprocess.TimeoutExpired):
            return None
        produced = work / (src.stem + "." + fmt)
        if not produced.exists() or produced.stat().st_size == 0:
            return None
        if fmt == "svg":
            try:
                ET.parse(produced)
            except ET.ParseError:
                return None
        if fmt == "png" and produced.read_bytes()[:8] != b"\x89PNG\r\n\x1a\n":
            return None
        dest = media_dir / produced.name
        shutil.move(str(produced), str(dest))
        return dest


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    target = Path(sys.argv[1])
    media_dir = target / "media" if (target / "media").is_dir() else target
    if not media_dir.is_dir():
        sys.exit(f"media directory not found: {target}")
    metafiles = sorted(p for p in media_dir.iterdir()
                       if p.suffix.lower() in METAFILE_SUFFIXES)
    if not metafiles:
        print("no EMF/WMF files found")
        return
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("LibreOffice (soffice) is required to convert EMF/WMF but was not found")
    failed = []
    for src in metafiles:
        dest = convert_one(soffice, src, media_dir, "svg")
        kind = "svg"
        if dest is None:
            dest = convert_one(soffice, src, media_dir, "png")
            kind = "png (fallback)"
        if dest is None:
            failed.append(src.name)
            print(f"FAILED: {src.name} could not be converted to SVG or PNG")
        else:
            print(f"converted: {src.name} -> {dest.name} [{kind}]")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
