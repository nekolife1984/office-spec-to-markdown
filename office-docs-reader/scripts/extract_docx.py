#!/usr/bin/env python3
"""Word (.docx) extraction: pandoc when available, else a stdlib fallback
(headings / paragraphs / tables / images). Can also be run standalone:

    python3 extract_docx.py <input.docx> -o <output_dir>
"""
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from extract_common import NS, esc, q, run_extraction


def docx_fallback(zf, src, out, media_dir):
    """Stdlib docx extractor: paragraphs (with heading styles), tables, images."""
    doc = ET.fromstring(zf.read("word/document.xml"))
    rels = {}
    try:
        rel_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        for rel in rel_root:
            rels[rel.get("Id")] = rel.get("Target")
    except KeyError:
        pass

    media_map = {}  # zip path -> output filename
    for name in zf.namelist():
        if name.startswith("word/media/") and not name.endswith("/"):
            fn = Path(name).name
            media_dir.mkdir(parents=True, exist_ok=True)
            (media_dir / fn).write_bytes(zf.read(name))
            media_map[name] = fn

    lines, n_head, n_table, n_img = [], 0, 0, 0

    def para_text(p):
        return "".join(t.text or "" for t in p.iter(q("w", "t")))

    def para_images(p):
        out_imgs = []
        for blip in p.iter(q("a", "blip")):
            rid = blip.get(q("r", "embed"))
            target = rels.get(rid, "")
            full = "word/" + target if not target.startswith("/") else target.lstrip("/")
            full = str(Path(full))  # normalise ../ etc. roughly
            for zp, fn in media_map.items():
                if Path(zp).name == Path(target).name:
                    out_imgs.append(fn)
        return out_imgs

    body = doc.find(q("w", "body"))
    for child in body:
        if child.tag == q("w", "p"):
            style_el = child.find(f"{q('w','pPr')}/{q('w','pStyle')}")
            style = style_el.get(q("w", "val")) if style_el is not None else ""
            text = para_text(child).strip()
            imgs = para_images(child)
            m = re.match(r"(?:Heading|見出し)\s*([1-6])", style or "")
            if m and text:
                lines.append(f"{'#' * int(m.group(1))} {text}")
                n_head += 1
            elif child.find(f".//{q('w','numPr')}") is not None and text:
                lines.append(f"- {text}")
            elif text:
                lines.append(text)
            for fn in imgs:
                lines.append(f"![図](media/{fn})")
                n_img += 1
            lines.append("")
        elif child.tag == q("w", "tbl"):
            rows = []
            for tr in child.findall(q("w", "tr")):
                row = []
                for tc in tr.findall(q("w", "tc")):
                    cell_text = "<br>".join(
                        esc(para_text(p).strip()) for p in tc.findall(q("w", "p")) if para_text(p).strip()
                    )
                    span_el = tc.find(f"{q('w','tcPr')}/{q('w','gridSpan')}")
                    span = int(span_el.get(q("w", "val"), 1)) if span_el is not None else 1
                    row.append(cell_text)
                    row.extend([""] * (span - 1))
                rows.append(row)
            if rows:
                width = max(len(r) for r in rows)
                rows = [r + [""] * (width - len(r)) for r in rows]
                lines.append("| " + " | ".join(rows[0]) + " |")
                lines.append("| " + " | ".join(["---"] * width) + " |")
                for r in rows[1:]:
                    lines.append("| " + " | ".join(r) + " |")
                lines.append("")
                n_table += 1
    md = "\n".join(lines)
    warnings = [
        "docx fallback extractor used (pandoc not found): SmartArt, text boxes, "
        "charts, headers/footers and tracked changes are NOT extracted. "
        "Install pandoc or convert those parts via PDF + screenshots."
    ]
    return md, {"engine": "stdlib-fallback", "headings": n_head, "tables": n_table,
                "images": n_img, "media_files": len(media_map)}, warnings


def convert_docx(src, out):
    media_dir = out / "media"
    stem = src.stem
    inv = {"source": src.name, "type": "docx"}
    warnings = []
    if shutil.which("pandoc"):
        media_dir.mkdir(parents=True, exist_ok=True)
        md_path = out / "extracted.md"
        subprocess.run(
            ["pandoc", str(src), "-t", "gfm", "--wrap=none",
             f"--extract-media={out}", "-o", str(md_path)],
            check=True)
        # pandoc puts images under <out>/media by default
        extracted = [p for p in out.rglob("*") if p.is_file() and p != md_path]
        inv.update({"engine": "pandoc", "media_files": len(extracted)})
        warnings.append("Check SmartArt / text boxes / Excel charts in the source: "
                        "pandoc does not convert them. Handle via the PDF route in SKILL.md.")
    else:
        with zipfile.ZipFile(src) as zf:
            md, stats, w = docx_fallback(zf, src, out, media_dir)
        (out / "extracted.md").write_text(md, encoding="utf-8")
        inv.update(stats)
        warnings.extend(w)
    inv["warnings"] = warnings
    return inv




if __name__ == "__main__":
    run_extraction(convert_docx, {".docx"})
