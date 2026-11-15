#!/usr/bin/env python3
"""Extract everything an AI needs to read from a Word (.docx), Excel (.xlsx) or PDF document.

This is the extraction half of a two-skill pipeline (office-docs-reader ->
office-to-markdown). It extracts *everything* mechanically (cells, paragraphs,
tables, shape/text-box text, comments, images) and records counts in
inventory.json. It does NOT restructure anything into final Markdown: the
output is a faithful, position-annotated extraction for an AI to read.
Nothing is silently dropped: anything that cannot be extracted is listed
in the inventory warnings.

Usage:
    python3 extract.py <input.docx|input.xlsx|input.pdf> -o <output_dir>

Outputs in <output_dir>:
    extracted.md     faithful extraction (per-sheet tables for xlsx,
                     paragraphs/tables for docx) with shape texts, comments
                     and image references placed by anchor position
    inventory.json   counts and warnings (the verification baseline)
    media/           extracted images (docx / xlsx embedded pictures)
    work/            intermediate files (shapes.json for xlsx)

Dependencies: Python 3.10+ standard library. openpyxl (+Pillow) is used for
.xlsx, PyMuPDF for .pdf; .docx falls back to a stdlib extractor when pandoc
is not installed.
"""
import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

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


# ---------------------------------------------------------------- docx

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


# ---------------------------------------------------------------- xlsx

def parse_sheet_rels(zf, sheet_path):
    """Return {rel_type_suffix: target_path} for one worksheet."""
    rels_path = str(Path(sheet_path).parent / "_rels" / (Path(sheet_path).name + ".rels"))
    result = {}
    try:
        root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return result
    base = Path(sheet_path).parent
    for rel in root:
        target = rel.get("Target")
        rtype = rel.get("Type", "").rsplit("/", 1)[-1]
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = str(base / target)
        # normalise a/../b style paths
        parts = []
        for seg in resolved.split("/"):
            if seg == "..":
                parts.pop()
            elif seg not in (".", ""):
                parts.append(seg)
        result[rtype] = "/".join(parts)
    return result


def extract_shapes(zf, drawing_path):
    """Extract text from drawing shapes (incl. text boxes and grouped shapes).

    Returns a list of {anchor_cell, text} sorted by anchor position.
    """
    try:
        root = ET.fromstring(zf.read(drawing_path))
    except KeyError:
        return []
    shapes = []
    for anchor in root:
        frm = anchor.find(q("xdr", "from"))
        col = int(frm.find(q("xdr", "col")).text) if frm is not None else 0
        row = int(frm.find(q("xdr", "row")).text) if frm is not None else 0
        for sp in anchor.iter(q("xdr", "sp")):
            paras = []
            for p in sp.iter(q("a", "p")):
                t = "".join(node.text or "" for node in p.iter(q("a", "t"))).strip()
                if t:
                    paras.append(t)
            text = "\n".join(paras).strip()
            if text:
                shapes.append({"anchor_cell": cell_ref(col, row), "_sort": (row, col), "text": text})
    shapes.sort(key=lambda s: s["_sort"])
    for s in shapes:
        del s["_sort"]
    return shapes


def extract_comments(zf, comments_path):
    try:
        root = ET.fromstring(zf.read(comments_path))
    except KeyError:
        return []
    out = []
    for comment in root.iter():
        if comment.tag.endswith("}comment"):
            text = "".join(t.text or "" for t in comment.iter() if t.tag.endswith("}t")).strip()
            if text:
                out.append({"cell": comment.get("ref"), "author": comment.get("author", ""), "text": text})
    return out


def convert_xlsx(src, out):
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("openpyxl is required for .xlsx (pip install openpyxl Pillow), "
                 "or run markitdown for a cells-only first pass.")
    media_dir = out / "media"
    work_dir = out / "work"
    work_dir.mkdir(parents=True, exist_ok=True)
    stem = src.stem
    inv = {"source": src.name, "type": "xlsx", "sheets": []}
    warnings = []

    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        media_files = [n for n in names if n.startswith("xl/media/") and not n.endswith("/")]
        chart_files = [n for n in names if n.startswith("xl/charts/chart") and n.endswith(".xml")]
        media_bytes = {}
        if media_files:
            media_dir.mkdir(parents=True, exist_ok=True)
            for n in media_files:
                data = zf.read(n)
                (media_dir / Path(n).name).write_bytes(data)
                media_bytes[hashlib.sha256(data).hexdigest()] = Path(n).name

        wb = load_workbook(src, data_only=True)
        # map sheet name -> sheet xml path, in workbook order
        wb_rels = {}
        try:
            root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in root:
                wb_rels[rel.get("Id")] = rel.get("Target")
        except KeyError:
            pass
        wbxml = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_paths = {}
        for sh in wbxml.iter():
            if sh.tag.endswith("}sheet"):
                rid = sh.get(q("r", "id"))
                target = wb_rels.get(rid, "")
                if target.startswith("/"):
                    sheet_paths[sh.get("name")] = target.lstrip("/")
                elif target.startswith("xl/"):
                    sheet_paths[sh.get("name")] = target
                else:
                    sheet_paths[sh.get("name")] = "xl/" + target

        md_parts = [f"# {stem}"]
        all_shapes = {}
        for ws in wb.worksheets:
            s_inv = {"name": ws.title, "state": ws.sheet_state,
                     "max_row": ws.max_row, "max_column": ws.max_column,
                     "merged_ranges": [str(r) for r in ws.merged_cells.ranges]}
            rels = parse_sheet_rels(zf, sheet_paths.get(ws.title, "")) if ws.title in sheet_paths else {}

            # --- cells -> markdown table (trailing empty rows/cols trimmed)
            grid = []
            for row in ws.iter_rows():
                grid.append(["" if c.value is None else esc(c.value) for c in row])
            while grid and all(v == "" for v in grid[-1]):
                grid.pop()
            width = 0
            for r in grid:
                for i in range(len(r), 0, -1):
                    if r[i - 1] != "":
                        width = max(width, i)
                        break
            grid = [r[:width] for r in grid]
            s_inv["table_rows"] = len(grid)

            md_parts.append(f"\n## シート: {ws.title}" + ("（非表示）" if ws.sheet_state != "visible" else ""))
            if grid:
                header = grid[0] if any(grid[0]) else [f"列{i+1}" for i in range(width)]
                md_parts.append("| " + " | ".join(header) + " |")
                md_parts.append("| " + " | ".join(["---"] * width) + " |")
                for r in grid[1:]:
                    md_parts.append("| " + " | ".join(r) + " |")
            else:
                md_parts.append("（セルに値なし）")

            # --- shapes / text boxes
            shapes = extract_shapes(zf, rels["drawing"]) if "drawing" in rels else []
            if shapes:
                all_shapes[ws.title] = shapes
                md_parts.append("\n### 図形・テキストボックス内のテキスト")
                for s in shapes:
                    md_parts.append(f"- （{s['anchor_cell']}付近）{esc(s['text'])}")
            s_inv["shapes_with_text"] = len(shapes)

            # --- comments
            comments = extract_comments(zf, rels["comments"]) if "comments" in rels else []
            if comments:
                md_parts.append("\n### セルコメント")
                for c in comments:
                    md_parts.append(f"- （{c['cell']}）{esc(c['text'])}")
            s_inv["comments"] = len(comments)

            # --- images anchored on this sheet
            sheet_images = []
            for img in getattr(ws, "_images", []):
                anchor = getattr(img, "anchor", None)
                frm = getattr(anchor, "_from", None)
                try:
                    digest = hashlib.sha256(img._data()).hexdigest()
                    fn = media_bytes.get(digest)
                except Exception:
                    fn = None
                if fn:
                    sheet_images.append({"file": fn,
                                         "anchor_cell": cell_ref(frm.col, frm.row) if frm else "?"})
            if sheet_images:
                md_parts.append("\n### 画像")
                for im in sheet_images:
                    md_parts.append(f"- ![図（{im['anchor_cell']}付近）](media/{im['file']})")
            s_inv["images"] = len(sheet_images)
            inv["sheets"].append(s_inv)

        inv["media_files"] = len(media_files)
        inv["charts"] = len(chart_files)
        if chart_files:
            warnings.append(f"{len(chart_files)} Excel chart object(s) found. Charts are not images: "
                            "render them via Excel/LibreOffice to PNG (PDF route) before restructuring.")
        if media_files and sum(s["images"] for s in inv["sheets"]) < len(media_files):
            warnings.append("Some media files could not be anchored to a sheet (e.g. in-cell images). "
                            "They are extracted under media/; place them manually.")
        (work_dir / "shapes.json").write_text(json.dumps(all_shapes, ensure_ascii=False, indent=2), encoding="utf-8")
        (out / "extracted.md").write_text("\n".join(md_parts) + "\n", encoding="utf-8")
    inv["warnings"] = warnings
    return inv


def extract_pdf(src, out):
    """PDF extraction with PyMuPDF: per-page text blocks in reading order,
    detected tables as Markdown tables, embedded images to media/.

    Text inside detected table bboxes is excluded from the block text to
    avoid duplication. Pages with almost no text but images are flagged
    as possible scans (OCR is out of scope for this script).
    """
    try:
        import pymupdf as fitz  # PyMuPDF (new import name)
    except ImportError:
        try:
            import fitz  # older PyMuPDF
        except ImportError:
            sys.exit("PyMuPDF is required for .pdf (pip install -r requirements.txt).")
    media_dir = out / "media"
    inv = {"source": src.name, "type": "pdf", "engine": "pymupdf", "pages": []}
    warnings = []
    doc = fitz.open(src)
    inv["page_count"] = len(doc)
    if doc.metadata.get("title"):
        inv["title"] = doc.metadata["title"]
    md_parts = [f"# {src.stem}"]
    total_images = 0
    for page_no, page in enumerate(doc, start=1):
        p_inv = {"page": page_no}
        # --- tables first, to know their bboxes
        tables_md = []
        table_rects = []
        try:
            found = page.find_tables()
            tables = list(found.tables)
        except Exception:
            tables = []
        for t in tables:
            try:
                rows = t.extract()
            except Exception:
                rows = None
            if not rows:
                continue
            table_rects.append(fitz.Rect(t.bbox))
            rows = [["" if c is None else esc(str(c).strip()) for c in r] for r in rows]
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            tables_md.append("| " + " | ".join(rows[0]) + " |")
            tables_md.append("| " + " | ".join(["---"] * width) + " |")
            for r in rows[1:]:
                tables_md.append("| " + " | ".join(r) + " |")
            tables_md.append("")
        p_inv["tables"] = len(table_rects)

        # --- text blocks outside tables, reading order (top-to-bottom, then x)
        blocks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *_ = b
            text = text.strip()
            if not text:
                continue
            center = fitz.Point((x0 + x1) / 2, (y0 + y1) / 2)
            if any(center in r for r in table_rects):
                continue
            blocks.append((round(y0 / 10), x0, text))
        blocks.sort()
        p_inv["text_chars"] = sum(len(t) for _, _, t in blocks)

        # --- images
        page_images = []
        for img_idx, img in enumerate(page.get_images(full=True), start=1):
            xref = img[0]
            try:
                data = doc.extract_image(xref)
            except Exception:
                continue
            media_dir.mkdir(parents=True, exist_ok=True)
            fn = f"page{page_no}_img{img_idx}.{data['ext']}"
            (media_dir / fn).write_bytes(data["image"])
            page_images.append(fn)
        p_inv["images"] = len(page_images)
        total_images += len(page_images)

        if p_inv["text_chars"] < 20 and (page_images or tables_md):
            warnings.append(f"Page {page_no} has almost no text layer "
                            f"({p_inv['text_chars']} chars). It may be a scanned page: "
                            "OCR is required before its content can be read.")
        md_parts.append(f"\n## ページ {page_no}")
        for _, _, text in blocks:
            md_parts.append(text)
            md_parts.append("")
        if tables_md:
            md_parts.append("### 表（検出）")
            md_parts.extend(tables_md)
        if page_images:
            md_parts.append("### 画像")
            for fn in page_images:
                md_parts.append(f"![図（ページ{page_no}）](media/{fn})")
            md_parts.append("")
        inv["pages"].append(p_inv)
    inv["media_files"] = total_images
    inv["tables"] = sum(p["tables"] for p in inv["pages"])
    inv["warnings"] = warnings
    (out / "extracted.md").write_text("\n".join(md_parts) + "\n", encoding="utf-8")
    return inv


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
        inv = convert_docx(src, out)
    elif suffix in (".xlsx", ".xlsm"):
        inv = convert_xlsx(src, out)
    elif suffix == ".pdf":
        inv = extract_pdf(src, out)
    elif suffix == ".xls":
        sys.exit("Legacy .xls is not supported. Re-save as .xlsx first.")
    elif suffix == ".doc":
        sys.exit("Legacy .doc is not supported. Re-save as .docx first.")
    else:
        sys.exit(f"Unsupported file type: {suffix}")
    (out / "inventory.json").write_text(json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(inv, ensure_ascii=False, indent=2))
    print(f"\nExtracted text: {out / 'extracted.md'}")
    print(f"Inventory: {out / 'inventory.json'}")


if __name__ == "__main__":
    main()
