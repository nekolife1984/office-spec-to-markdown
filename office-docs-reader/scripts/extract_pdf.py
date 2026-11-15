#!/usr/bin/env python3
"""PDF extraction (PyMuPDF): per-page text blocks in reading order, detected
tables and embedded images. Can also be run standalone:

    python3 extract_pdf.py <input.pdf> -o <output_dir>
"""
import sys
from pathlib import Path

from extract_common import esc, run_extraction


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
        # Span-level extraction carries font decoration: bold/italic runs
        # become Markdown; coloured runs are counted (Markdown has no
        # colour) so the restructuring step can make their meaning explicit.
        styled = {"bold": 0, "italic": 0, "colored": 0}
        blocks = []
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0:
                continue
            x0, y0, x1, y1 = block["bbox"]
            center = fitz.Point((x0 + x1) / 2, (y0 + y1) / 2)
            if any(center in r for r in table_rects):
                continue
            lines_md = []
            plain_len = 0
            for line in block["lines"]:
                runs = []  # (bold, italic, text) with adjacent same-style spans merged
                for span in line["spans"]:
                    text = span["text"]
                    if not text:
                        continue
                    plain_len += len(text)
                    fname = span.get("font", "").lower()
                    bold = bool(span.get("flags", 0) & 16) or "bold" in fname
                    italic = bool(span.get("flags", 0) & 2) or "italic" in fname or "oblique" in fname
                    if text.strip():
                        if bold: styled["bold"] += 1
                        if italic: styled["italic"] += 1
                        if span.get("color", 0) not in (0, None): styled["colored"] += 1
                    if runs and runs[-1][0] == bold and runs[-1][1] == italic:
                        runs[-1] = (bold, italic, runs[-1][2] + text)
                    else:
                        runs.append((bold, italic, text))
                line_md = ""
                for bold, italic, text in runs:
                    t = text
                    if t.strip():
                        if italic: t = f"*{t}*"
                        if bold: t = f"**{t}**"
                    line_md += t
                if line_md.strip():
                    lines_md.append(line_md.strip())
            if lines_md:
                blocks.append((round(y0 / 10), x0, "\n".join(lines_md)))
        blocks.sort()
        p_inv["text_chars"] = sum(len(t) for _, _, t in blocks)
        p_inv["styled_spans"] = styled
        if styled["colored"]:
            warnings.append(f"Page {page_no} has {styled['colored']} coloured text run(s), "
                            "which Markdown cannot represent. Make their meaning explicit "
                            "when restructuring.")

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




if __name__ == "__main__":
    run_extraction(extract_pdf, {".pdf"})
