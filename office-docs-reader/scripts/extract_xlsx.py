#!/usr/bin/env python3
"""Excel (.xlsx/.xlsm) extraction: cells via openpyxl, plus shape/text-box
text, cell comments and anchored images parsed from the file's XML parts.
Can also be run standalone:

    python3 extract_xlsx.py <input.xlsx> -o <output_dir>
"""
import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from extract_common import cell_ref, esc, q, run_extraction


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


def _wrap_md(text, bold=False, italic=False, strike=False, underline=False):
    """Wrap one text run with Markdown decoration for the given font flags."""
    if not text.strip():
        return text
    if strike:
        text = f"~~{text}~~"
    if underline:
        text = f"<u>{text}</u>"
    if italic:
        text = f"*{text}*"
    if bold:
        text = f"**{text}**"
    return text


def _is_colored(font):
    """True for visually chromatic font colours.

    Grayscale colours (black, white, grays such as the near-black body
    text 262626 or white header text) are document design, not semantic
    colour coding, so they are not counted.
    """
    color = getattr(font, "color", None)
    if color is None or getattr(color, "type", None) != "rgb":
        return False
    rgb = str(getattr(color, "rgb", "") or "")[-6:]
    if len(rgb) != 6:
        return False
    try:
        r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    return max(r, g, b) - min(r, g, b) > 16


def cell_md(cell, counts):
    """Cell value as Markdown, carrying font decoration.

    Bold/italic/strike/underline become Markdown (or <u>); font colour
    cannot be represented in Markdown, so coloured runs are only counted
    in the inventory (counts["colored"]) for the restructuring step to
    make their meaning explicit. Rich-text runs are decorated per run.
    """
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, CellRichText):
        parts = []
        for block in v:
            if isinstance(block, TextBlock):
                if not block.text.strip():
                    parts.append(esc(block.text))
                    continue
                f = block.font
                b = bool(f and f.b)
                i = bool(f and f.i)
                st = bool(f and f.strike)
                u = bool(f and f.u)
                if b: counts["bold"] += 1
                if i: counts["italic"] += 1
                if st: counts["strike"] += 1
                if u: counts["underline"] += 1
                if f and _is_colored(f): counts["colored"] += 1
                parts.append(_wrap_md(esc(block.text), b, i, st, u))
            else:
                parts.append(esc(str(block)))
        return "".join(parts)
    if isinstance(v, str):
        if not v.strip():
            return ""
        f = cell.font
        b = bool(f and f.b)
        i = bool(f and f.i)
        st = bool(f and f.strike)
        u = bool(f and f.u)
        if b: counts["bold"] += 1
        if i: counts["italic"] += 1
        if st: counts["strike"] += 1
        if u: counts["underline"] += 1
        if f and _is_colored(f): counts["colored"] += 1
        return _wrap_md(esc(v), b, i, st, u)
    return esc(v)


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
        media_duplicates = []
        if media_files:
            media_dir.mkdir(parents=True, exist_ok=True)
            for n in media_files:
                data = zf.read(n)
                digest = hashlib.sha256(data).hexdigest()
                if digest in media_bytes:
                    # identical image embedded more than once: keep one file
                    media_duplicates.append(Path(n).name)
                    continue
                (media_dir / Path(n).name).write_bytes(data)
                media_bytes[digest] = Path(n).name

        try:
            wb = load_workbook(src, data_only=True, rich_text=True)
        except TypeError:
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
            styled = {"bold": 0, "italic": 0, "strike": 0, "underline": 0, "colored": 0}
            grid = []
            for row in ws.iter_rows():
                grid.append([cell_md(c, styled) for c in row])
            s_inv["styled_cells"] = styled
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

        colored_total = sum(s.get("styled_cells", {}).get("colored", 0) for s in inv["sheets"])
        if colored_total:
            warnings.append(f"{colored_total} cell(s) use font colour, which Markdown cannot "
                            "represent. They are counted per sheet in styled_cells; make the "
                            "meaning explicit (a note or a status column) when restructuring.")
        inv["media_files"] = len(media_bytes)
        if media_duplicates:
            inv["media_duplicates"] = media_duplicates
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




if __name__ == "__main__":
    run_extraction(convert_xlsx, {".xlsx", ".xlsm"})
