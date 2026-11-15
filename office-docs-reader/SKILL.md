---
name: office-docs-reader
description: Read Word (.docx) and Excel (.xlsx) documents completely for AI analysis — extract paragraphs, headings, tables, cell values, shape and text-box text, cell comments, embedded images and charts information, with positions and counts. Use whenever an AI agent needs to read, understand, search, summarize or answer questions about a .docx/.xlsx file (including 仕様書・設計書), or as the extraction step before converting Office documents to Markdown with the office-to-markdown skill. Do NOT rely on reading the binary directly or on tools that only extract cell/body text — shapes, comments and images are silently dropped by those.
---

# Office文書リーダー（AIがWord/Excelを読む）

## Purpose

Give the agent a **complete and faithful** view of a .docx/.xlsx file:
not just body text and cell values, but also the content that naive
extractors silently drop — text inside shapes, text boxes and callouts,
cell comments, embedded images (with their positions), charts, hidden
sheets and merged-cell structure.

This skill is about *reading*. It does not restructure or rewrite.
For turning the extracted material into polished Markdown, hand the
output to the **office-to-markdown** skill.

## Workflow

### 1. Extract

```bash
python3 <skill>/scripts/extract.py <file.docx|file.xlsx> -o <out_dir>
```

Outputs:

```
<out_dir>/
  extracted.md     faithful extraction, annotated with positions
  inventory.json   counts and warnings — read this first
  media/           extracted images
  work/shapes.json shape/text-box texts with anchor cells (xlsx)
```

- **.docx**: uses pandoc when installed (best fidelity), otherwise a
  stdlib fallback (headings, paragraphs, tables, images).
- **.xlsx**: uses openpyxl for cells plus direct parsing of the file's
  XML parts — `xl/drawings/*.xml` for shape/text-box text and anchors,
  `xl/comments*.xml` for comments, `xl/media/*` for pictures.
  MarkItDown is an acceptable cells-only supplement, never the sole
  source: it extracts neither images nor shape text.

### 2. Read the inventory before the text

`inventory.json` tells you what the document *contains*: sheets (and
whether they are hidden), merged ranges, shapes with text, comments,
media files, charts, and warnings. If a count is zero where the user
expects content, stop and investigate instead of answering from a
partial extraction.

### 3. Answer / hand off from the extraction

- Read `extracted.md` for body text, tables and cell values.
- Read `work/shapes.json` for shape and text-box text (anchor cell =
  which part of the sheet it annotates).
- Look at the images in `media/` when the question is about a figure,
  screenshot or diagram. An image's anchor cell tells you where it
  belongs; the surrounding cells tell you what it illustrates.
- For layout-dependent questions ("what is next to this figure?",
  reading order of a free-form Excel sheet), export the source to PDF
  (Excel or LibreOffice headless) and read the rendered pages as images.
  The PDF is the ground truth for layout; the extraction is the ground
  truth for exact text and numbers.

## What cannot be extracted (say so, do not guess)

- Excel charts and SmartArt are objects, not text or images. The
  inventory counts them; read them via a PDF/PNG rendering.
- Grouped or complex diagrams: extract their texts, but get the
  structure (arrows, grouping) from a rendering.
- In-cell Excel images (セル内画像) may not anchor to a sheet; they are
  extracted to `media/` and flagged in `warnings`.
- Formula cells yield cached values. When the formula itself matters,
  re-read with `openpyxl` (`data_only=False`) and report both.
- Headers/footers, tracked changes and colour-only meaning are not in
  the extraction unless explicitly requested (pandoc
  `--track-changes=all` for Word change history).

## Rules for answering from Office documents

- Cite positions: sheet name + cell range, shape anchor, or page of the
  PDF rendering — not just "the document says".
- Never claim a document has no figures/comments/diagrams without
  checking `inventory.json` counts.
- When extraction warnings exist, mention what might be missing in
  your answer instead of presenting partial content as complete.

## Details

See `references/extraction-notes.md` for the file-format map (where each
kind of content lives inside .docx/.xlsx) and troubleshooting.
