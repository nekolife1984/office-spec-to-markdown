---
name: office-spec-to-markdown
description: Convert Word (.docx) and Excel (.xlsx) specification documents that contain text, tables, shapes, text boxes, comments and images into structured Markdown (GitHub Flavored Markdown). Use when the user wants to migrate a Word/Excel 仕様書・設計書 to Markdown, make it Git-managed or AI-readable, or asks how to convert docx/xlsx including figures and diagrams. Works with GitHub Copilot (VS Code agent mode / Copilot CLI), Claude, Codex and other agents that can run local commands.
---

# Office仕様書 → Markdown変換

## Core principle

Do NOT ask the LLM to "read the .docx/.xlsx and rewrite it as Markdown" in one
shot. Split the job in three:

1. **Extract mechanically** (scripts in this skill) — every paragraph, cell,
   table, shape text, comment and image, with counts.
2. **Restructure with the LLM** — apply the conversion rules in
   `references/conversion-rules.md` to the extracted material.
3. **Verify by counts** — compare the result against `inventory.json` with
   `scripts/verify.py`. A conversion that cannot show its counts is not done.

## Workflow

### 1. Inventory first

Run the converter on the source file. It always writes `inventory.json`:

```bash
python3 <skill>/scripts/convert.py <file.docx|file.xlsx> -o <out_dir>
```

Read `inventory.json` before touching the Markdown. Note especially:
sheets and their state (hidden sheets!), merged ranges, shapes with text,
comments, media files, charts, and the `warnings` array.

### 2. Convert

- **.docx** — the script uses pandoc when installed
  (`pandoc <file> -t gfm --wrap=none --extract-media=<out_dir>`), otherwise a
  stdlib fallback extractor (headings, paragraphs, tables, images only).
  Recommend pandoc for production use.
- **.xlsx** — the script uses openpyxl plus direct XML parsing of the file:
  cells become per-sheet Markdown tables, shape/text-box text is extracted
  from `xl/drawings/*.xml` with its anchor cell, comments from
  `xl/comments*.xml`, and embedded pictures are extracted to `media/` with
  their anchor cell. MarkItDown (`markitdown <file.xlsx>`) is an acceptable
  first pass for cells only — it does not extract images or shape text.

### 3. Handle what Markdown cannot express

Decide per item, following `references/conversion-rules.md`:

- Flowcharts, sequence/state diagrams drawn with shapes → redraw as Mermaid
  AND keep a PNG rendering alongside.
- Excel charts, SmartArt, complex grouped diagrams → render the source to
  PDF (Excel or LibreOffice headless), export PNG, reference it from the
  Markdown, and add a text description of what it shows.
- Merged-cell or multi-line tables that break GFM pipe tables → keep an HTML
  `<table>` or split the table; never silently drop rows.
- Meaning carried by colour, grey-out or strikethrough → make it explicit
  (`<del>`, notes, a status column).

### 4. Restructure (the LLM's job)

- Rebuild heading hierarchy from the document's own structure (Word styles /
  Excel sheet types), not from visual size.
- Absorb shape and comment text into the section it belongs to (use the
  anchor cell as the hint). Remove shape-name headings such as
  "正方形/長方形 15".
- For Excel, classify each sheet first (screen spec / item definition /
  data table / figure-only / prose) and give each type its own structure.
  Figure-only sheets may stay as PDF/PNG references instead of forced Markdown.
- Keep figure numbering and cross-references ("図1", "表2") consistent with
  the extracted media files.

### 5. Verify

```bash
python3 <skill>/scripts/verify.py <out_dir>
```

Then check by counts: headings vs. source chapters, table rows, shape texts
extracted vs. shapes in inventory, media files all referenced, no broken
links. Report any item deliberately left as image/PDF to the user.

## Using with GitHub Copilot

- In VS Code, put the source file in the workspace and run this skill in
  agent mode, or use the prompt template in
  `templates/convert-spec.prompt.md` (save it as
  `.github/prompts/convert-spec.prompt.md` in the target repository).
- Copilot Chat cannot reliably read .docx/.xlsx binaries directly. Always
  run the extraction script first and let Copilot work on the extracted
  Markdown/JSON. For layout questions, also export the source to PDF and
  attach it — Copilot can read PDFs and images.

## Output layout

```
<out_dir>/
  <stem>.md        final Markdown (after restructuring)
  inventory.json   extraction counts and warnings
  media/           extracted images
  work/            intermediates (shapes.json etc.) — keep for audit, or delete after verification
```

## Known limits

- Legacy .doc / .xls are not supported; re-save as .docx / .xlsx first.
- In-cell Excel images (セル内画像) may not anchor to a sheet; they are still
  extracted to `media/` and flagged in `warnings`.
- Formula cells are converted with their cached values, not the formulas.
- Headers/footers, font colours and page layout have no Markdown equivalent
  and are dropped — say so in the delivery report.
