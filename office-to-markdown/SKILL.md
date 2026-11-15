---
name: office-to-markdown
description: Convert Word (.docx) and Excel (.xlsx) documents — especially specification documents (仕様書・設計書) with tables, shapes, text boxes, comments and images — into structured, Git-friendly Markdown (GFM). Use when the user wants a .docx/.xlsx turned into a .md file, migrated to docs-as-code / Git management, or restructured for readers and AI. This skill consumes the extraction produced by the office-docs-reader skill (extracted.md + inventory.json + media/) and restructures it under explicit conversion rules, then verifies the result by counts. Do not use it to merely read or summarize an Office document — that is office-docs-reader's job.
---

# Office文書 → Markdown変換

## Purpose

Turn an Office document into Markdown that a human would be happy to
maintain in Git: real heading hierarchy, clean tables, figures as
referenced files, diagrams as Mermaid where sensible.

Conversion is **restructuring**, not extraction. Extraction is the
**office-docs-reader** skill's job; this skill starts from its output.

## Inputs

An extraction directory produced by office-docs-reader:

```
<dir>/
  extracted.md     faithful extraction with positions
  inventory.json   counts and warnings (the contract)
  media/           extracted images
  work/shapes.json shape/text-box texts with anchors (xlsx)
```

If you only have the source .docx/.xlsx, run office-docs-reader first.
Never restructure directly from the binary.

## Workflow

### 1. Read the contract

Read `inventory.json` and note: sections/sheets, tables, shapes with
text, comments, media files, charts, warnings. These counts are what
the final Markdown must reconcile with.

### 2. Plan the structure before writing

- Word: heading hierarchy from the document's styles/sections.
- Excel: classify each sheet first (data table / screen spec / prose /
  figure-only / cover & history) — each type gets a different structure.
  See `references/conversion-rules.md`.
- Decide per diagram: Mermaid (simple flow/sequence/state/ER) or PNG
  reference + text description (screenshots, charts, dense diagrams).
  Figure-only Excel sheets may stay PDF/PNG references entirely.

### 3. Restructure

- Absorb shape, text-box and comment text into the section it belongs
  to (anchor position is the hint). Remove auto shape names
  ("正方形/長方形 15") — keep only the meaning.
- Tables that GFM cannot express (merged cells, multi-paragraph cells):
  use an HTML `<table>`, split the table, or use a definition list plus
  a simple table. Never drop rows or columns.
- Make implicit meaning explicit: colour, grey-out, strikethrough or
  fill that carries meaning becomes text (`<del>`, a status column, a
  note). Decoration is dropped, once, and mentioned in the report.
- Keep figure/table numbering and cross-references consistent with the
  files in `media/`.
- Write the final file as `<stem>.md` in the extraction directory (next
  to `extracted.md`, which stays as the audit trail).

### 4. Verify

```bash
python3 <skill>/scripts/verify.py <dir> [<stem>.md]
```

Reconcile: headings vs. source chapters, table rows/columns, shape
texts placed = shapes in inventory, comments all placed or explicitly
dropped with a reason, every media file referenced, no broken links,
charts/SmartArt present as PNG + description.

### 5. Report

Tell the user: what was converted, what stayed as image/PDF and why,
what was dropped (page layout, colours, headers/footers), and every
judgement call. A conversion without this report is not done.

## Using with GitHub Copilot

Copy `templates/convert-spec.prompt.md` to
`.github/prompts/convert-spec.prompt.md` in the target repository and
run `/convert-spec` in agent mode. The prompt drives both skills in
sequence (reader extraction → this skill's restructuring → verify).

## Rules reference

All detailed rules — Word specifics, Excel sheet-type handling, merged
cells, formulas, hidden sheets, the full verification checklist — live
in `references/conversion-rules.md`. Read it before restructuring.
