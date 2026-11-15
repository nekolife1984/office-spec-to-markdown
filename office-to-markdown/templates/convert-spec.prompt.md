---
mode: agent
description: Convert a Word/Excel specification document in this workspace to Markdown using the office-docs-reader and office-to-markdown skills
---

Convert the Office specification document I specify (a .docx or .xlsx file in
this workspace) into Markdown, using the two skills in sequence:

1. Read it with the office-docs-reader skill first and show me the inventory
   counts before restructuring anything:
   `python3 <reader-skill-path>/scripts/extract.py <file> -o <output-dir>`
2. Restructure the extracted material with the office-to-markdown skill,
   according to its `references/conversion-rules.md`. In particular:
   - absorb shape, text-box and comment text into the sections they belong to
   - classify each Excel sheet by type before converting it
   - redraw simple diagrams as Mermaid and keep complex figures as PNG
     references with a text description
   - never drop table rows, shape texts or comments silently
3. Run `python3 <converter-skill-path>/scripts/verify.py <output-dir>` and reconcile
   the counts with inventory.json.
4. Report: what was converted, what stayed as image/PDF and why, what was
   dropped (layout, colours, headers/footers), and any judgement calls.

Do not read the .docx/.xlsx binary directly as the source of truth, and do
not commit the result until I confirm it.
