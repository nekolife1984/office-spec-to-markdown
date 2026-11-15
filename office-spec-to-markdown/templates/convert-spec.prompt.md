---
mode: agent
description: Convert a Word/Excel specification document in this workspace to Markdown using the office-spec-to-markdown skill
---

Convert the Office specification document I specify (a .docx or .xlsx file in
this workspace) into Markdown, following the office-spec-to-markdown skill:

1. Run the skill's extraction script first and show me the inventory counts
   before restructuring anything:
   `python3 <skill-path>/scripts/convert.py <file> -o <output-dir>`
2. Restructure the extracted material according to the skill's
   `references/conversion-rules.md`. In particular:
   - absorb shape, text-box and comment text into the sections they belong to
   - classify each Excel sheet by type before converting it
   - redraw simple diagrams as Mermaid and keep complex figures as PNG
     references with a text description
   - never drop table rows, shape texts or comments silently
3. Run `python3 <skill-path>/scripts/verify.py <output-dir>` and reconcile
   the counts with inventory.json.
4. Report: what was converted, what stayed as image/PDF and why, what was
   dropped (layout, colours, headers/footers), and any judgement calls.

Do not read the .docx/.xlsx binary directly as the source of truth, and do
not commit the result until I confirm it.
