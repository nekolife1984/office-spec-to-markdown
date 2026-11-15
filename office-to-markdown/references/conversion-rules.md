# Conversion rules — Word / Excel仕様書 → Markdown

Detailed rules for the office-to-markdown skill. Apply them during the
restructuring step, working from the extraction produced by the
office-docs-reader skill (`extracted.md`, `inventory.json`, `media/`,
`work/shapes.json`). When a rule forces a judgement call, note it in the
delivery report instead of deciding silently.

## Common rules (Word & Excel)

1. **Extract first, rewrite second.** Never restructure from the binary
   directly. Work from office-docs-reader's extraction: `extracted.md`,
   `shapes.json` and `inventory.json`.
2. **Counts are the contract.** Headings, tables, shape texts, comments and
   media files in the output must reconcile with `inventory.json`.
3. **Figures: drawio.svg by default, PNG when difficult.**
   - Diagrams made of boxes, arrows and text (flowcharts, screen
     transitions, state diagrams, configuration/architecture diagrams,
     org charts) → rebuild as **`.drawio.svg`**: an SVG that renders in
     Markdown and embeds the diagrams.net (draw.io) graph, so the figure
     stays editable after conversion. For simple node/edge diagrams use
     `scripts/make_drawio_svg.py` (JSON in, .drawio.svg out). For richer
     diagrams, author the .drawio source and export SVG from
     diagrams.net with the diagram embedded.
   - Validate every rebuilt figure: the embedded graph must parse, and
     its node/edge counts and labels must match the source figure.
     Keep the source rendering (PNG) in the work folder until the
     rebuild is verified.
   - **Difficult figures stay PNG**: photos, screenshots, dense or
     highly detailed diagrams, data charts (graphs), handwritten or
     scanned figures — anything where redrawing risks losing or
     inventing information. PNG reference plus a short text description
     (what it shows, key values, conclusion).
   - A figure that is partly simple: rebuild the structure as
     drawio.svg and note in the description what was simplified.
4. **Tables that GFM cannot express** (merged cells, multi-paragraph cells,
   nested tables): keep an HTML `<table>`, split into several simple
   tables, or turn the header part into a definition list plus a simple
   table. Dropping rows or columns is never acceptable.
5. **Implicit meaning must become explicit.** Colour, grey-out, highlight,
   strikethrough, cell fill — if it carries meaning in the source, express
   it in text (`<del>`, a status column, a note). If it is decoration only,
   drop it and say so once in the report.
6. **Cross-references.** Keep "図1 / 表2 / 項番" numbering stable. After
   restructuring, check that every reference resolves to an existing
   heading, table or image.

## Word (.docx) rules

- Heading levels come from paragraph styles (見出し1/Heading 1 …), not from
  font size or bold.
- Use pandoc with `-t gfm --wrap=none --extract-media=<dir>` as the primary
  route. `--wrap=none` keeps one paragraph per line so diffs stay readable.
- SmartArt, text boxes, floating shapes and Excel charts embedded in Word
  are not converted by pandoc. Detect them (they appear as images or not at
  all), then: export the document to PDF, cut the region out as PNG, and
  either rebuild it as drawio.svg (boxes/arrows/text figures) or keep
  it as PNG with a text description (difficult figures).
- Tracked changes: convert the accepted state by default. Only use
  `--track-changes=all` when the user explicitly wants the change history.
- Headers, footers, page numbers, TOC fields: drop. Regenerate a TOC from
  the Markdown headings if needed.
- Footnotes survive pandoc conversion as `[^n]`; check they are not
  duplicated in the body text.

## Excel (.xlsx) rules

Excel specification sheets are free-form layouts on a cell grid, so classify
each sheet before converting:

| Sheet type | How to convert |
|---|---|
| Data table (項目定義・一覧) | One Markdown table per logical table. Fill merged header cells down/right so every row knows its group. |
| Screen / form spec (画面仕様) | Rebuild as headings + item tables. The screenshot stays as an image; item numbers in the image must match the table rows. |
| Prose (概要・説明文) | Merge the scattered cells into paragraphs in reading order (top→bottom, left→right). Do not ship a giant sparse table. |
| Figure-only (フロー図・構成図) | Rebuild the figure as drawio.svg and reference it from a short Markdown section with a text walkthrough. If the figure is too dense or detailed to rebuild faithfully, reference a PNG rendering instead and say why. |
| Cover / history (表紙・改訂履歴) | Cover → title + metadata list. History → a simple table. |

Mechanics:

- **Merged cells**: in tables, propagate the merged value to all covered
  cells when it expresses grouping; keep only the top-left value when it is
  a title spanning columns (then make it a heading instead).
- **Shape / text-box / callout text** often holds real specification text.
  Extract all of it (office-docs-reader does this from `xl/drawings/*.xml`),
  sort by anchor cell, and absorb each piece into the section it annotates.
  Never keep auto shape names ("正方形/長方形 15") as headings.
- **Cell comments** are specification notes: extract and merge them into
  the relevant item description, marked as notes.
- **Hidden sheets, rows and columns**: list them in the report. Convert
  hidden sheets only when the user confirms they are not scratch work.
- **Formulas**: output cached values. If the formula itself is the
  specification (calculation rules), extract formulas separately
  (`openpyxl` with `data_only=False`) and present value + formula.
- **Charts** are objects, not images. Render via Excel/LibreOffice to PNG.
  Also extract the underlying data range as a table when the numbers matter.
- **In-cell images** (セル内画像, newer Excel) do not anchor like floating
  pictures. They are extracted to `media/`; place them by matching the
  cell's rich-value metadata or by visual check against a PDF rendering.
- **Reading order** is ambiguous in free-form sheets. When in doubt, export
  the sheet to PDF and use the rendered layout as the ground truth for
  order and for which shape belongs to which table.

## Verification checklist

- [ ] Sheet/section count matches the source (including a decision on hidden sheets)
- [ ] Every table's row and column counts reconciled
- [ ] Shape texts in output = shapes with text in `inventory.json`
- [ ] Comments all placed or explicitly dropped with a reason
- [ ] Media files in `media/` all referenced; no broken image links (`verify.py`)
- [ ] Figures: rebuilt as drawio.svg (embedded graph parses, node/edge counts and labels match the source) or, for difficult ones, PNG + description present
- [ ] Cross-references (図・表・項番) all resolve
- [ ] Delivery report lists: what was dropped (layout, colours, headers),
      what stayed as image/PDF, and every judgement call
