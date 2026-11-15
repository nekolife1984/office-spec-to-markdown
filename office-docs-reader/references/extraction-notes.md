# Extraction notes — where content lives inside .docx / .xlsx

Both formats are ZIP archives of XML parts. When a library does not
surface some content, read the part directly instead of concluding the
content is absent.

## Word (.docx)

| Content | Location |
|---|---|
| Body, headings, tables | `word/document.xml` (styles in `word/styles.xml`) |
| Images | `word/media/*`, referenced by `a:blip/@r:embed` → `word/_rels/document.xml.rels` |
| Footnotes / endnotes | `word/footnotes.xml`, `word/endnotes.xml` |
| Comments | `word/comments.xml` |
| Headers / footers | `word/header*.xml`, `word/footer*.xml` |
| Text boxes / shapes | drawing elements inside `document.xml` (`w:txbxContent`) — pandoc drops most of these |
| SmartArt | `word/diagrams/*` plus a fallback drawing in `document.xml` |

## Excel (.xlsx)

| Content | Location |
|---|---|
| Cell values, merged ranges | `xl/worksheets/sheetN.xml` (+ `xl/sharedStrings.xml`) |
| Sheet order, names, visibility | `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` |
| Per-sheet drawings/comments links | `xl/worksheets/_rels/sheetN.xml.rels` |
| Shapes, text boxes, callouts, pictures (anchors + text) | `xl/drawings/drawingN.xml` — shapes are `xdr:sp` (also nested in `grpSp`), text in `a:t`, position in the anchor's `xdr:from` |
| Cell comments | `xl/comments/commentN.xml` (note the singular folder name varies: `xl/comments*`) |
| Pictures (files) | `xl/media/*` |
| Charts | `xl/charts/chartN.xml` — objects, not images; render to read |
| In-cell images | cell `vm` attribute → `xl/metadata.xml` → `xl/richData/*` → `xl/media/*` (openpyxl does not load these) |

## Troubleshooting

- **Images extracted but not anchored**: match by content. openpyxl's
  `ws._images` gives anchors for floating pictures; hash the image bytes
  (`img._data()`) and match them against `xl/media/*` files, whose names
  alone say nothing about position.
- **A sheet seems missing**: check `sheet_state` in the inventory —
  hidden and very-hidden sheets are real content until the user says
  they are scratch.
- **Numbers look wrong**: date/time cells are serial numbers in the
  XML; libraries like openpyxl convert them via the cell style. Formula
  cells show the cached value from the last Excel save, which can be
  stale if the file was produced by another tool.
- **Japanese text garbled**: the XML is UTF-8; garbling usually comes
  from reading shared strings with the wrong encoding in hand-rolled
  parsers. Prefer openpyxl/pandoc over ad-hoc string handling.
- **Huge sheets**: free-form spec sheets can have a large used range
  full of empty formatted cells. Trim trailing all-empty rows/columns
  (extract.py does) before feeding text to an LLM.
