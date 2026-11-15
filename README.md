# office-spec-to-markdown

Word（.docx）・Excel（.xlsx）で作られた仕様書・設計書を、図表・図形・
テキストボックス・コメント・画像まで含めて Markdown（GFM）へ変換するための
AIエージェント向けスキルです。

An agent skill for converting Word/Excel specification documents — including
tables, shapes, text boxes, comments and images — into structured Markdown.

## Why not just "ask the AI to convert it"?

LLMs cannot reliably read Office binaries in one shot, and naive converters
silently drop exactly the parts that matter in specification documents:
text inside shapes and callouts, cell comments, charts, and merged-cell
structure. This skill splits the work into three verifiable steps:

1. **Extract mechanically** — `scripts/convert.py` pulls out paragraphs,
   cells, tables, shape text (from the drawing XML), comments and images,
   and records counts in `inventory.json`.
2. **Restructure with the LLM** — following `references/conversion-rules.md`
   (sheet-type classification for Excel, Mermaid vs. PNG decisions for
   diagrams, making colour/grey-out meaning explicit, …).
3. **Verify by counts** — `scripts/verify.py` reconciles the output with
   the inventory: no broken image links, no unreferenced media, no lost
   shape texts.

## Layout

```
office-spec-to-markdown/
  SKILL.md                          the skill itself
  scripts/
    convert.py                      .docx/.xlsx -> Markdown + inventory.json + media/
    verify.py                       link/count verification against inventory.json
  references/
    conversion-rules.md             detailed Word/Excel conversion rules
  templates/
    convert-spec.prompt.md          GitHub Copilot prompt file (.github/prompts/)
examples/
  make_samples.py                   generates sample .docx/.xlsx for testing
```

## Quick start

```bash
# Word (uses pandoc when installed, stdlib fallback otherwise)
python3 office-spec-to-markdown/scripts/convert.py spec.docx -o out/

# Excel (requires openpyxl + Pillow)
pip install openpyxl Pillow
python3 office-spec-to-markdown/scripts/convert.py spec.xlsx -o out/

# Verify
python3 office-spec-to-markdown/scripts/verify.py out/
```

## Use with your agent

- **Agent Skills compatible agents** (GitHub Copilot in VS Code, Claude
  Code, Codex, …): copy or clone the `office-spec-to-markdown/` folder into
  your skills directory, or reference it from your workspace.
- **GitHub Copilot prompt file**: copy `templates/convert-spec.prompt.md`
  to `.github/prompts/convert-spec.prompt.md` in your repository and run
  `/convert-spec` in Copilot Chat (agent mode).

## Test

```bash
python3 examples/make_samples.py /tmp/spec-samples
python3 office-spec-to-markdown/scripts/convert.py /tmp/spec-samples/sample.xlsx -o /tmp/out-xlsx
python3 office-spec-to-markdown/scripts/convert.py /tmp/spec-samples/sample.docx -o /tmp/out-docx
python3 office-spec-to-markdown/scripts/verify.py /tmp/out-xlsx
python3 office-spec-to-markdown/scripts/verify.py /tmp/out-docx
```

## License

MIT
