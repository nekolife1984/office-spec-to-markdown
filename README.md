# Office skills: read Word/Excel, then convert to Markdown

Word（.docx）・Excel（.xlsx）をAIが読むためのスキルと、仕様書・設計書を
Markdown（GFM）へ変換するためのスキル、2本セットです。

Two agent skills that separate *reading* Office documents from *converting*
them:

| Skill | Job | Use it when |
|---|---|---|
| **office-docs-reader** | Extract everything an AI needs to read a .docx/.xlsx: body text, tables, cells, shape/text-box text, comments, images with positions, counts | An AI must read, search, summarize or answer questions about a Word/Excel file |
| **office-to-markdown** | Restructure the reader's extraction into maintainable Markdown under explicit rules, verified by counts | A Word/Excel document (especially a 仕様書・設計書) must become a .md file for Git / docs-as-code |

## Why split reading and converting?

Naive one-shot conversion silently drops exactly the parts that matter in
specification documents: text inside shapes and callouts, cell comments,
charts, merged-cell structure. Reading is also useful on its own (Q&A,
search, summarization) without any conversion. So:

1. **office-docs-reader** extracts mechanically and faithfully —
   `scripts/extract.py` produces `extracted.md`, `inventory.json`,
   `media/` and `work/shapes.json`. Nothing is restructured, nothing is
   silently dropped; un-extractable items (charts, SmartArt) are counted
   and flagged.
2. **office-to-markdown** consumes that extraction and restructures it
   under `references/conversion-rules.md` (Excel sheet-type
   classification, Mermaid vs. PNG decisions, making colour/grey-out
   meaning explicit, …), then reconciles the result with the inventory
   via `scripts/verify.py`.

## Layout

```
office-docs-reader/
  SKILL.md
  scripts/extract.py                .docx/.xlsx -> extracted.md + inventory.json + media/
  references/extraction-notes.md    where each content type lives in the file formats
office-to-markdown/
  SKILL.md
  scripts/verify.py                 link/count verification against inventory.json
  references/conversion-rules.md    detailed Word/Excel conversion rules
  templates/convert-spec.prompt.md  GitHub Copilot prompt file (.github/prompts/)
examples/
  make_samples.py                   generates sample .docx/.xlsx for testing
```

## Quick start

```bash
# 1. Read / extract (docx: pandoc when installed, stdlib fallback otherwise;
#    xlsx: pip install openpyxl Pillow)
python3 office-docs-reader/scripts/extract.py spec.xlsx -o out/
# -> out/extracted.md, out/inventory.json, out/media/, out/work/shapes.json

# 2. Convert: restructure out/extracted.md into out/spec.md following
#    office-to-markdown/references/conversion-rules.md (this is the agent's job)

# 3. Verify the result against the inventory
python3 office-to-markdown/scripts/verify.py out/ spec.md
```

## Use with your agent

- **Agent Skills compatible agents** (GitHub Copilot in VS Code, Claude
  Code, Codex, …): copy or clone both skill folders into your skills
  directory, or reference them from your workspace. They also work
  independently — install only the reader if you never convert.
- **GitHub Copilot prompt file**: copy
  `office-to-markdown/templates/convert-spec.prompt.md` to
  `.github/prompts/convert-spec.prompt.md` in your repository and run
  `/convert-spec` in Copilot Chat (agent mode). It drives both skills in
  sequence.

## Test

```bash
python3 examples/make_samples.py /tmp/spec-samples
python3 office-docs-reader/scripts/extract.py /tmp/spec-samples/sample.xlsx -o /tmp/out-xlsx
python3 office-docs-reader/scripts/extract.py /tmp/spec-samples/sample.docx -o /tmp/out-docx
# after restructuring extracted.md into a final .md in the same folder:
python3 office-to-markdown/scripts/verify.py /tmp/out-xlsx
```

## License

MIT
