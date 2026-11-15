#!/usr/bin/env python3
"""Reformat Japanese Markdown to one sentence per line.

Joins hard-wrapped paragraph/list lines, then splits at sentence ends
(。？！). Code blocks, tables, headings, YAML frontmatter are untouched.
"""
import re
import sys
from pathlib import Path

LIST_RE = re.compile(r"^(\s*)([-*+]|\d+\.)(\s+)(.*)$")
HEAD_RE = re.compile(r"^#{1,6}\s")
CLOSERS = "」）』】〕"


def split_sentences(text):
    parts, buf, i = [], "", 0
    in_bold = False
    in_code = False
    while i < len(text):
        if text.startswith("**", i):
            buf += "**"
            if not in_code:
                in_bold = not in_bold
            i += 2
            continue
        ch = text[i]
        if ch == "`":
            in_code = not in_code
        buf += ch
        if ch in "。？！" and not in_bold and not in_code:
            j = i + 1
            while j < len(text) and text[j] in CLOSERS:
                buf += text[j]
                j += 1
            parts.append(buf)
            buf = ""
            i = j
            continue
        i += 1
    if buf:
        parts.append(buf)
    return [p.strip() for p in parts if p.strip()]


def join_lines(lines):
    text = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not text:
            text = ln
        elif (text[-1].isascii() and text[-1].isalnum()
              and ln[0].isascii() and ln[0].isalnum()):
            text += " " + ln
        else:
            text += ln
    return text


def emit_block(out, block_lines):
    """block_lines: list of raw lines forming one paragraph or one list item group."""
    # Split into list items and paragraphs preserving structure
    i = 0
    while i < len(block_lines):
        m = LIST_RE.match(block_lines[i])
        if m:
            indent, marker, sp, first = m.groups()
            content_indent = " " * (len(indent) + len(marker) + len(sp))
            item_lines = [first]
            i += 1
            while i < len(block_lines) and not LIST_RE.match(block_lines[i]):
                item_lines.append(block_lines[i])
                i += 1
            text = join_lines(item_lines)
            sentences = split_sentences(text)
            if not sentences:
                out.append(f"{indent}{marker} ")
                continue
            out.append(f"{indent}{marker} {sentences[0]}")
            for s in sentences[1:]:
                out.append(f"{content_indent}{s}")
        else:
            para = []
            while i < len(block_lines) and not LIST_RE.match(block_lines[i]):
                para.append(block_lines[i])
                i += 1
            for s in split_sentences(join_lines(para)):
                out.append(s)


def reformat(text):
    lines = text.split("\n")
    out = []
    i = 0
    # YAML frontmatter
    if lines and lines[0].strip() == "---":
        j = 1
        while j < len(lines) and lines[j].strip() != "---":
            j += 1
        out.extend(lines[:j + 1])
        i = j + 1
    in_code = False
    block = []

    def flush():
        nonlocal block
        if block:
            emit_block(out, block)
            block = []

    while i < len(lines):
        ln = lines[i]
        if ln.lstrip().startswith("```"):
            flush()
            in_code = not in_code
            out.append(ln)
            i += 1
            continue
        if in_code:
            out.append(ln)
            i += 1
            continue
        if not ln.strip():
            flush()
            out.append("")
            i += 1
            continue
        if HEAD_RE.match(ln) or ln.lstrip().startswith("|") or ln.strip() == "---":
            flush()
            out.append(ln)
            i += 1
            continue
        block.append(ln)
        i += 1
    flush()
    return "\n".join(out)


def normalized(s):
    return re.sub(r"\s+", "", s)


def main():
    for name in sys.argv[1:]:
        p = Path(name)
        orig = p.read_text(encoding="utf-8")
        new = reformat(orig)
        if normalized(orig) != normalized(new):
            print(f"MISMATCH (content changed): {name}")
            # locate first difference
            a, b = normalized(orig), normalized(new)
            for k in range(min(len(a), len(b))):
                if a[k] != b[k]:
                    print("  orig:", a[max(0, k - 30):k + 30])
                    print("  new :", b[max(0, k - 30):k + 30])
                    break
            print("  len:", len(a), "->", len(b))
            continue
        p.write_text(new, encoding="utf-8")
        print(f"OK: {name}")


if __name__ == "__main__":
    main()
