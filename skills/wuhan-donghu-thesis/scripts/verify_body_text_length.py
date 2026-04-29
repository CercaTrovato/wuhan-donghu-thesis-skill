#!/usr/bin/env python3
"""Verify thesis body text length.

Counts visible non-space characters from the first body chapter heading
(`1 ...`) to the heading before references. This is stricter than whole-document
length checks because cover pages, abstracts, TOC, references, and acknowledgments
do not satisfy the school requirement for body text length.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_block_items(document: Document):
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def is_body_heading(paragraph: Paragraph, text: str) -> bool:
    style_name = paragraph.style.name if paragraph.style is not None else ""
    return style_name in {"Thesis Heading 1", "Heading 1"} and re.match(r"^[1-9]\d*\s+", text)


def is_references_heading(paragraph: Paragraph, text: str) -> bool:
    style_name = paragraph.style.name if paragraph.style is not None else ""
    return style_name in {"Thesis Heading 1", "Heading 1"} and text.startswith("参考文献")


def table_text(table: Table) -> str:
    # python-docx repeats merged-cell text; de-duplicate within each table for a
    # stable visible-text approximation.
    parts: list[str] = []
    seen: set[str] = set()
    for row in table.rows:
        for cell in row.cells:
            text = "\n".join(p.text for p in cell.paragraphs).strip()
            if text and text not in seen:
                seen.add(text)
                parts.append(text)
    return "\n".join(parts)


def count_body_text(docx_path: Path) -> tuple[int, list[tuple[str, int]]]:
    document = Document(str(docx_path))
    in_body = False
    parts: list[str] = []
    sections: list[tuple[str, int]] = []
    current_index: int | None = None

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not in_body and is_body_heading(block, text):
                in_body = True
            if in_body and is_references_heading(block, text):
                break
            if not in_body:
                continue
            parts.append(text)
            if is_body_heading(block, text):
                sections.append((text, 0))
                current_index = len(sections) - 1
            if current_index is not None:
                cleaned = re.sub(r"\s+", "", text)
                title, count = sections[current_index]
                sections[current_index] = (title, count + len(cleaned))
        else:
            if not in_body:
                continue
            text = table_text(block)
            parts.append(text)
            if current_index is not None:
                cleaned = re.sub(r"\s+", "", text)
                title, count = sections[current_index]
                sections[current_index] = (title, count + len(cleaned))

    total = len(re.sub(r"\s+", "", "\n".join(parts)))
    return total, sections


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify body text non-space visible character count.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--min", type=int, default=20000, dest="minimum")
    args = parser.parse_args()

    total, sections = count_body_text(args.docx)
    print(f"BODY_NONSPACE_VISIBLE_CHARS={total}")
    for title, count in sections:
        print(f"BODY_SECTION_CHARS {title}={count}")
    if total < args.minimum:
        print("BODY_TEXT_LENGTH_CHECK=FAIL")
        print(f"REQUIRED_MIN={args.minimum}")
        return 1
    print("BODY_TEXT_LENGTH_CHECK=PASS")
    print(f"REQUIRED_MIN={args.minimum}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
