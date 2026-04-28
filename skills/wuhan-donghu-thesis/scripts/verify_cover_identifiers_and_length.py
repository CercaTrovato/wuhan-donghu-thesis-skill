# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}


def w_attr(el: etree._Element | None, name: str) -> str | None:
    if el is None:
        return None
    return el.get(f"{{{W_NS}}}{name}")


def paragraph_text(p: etree._Element) -> str:
    chunks: list[str] = []
    for r in p.xpath("./w:r", namespaces=NS):
        chunks.extend(r.xpath("./w:t/text()", namespaces=NS))
        if r.xpath("./w:tab", namespaces=NS):
            chunks.append("\t")
        if r.xpath("./w:br", namespaces=NS):
            chunks.append("\n")
    return "".join(chunks)


def run_text(r: etree._Element) -> str:
    chunks = r.xpath("./w:t/text()", namespaces=NS)
    if r.xpath("./w:tab", namespaces=NS):
        chunks.append("\t")
    return "".join(chunks)


def run_has_preserved_space(r: etree._Element) -> bool:
    for t in r.xpath("./w:t", namespaces=NS):
        if t.get(f"{{{XML_NS}}}space") == "preserve":
            return True
    return False


def run_is_underlined(r: etree._Element) -> bool:
    u = r.find("w:rPr/w:u", NS)
    if u is None:
        return False
    value = w_attr(u, "val")
    return value in (None, "single")


def is_right_aligned(p: etree._Element) -> bool:
    jc = p.find("w:pPr/w:jc", NS)
    return w_attr(jc, "val") == "right"


def tab_positions(p: etree._Element) -> set[str]:
    return {
        pos
        for pos in (
            w_attr(tab, "pos")
            for tab in p.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS)
        )
        if pos
    }


@dataclass
class IdentifierLine:
    para_index: int
    text: str
    paragraph: etree._Element

    @property
    def suffix(self) -> str:
        return self.text.split("：", 1)[1] if "：" in self.text else ""

    @property
    def underlined_suffix(self) -> str:
        pieces: list[str] = []
        seen_label = False
        for r in self.paragraph.xpath("./w:r", namespaces=NS):
            txt = run_text(r)
            if not seen_label:
                if "：" in txt:
                    seen_label = True
                    tail = txt.split("：", 1)[1]
                    if tail and run_is_underlined(r):
                        pieces.append(tail)
                continue
            if run_is_underlined(r):
                pieces.append(txt)
        return "".join(pieces)

    @property
    def underlined_runs_preserve_spaces(self) -> bool:
        for r in self.paragraph.xpath("./w:r", namespaces=NS):
            if run_is_underlined(r) and (" " in run_text(r)):
                if not run_has_preserved_space(r):
                    return False
        return True


def load_document(docx_path: Path) -> etree._Element:
    with ZipFile(docx_path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def find_identifier_line(paragraphs: list[etree._Element], label: str) -> IdentifierLine | None:
    for index, p in enumerate(paragraphs[:15], start=1):
        text = paragraph_text(p)
        if text.startswith(label):
            return IdentifierLine(index, text, p)
    return None


def count_nonspace_chars(root: etree._Element) -> int:
    text = "".join(root.xpath("//w:t/text()", namespaces=NS))
    return len(re.sub(r"\s+", "", text))


def check_identifier_line(
    failures: list[str],
    line: IdentifierLine | None,
    *,
    label: str,
    expected_value: str | None,
    min_suffix_chars: int,
    required_tabs: set[str],
) -> None:
    if line is None:
        failures.append(f"{label}: 未找到右上角字段行")
        return

    if not is_right_aligned(line.paragraph):
        failures.append(f"{label}: 段落必须右对齐")

    actual_tabs = tab_positions(line.paragraph)
    missing_tabs = sorted(required_tabs - actual_tabs)
    if missing_tabs:
        failures.append(f"{label}: 缺少范本制表位 {', '.join(missing_tabs)} twips")

    suffix = line.suffix
    underlined_suffix = line.underlined_suffix
    if expected_value and expected_value not in suffix:
        failures.append(f"{label}: 未包含指定值 {expected_value}")

    if len(underlined_suffix) < min_suffix_chars:
        failures.append(f"{label}: 冒号后的值或占位线必须使用下划线保留，当前下划线内容长度 {len(underlined_suffix)}")

    if not line.underlined_runs_preserve_spaces:
        failures.append(f"{label}: 下划线空格必须使用 xml:space='preserve'，否则 Word 中横线长度会丢失")


def run_checks(
    docx_path: Path,
    *,
    student_id: str | None,
    archive_no: str | None,
    min_nonspace_chars: int,
    skip_length_check: bool,
) -> tuple[list[str], int]:
    root = load_document(docx_path)
    paragraphs = root.xpath("//w:body/w:p", namespaces=NS)
    failures: list[str] = []

    student_line = find_identifier_line(paragraphs, "学号：")
    archive_line = find_identifier_line(paragraphs, "档号：")

    check_identifier_line(
        failures,
        student_line,
        label="学号",
        expected_value=student_id,
        min_suffix_chars=max(12, len(student_id or "")),
        required_tabs={"8280"},
    )
    check_identifier_line(
        failures,
        archive_line,
        label="档号",
        expected_value=archive_no,
        min_suffix_chars=max(12, len(archive_no or "")),
        required_tabs={"7920", "8280"},
    )

    if student_line and archive_line and archive_line.para_index != student_line.para_index + 1:
        failures.append("学号/档号: 档号行必须紧跟学号行，保持范本第 2、3 段结构")

    nonspace_chars = count_nonspace_chars(root)
    if not skip_length_check and nonspace_chars < min_nonspace_chars:
        failures.append(f"总字数: 非空白可见字符数 {nonspace_chars}，少于要求 {min_nonspace_chars}")

    return failures, nonspace_chars


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify cover identifier lines and thesis length.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--student-id", help="Expected student number on the Chinese cover.")
    parser.add_argument("--archive-no", help="Expected archive number. If omitted, the underlined placeholder still must remain.")
    parser.add_argument("--min-nonspace-chars", type=int, default=20000, help="Minimum visible non-space character count. Default: 20000.")
    parser.add_argument("--skip-length-check", action="store_true", help="Only check the cover identifier lines.")
    args = parser.parse_args()

    failures, nonspace_chars = run_checks(
        args.docx.resolve(),
        student_id=args.student_id,
        archive_no=args.archive_no,
        min_nonspace_chars=args.min_nonspace_chars,
        skip_length_check=args.skip_length_check,
    )

    print(f"NONSPACE_VISIBLE_CHARS={nonspace_chars}")
    if failures:
        print("COVER_IDENTIFIERS_AND_LENGTH_CHECK=FAIL")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("COVER_IDENTIFIERS_AND_LENGTH_CHECK=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
