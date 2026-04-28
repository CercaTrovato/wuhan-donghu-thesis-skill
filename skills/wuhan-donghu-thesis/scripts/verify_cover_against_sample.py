# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
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


def paragraph_alignment(p: etree._Element) -> str | None:
    return w_attr(p.find("w:pPr/w:jc", NS), "val")


def tab_positions(p: etree._Element) -> list[tuple[str | None, str | None]]:
    return [
        (w_attr(tab, "val"), w_attr(tab, "pos"))
        for tab in p.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS)
    ]


def run_is_underlined(r: etree._Element) -> bool:
    u = r.find("w:rPr/w:u", NS)
    if u is None:
        return False
    return w_attr(u, "val") in (None, "single")


def run_has_preserved_space(r: etree._Element) -> bool:
    return any(
        t.get(f"{{{XML_NS}}}space") == "preserve"
        for t in r.xpath("./w:t", namespaces=NS)
    )


def run_size(r: etree._Element) -> str | None:
    return w_attr(r.find("w:rPr/w:sz", NS), "val")


def run_east_asia_font(r: etree._Element) -> str | None:
    return w_attr(r.find("w:rPr/w:rFonts", NS), "eastAsia")


def suffix_run_texts(p: etree._Element, label: str) -> list[tuple[etree._Element, str]]:
    full_text = paragraph_text(p)
    label_start = full_text.find(label)
    if label_start < 0:
        return []
    label_end = label_start + len(label)

    runs: list[tuple[etree._Element, str]] = []
    position = 0
    for r in p.xpath("./w:r", namespaces=NS):
        txt = run_text(r)
        next_position = position + len(txt)
        if next_position <= label_end:
            position = next_position
            continue
        if position >= label_end:
            runs.append((r, txt))
        else:
            runs.append((r, txt[label_end - position :]))
        position = next_position
    return runs


def underlined_suffix_text(p: etree._Element, label: str) -> str:
    return "".join(txt for r, txt in suffix_run_texts(p, label) if run_is_underlined(r))


def display_width(text: str) -> int:
    width = 0
    for ch in text:
        if ch in "\t\r\n":
            continue
        if ch == " " or ord(ch) < 128:
            width += 1
        else:
            width += 2
    return width


def load_paragraphs(path: Path) -> list[etree._Element]:
    with ZipFile(path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    return root.xpath("//w:body/w:p", namespaces=NS)


@dataclass
class LineSignature:
    index: int
    label: str
    text: str
    alignment: str | None
    tabs: list[tuple[str | None, str | None]]
    label_font: str | None
    label_size: str | None
    suffix_size: str | None
    suffix_min_len: int
    paragraph: etree._Element


def find_line(paragraphs: list[etree._Element], label: str) -> LineSignature | None:
    for index, p in enumerate(paragraphs[:25], start=1):
        text = paragraph_text(p)
        if not text.startswith(label):
            continue

        runs = p.xpath("./w:r", namespaces=NS)
        label_run = next((r for r in runs if label in run_text(r)), None)
        underlined = underlined_suffix_text(p, label)
        first_suffix_run = next((r for r, txt in suffix_run_texts(p, label) if txt), None)

        return LineSignature(
            index=index,
            label=label,
            text=text,
            alignment=paragraph_alignment(p),
            tabs=tab_positions(p),
            label_font=run_east_asia_font(label_run) if label_run is not None else None,
            label_size=run_size(label_run) if label_run is not None else None,
            suffix_size=run_size(first_suffix_run) if first_suffix_run is not None else None,
            suffix_min_len=max(12, len(underlined.strip())),
            paragraph=p,
        )
    return None


def compare_line(
    failures: list[str],
    sample: LineSignature | None,
    candidate: LineSignature | None,
    *,
    expected_value: str | None,
) -> None:
    label = sample.label if sample else (candidate.label if candidate else "字段")
    if sample is None:
        failures.append(f"{label}: 范本中未找到字段行，无法建立对比基准")
        return
    if candidate is None:
        failures.append(f"{label}: 候选文档中未找到字段行")
        return

    if candidate.index != sample.index:
        failures.append(f"{label}: 段落序号 {candidate.index} 与范本 {sample.index} 不一致")
    if candidate.alignment != sample.alignment:
        failures.append(f"{label}: 段落对齐 {candidate.alignment} 与范本 {sample.alignment} 不一致")
    if candidate.tabs != sample.tabs:
        failures.append(f"{label}: 制表位 {candidate.tabs} 与范本 {sample.tabs} 不一致")
    if candidate.label_font != sample.label_font:
        failures.append(f"{label}: 标签字体 {candidate.label_font} 与范本 {sample.label_font} 不一致")
    if candidate.label_size != sample.label_size:
        failures.append(f"{label}: 标签字号 {candidate.label_size} 与范本 {sample.label_size} 不一致")
    if candidate.suffix_size != sample.suffix_size:
        failures.append(f"{label}: 冒号后字号 {candidate.suffix_size} 与范本 {sample.suffix_size} 不一致")

    if expected_value and expected_value not in candidate.text:
        failures.append(f"{label}: 未包含指定值 {expected_value}")

    suffix = underlined_suffix_text(candidate.paragraph, label)
    if len(suffix) < sample.suffix_min_len:
        failures.append(
            f"{label}: 下划线内容长度 {len(suffix)} 小于范本基准 {sample.suffix_min_len}"
        )

    for r, txt in suffix_run_texts(candidate.paragraph, label):
        if txt and not run_is_underlined(r):
            failures.append(f"{label}: 冒号后存在未加下划线的 run：{txt!r}")
        if " " in txt and not run_has_preserved_space(r):
            failures.append(f"{label}: 下划线空格未使用 xml:space='preserve'")


def compare_fixed_slot_group(
    failures: list[str],
    *,
    group_name: str,
    labels: list[str],
    sample_paragraphs: list[etree._Element],
    candidate_paragraphs: list[etree._Element],
) -> None:
    sample_widths: dict[str, int] = {}
    candidate_widths: dict[str, int] = {}

    for label in labels:
        sample = find_line(sample_paragraphs, label)
        candidate = find_line(candidate_paragraphs, label)
        if sample is None:
            failures.append(f"{group_name}/{label}: 范本中未找到字段行")
            continue
        if candidate is None:
            failures.append(f"{group_name}/{label}: 候选文档中未找到字段行")
            continue
        sample_widths[label] = display_width(underlined_suffix_text(sample.paragraph, label))
        candidate_widths[label] = display_width(underlined_suffix_text(candidate.paragraph, label))

    if not sample_widths or not candidate_widths:
        return

    expected_width = max(sample_widths.values())
    for label, actual_width in candidate_widths.items():
        if actual_width != expected_width:
            failures.append(
                f"{group_name}/{label}: 下划线固定槽显示宽度 {actual_width}，应与范本同组宽度 {expected_width} 一致"
            )

    unique_candidate_widths = sorted(set(candidate_widths.values()))
    if len(unique_candidate_widths) > 1:
        failures.append(
            f"{group_name}: 同组字段下划线宽度不一致，当前为 {', '.join(map(str, unique_candidate_widths))}"
        )


def run_checks(
    sample_path: Path,
    candidate_path: Path,
    *,
    student_id: str | None,
    archive_no: str | None,
) -> list[str]:
    sample_paragraphs = load_paragraphs(sample_path)
    candidate_paragraphs = load_paragraphs(candidate_path)

    sample_student = find_line(sample_paragraphs, "学号：")
    sample_archive = find_line(sample_paragraphs, "档号：")
    candidate_student = find_line(candidate_paragraphs, "学号：")
    candidate_archive = find_line(candidate_paragraphs, "档号：")

    failures: list[str] = []
    compare_line(failures, sample_student, candidate_student, expected_value=student_id)
    compare_line(failures, sample_archive, candidate_archive, expected_value=archive_no)

    if sample_student and sample_archive and candidate_student and candidate_archive:
        sample_gap = sample_archive.index - sample_student.index
        candidate_gap = candidate_archive.index - candidate_student.index
        if candidate_gap != sample_gap:
            failures.append(f"学号/档号: 两行间距结构 {candidate_gap} 与范本 {sample_gap} 不一致")

    compare_fixed_slot_group(
        failures,
        group_name="学号/档号",
        labels=["学号：", "档号："],
        sample_paragraphs=sample_paragraphs,
        candidate_paragraphs=candidate_paragraphs,
    )
    compare_fixed_slot_group(
        failures,
        group_name="封面信息栏",
        labels=["院（系）名称：", "专业名称：", "学生姓名：", "指导教师："],
        sample_paragraphs=sample_paragraphs,
        candidate_paragraphs=candidate_paragraphs,
    )

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare the Chinese cover identifier lines against the sample DOCX."
    )
    parser.add_argument("--sample", required=True, type=Path, help="Path to 范本.docx.")
    parser.add_argument("--candidate", required=True, type=Path, help="Path to the generated DOCX.")
    parser.add_argument("--student-id", help="Expected student number in the candidate.")
    parser.add_argument("--archive-no", help="Expected archive number in the candidate.")
    args = parser.parse_args()

    failures = run_checks(
        args.sample.resolve(),
        args.candidate.resolve(),
        student_id=args.student_id,
        archive_no=args.archive_no,
    )

    if failures:
        print("COVER_SAMPLE_COMPARISON=FAIL")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("COVER_SAMPLE_COMPARISON=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
