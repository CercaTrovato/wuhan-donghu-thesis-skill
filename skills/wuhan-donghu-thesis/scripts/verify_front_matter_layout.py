# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

import win32com.client
from lxml import etree


@dataclass
class ParaInfo:
    index: int
    page: int
    text: str
    line_rule: int
    line: float
    before: float
    after: float
    first: float

    @property
    def compact(self) -> str:
        return "".join(self.text.split()).replace("\u3000", "")


def clean_text(value: str) -> str:
    return value.replace("\r", "").replace("\a", "").strip()


def find_first(paragraphs: list[ParaInfo], needle: str, min_page: int = 1) -> ParaInfo | None:
    compact_needle = "".join(needle.split()).replace("\u3000", "")
    for para in paragraphs:
        if para.page >= min_page and compact_needle in para.compact:
            return para
    return None


def empty_before(paragraphs: list[ParaInfo], para: ParaInfo, count: int) -> bool:
    before = [p for p in paragraphs if p.index < para.index][-count:]
    return len(before) == count and all(p.text == "" for p in before)


def check_para_format(failures: list[str], label: str, para: ParaInfo | None, *, line_rule: int, line: float, before: float | None = None, after: float | None = None, first: float | None = None, tol: float = 0.6) -> None:
    if para is None:
        return
    if para.line_rule != line_rule:
        failures.append(f"{label}: 行距规则应为 {line_rule}，实际 {para.line_rule}")
    if abs(para.line - line) > tol:
        failures.append(f"{label}: 行距值应为 {line}，实际 {para.line}")
    if before is not None and abs(para.before - before) > tol:
        failures.append(f"{label}: 段前应为 {before}，实际 {para.before}")
    if after is not None and abs(para.after - after) > tol:
        failures.append(f"{label}: 段后应为 {after}，实际 {para.after}")
    if first is not None and abs(para.first - first) > tol:
        failures.append(f"{label}: 首行缩进应为 {first}，实际 {para.first}")


def run_checks(docx_path: Path, title_cn: str, title_en: str) -> list[str]:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path), False, True, False)
        doc.Repaginate()
        paragraphs: list[ParaInfo] = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs.Item(i)
            page = para.Range.Information(3)
            text = clean_text(para.Range.Text)
            fmt = para.Format
            paragraphs.append(
                ParaInfo(
                    i,
                    page,
                    text,
                    int(fmt.LineSpacingRule),
                    round(float(fmt.LineSpacing), 1),
                    round(float(fmt.SpaceBefore), 1),
                    round(float(fmt.SpaceAfter), 1),
                    round(float(fmt.FirstLineIndent), 1),
                )
            )
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()

    failures: list[str] = []
    with ZipFile(docx_path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for idx, sect in enumerate(root.xpath("//w:sectPr", namespaces=ns), 1):
            grid = sect.find("w:docGrid", ns)
            if grid is None:
                failures.append(f"第 {idx} 节: 缺少 docGrid")
                continue
            line_pitch = grid.get(f"{{{ns['w']}}}linePitch")
            grid_type = grid.get(f"{{{ns['w']}}}type")
            if line_pitch != "312":
                failures.append(f"第 {idx} 节: docGrid linePitch 应为 312，实际 {line_pitch}")
            if grid_type != "linesAndChars":
                failures.append(f"第 {idx} 节: docGrid type 应为 linesAndChars，实际 {grid_type}")

    declaration = find_first(paragraphs, "独创性声明", 3)
    authorization = find_first(paragraphs, "学位论文版权使用授权书", 3)
    cn_title = find_first(paragraphs, title_cn, 3)
    cn_abs = find_first(paragraphs, "摘要", 3)
    cn_keywords = find_first(paragraphs, "关键词：", 3)
    en_title = find_first(paragraphs, title_en, 3)
    en_abs = find_first(paragraphs, "ABSTRACT", 3)
    en_keywords = find_first(paragraphs, "Key words:", 3)
    toc = find_first(paragraphs, "目录", 3)
    declaration_body = find_first(paragraphs, "本人声明所呈交的学位论文", 3)
    declaration_signature = find_first(paragraphs, "学位论文作者签名：", 3)
    declaration_date = find_first(paragraphs, "日期：", 3)
    authorization_body = find_first(paragraphs, "本学位论文作者完全了解学校有关保留", 3)
    confidentiality = find_first(paragraphs, "保密", 3)

    expected_pages = [
        ("独创性声明", declaration, 3),
        ("学位论文版权使用授权书", authorization, 3),
        ("中文题目", cn_title, 4),
        ("中文摘要标题", cn_abs, 4),
        ("中文关键词", cn_keywords, 4),
        ("英文题目", en_title, 5),
        ("ABSTRACT", en_abs, 5),
    ]
    for label, para, expected_page in expected_pages:
        if para is None:
            failures.append(f"{label}: 未找到")
        elif para.page != expected_page:
            failures.append(f"{label}: 应在第 {expected_page} 页，实际第 {para.page} 页")

    if toc is None:
        failures.append("目录: 未找到")
    elif en_abs is not None and toc.page <= en_abs.page:
        failures.append(f"目录: 应在英文摘要之后，实际第 {toc.page} 页")
    if en_keywords is None:
        failures.append("英文关键词: 未找到")
    elif en_abs is not None and en_keywords.page < en_abs.page:
        failures.append(f"英文关键词: 应在 ABSTRACT 之后，实际第 {en_keywords.page} 页")
    elif toc is not None and en_keywords.page >= toc.page:
        failures.append(f"英文关键词: 应在目录之前，实际第 {en_keywords.page} 页")

    page3_text = "\n".join(p.text for p in paragraphs if p.page == 3)
    if title_cn in page3_text or "摘" in page3_text:
        failures.append("第 3 页: 不应包含中文题目或摘要内容")
    for required in ["保密", "不保密", "指导教师签名"]:
        if required not in page3_text:
            failures.append(f"授权书固定内容: 缺少“{required}”")

    if cn_title is not None and not empty_before(paragraphs, cn_title, 2):
        failures.append("中文摘要页: 中文题目前应有两个空段")
    if en_title is not None and not empty_before(paragraphs, en_title, 2):
        failures.append("英文摘要页: 英文题目前应有两个空段")

    for label, first, second in [
        ("中文摘要页", cn_title, cn_abs),
        ("英文摘要页", en_title, en_abs),
    ]:
        if first is not None and second is not None:
            between = [p for p in paragraphs if first.index < p.index < second.index]
            if not any(p.text == "" for p in between):
                failures.append(f"{label}: 题目和摘要标题之间应有空段")

    check_para_format(failures, "独创性声明标题", declaration, line_rule=4, line=23, before=0, after=0, first=0)
    check_para_format(failures, "独创性声明正文", declaration_body, line_rule=1, line=18, before=0.3, after=0, first=28)
    check_para_format(failures, "独创性声明签名", declaration_signature, line_rule=1, line=18, after=6, first=228)
    check_para_format(failures, "独创性声明日期", declaration_date, line_rule=1, line=18, after=6, first=257.5)
    check_para_format(failures, "授权书标题", authorization, line_rule=1, line=18, before=7.8, after=7.8, first=0)
    check_para_format(failures, "授权书正文", authorization_body, line_rule=1, line=18, before=0, after=0, first=28)
    check_para_format(failures, "授权书保密行", confidentiality, line_rule=4, line=23, before=0, after=0, first=0)
    check_para_format(failures, "中文摘要标题", cn_abs, line_rule=0, line=12, before=0, after=0, first=0)
    check_para_format(failures, "中文关键词", cn_keywords, line_rule=0, line=12, before=0, after=0, first=0)
    check_para_format(failures, "ABSTRACT", en_abs, line_rule=0, line=12, before=0, after=0, first=0)
    check_para_format(failures, "英文关键词", en_keywords, line_rule=1, line=18, before=0, after=0, first=0)

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify Wuhan Donghu thesis front matter pagination/layout.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--title-cn", required=True, help="Chinese thesis title exactly as it appears on the abstract page.")
    parser.add_argument("--title-en", required=True, help="English thesis title exactly as it appears on the English abstract page.")
    args = parser.parse_args()

    failures = run_checks(args.docx.resolve(), args.title_cn, args.title_en)
    if failures:
        print("FRONT_MATTER_CHECK=FAIL")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("FRONT_MATTER_CHECK=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
