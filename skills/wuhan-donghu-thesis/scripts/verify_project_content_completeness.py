# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


@dataclass
class FigureContext:
    caption: str
    text_before: str
    alt_text: str


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def paragraph_alt_text(p: ET.Element) -> str:
    values: list[str] = []
    for doc_pr in p.findall(".//wp:docPr", NS):
        for attr in ("title", "descr", "name"):
            value = doc_pr.attrib.get(attr)
            if value:
                values.append(value)
    return " ".join(values)


def read_docx_text_and_figures(docx_path: Path) -> tuple[str, list[FigureContext]]:
    with ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    all_text: list[str] = []
    figures: list[FigureContext] = []
    recent_text: list[str] = []
    last_alt = ""
    for child in root.findall(".//w:body/*", NS):
        if not child.tag.endswith("}p"):
            if child.tag.endswith("}tbl"):
                table_text = paragraph_text(child)
                if table_text:
                    all_text.append(table_text)
                    recent_text.append(table_text)
                    recent_text = recent_text[-4:]
            continue
        text = paragraph_text(child)
        alt = paragraph_alt_text(child)
        if alt:
            last_alt = alt
        if text:
            all_text.append(text)
            if re.match(r"^图\d+-\d+", text):
                figures.append(FigureContext(text, " ".join(recent_text[-3:]), last_alt))
                last_alt = ""
            else:
                recent_text.append(text)
                recent_text = recent_text[-4:]
    return "\n".join(all_text), figures


def load_requirements(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("requirements JSON must be an object")
    return data


def check_required_terms(text: str, items: list) -> list[str]:
    missing: list[str] = []
    for item in items:
        if isinstance(item, str):
            label = item
            terms = [item]
            mode = "all"
        else:
            label = item.get("label") or item.get("term") or item.get("terms") or item.get("any")
            terms = item.get("terms")
            mode = "all"
            if terms is None:
                terms = item.get("any")
                mode = "any"
            if isinstance(terms, str):
                terms = [terms]
            if not isinstance(terms, list) or not terms:
                missing.append(f"{label}: invalid requirement")
                continue
        if mode == "any":
            ok = any(str(term) in text for term in terms)
        else:
            ok = all(str(term) in text for term in terms)
        if not ok:
            missing.append(str(label))
    return missing


def find_figure(figures: list[FigureContext], pattern: str) -> FigureContext | None:
    compiled = re.compile(pattern)
    for figure in figures:
        if compiled.search(figure.caption):
            return figure
    return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Verify thesis content covers project evidence, source functions, and ER diagram themes."
    )
    parser.add_argument("docx_path")
    parser.add_argument("--requirements", required=True, help="JSON file describing required terms and ER terms")
    args = parser.parse_args(argv)

    docx_path = Path(args.docx_path)
    requirements = load_requirements(Path(args.requirements))
    visible_text, figures = read_docx_text_and_figures(docx_path)

    missing: list[str] = []
    missing.extend(f"正文缺少：{label}" for label in check_required_terms(visible_text, requirements.get("required_terms", [])))

    er_terms = requirements.get("er_required_terms", [])
    if er_terms:
        pattern = requirements.get("er_caption_pattern", r"图\d+-\d+.*E-R")
        er_figure = find_figure(figures, pattern)
        if er_figure is None:
            missing.append(f"未找到 ER 图题：{pattern}")
        else:
            er_context = "\n".join([er_figure.caption, er_figure.text_before, er_figure.alt_text])
            for term in er_terms:
                if str(term) not in er_context:
                    missing.append(f"ER 图缺少主题：{term}")

    print(f"VISIBLE_TEXT_CHARS={len(re.sub(r'\\s+', '', visible_text))}")
    print(f"FIGURE_CONTEXTS={len(figures)}")
    if missing:
        print("PROJECT_CONTENT_COMPLETENESS=FAIL")
        for item in missing:
            print(f"- {item}")
        return 1

    print("PROJECT_CONTENT_COMPLETENESS=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
