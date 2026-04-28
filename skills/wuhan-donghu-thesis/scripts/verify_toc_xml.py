# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"


def qn(local_name: str) -> str:
    return W + local_name


def attrs(element):
    if element is None:
        return {}
    return {key.split("}", 1)[-1]: value for key, value in element.attrib.items()}


def verify(path: Path, expected_tab_pos: str, expected_line: str) -> int:
    with ZipFile(path) as zf:
        document_xml = etree.fromstring(zf.read("word/document.xml"))

    toc_paragraphs = []
    for paragraph in document_xml.findall(".//w:p", NS):
        p_style = paragraph.find("w:pPr/w:pStyle", NS)
        if p_style is None:
            continue
        style_id = p_style.get(qn("val"))
        if style_id in {"TOC1", "TOC2", "TOC3"}:
            toc_paragraphs.append((style_id, paragraph))

    bad_ind = 0
    bad_spacing = 0
    bad_tabs = 0
    hardcoded_leaders = 0

    for _, paragraph in toc_paragraphs:
        text = "".join(t.text or "" for t in paragraph.findall(".//w:t", NS))
        if "\u2026" in text:
            hardcoded_leaders += 1

        ppr = paragraph.find("w:pPr", NS)
        ind = attrs(ppr.find("w:ind", NS) if ppr is not None else None)
        spacing = attrs(ppr.find("w:spacing", NS) if ppr is not None else None)
        tabs = ppr.xpath("./w:tabs/w:tab", namespaces=NS) if ppr is not None else []
        tab_attrs = attrs(tabs[0]) if tabs else {}

        if ind.get("left") != "0" or ind.get("firstLine") != "0":
            bad_ind += 1
        if spacing.get("line") != expected_line or spacing.get("lineRule") != "exact":
            bad_spacing += 1
        if tab_attrs.get("val") != "right" or tab_attrs.get("leader") != "dot" or tab_attrs.get("pos") != expected_tab_pos:
            bad_tabs += 1

    levels = sorted({style_id for style_id, _ in toc_paragraphs})
    print(f"TOC_PARAS={len(toc_paragraphs)}")
    print(f"LEVELS={','.join(levels)}")
    print(f"BAD_IND={bad_ind}")
    print(f"BAD_SPACING={bad_spacing}")
    print(f"BAD_TABS={bad_tabs}")
    print(f"HARDCODED_LEADERS={hardcoded_leaders}")

    if not toc_paragraphs:
        return 1
    if bad_ind or bad_spacing or bad_tabs or hardcoded_leaders:
        return 1
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx_path")
    parser.add_argument("--tab-pos", default="8504")
    parser.add_argument("--line", default="460")
    args = parser.parse_args()
    return verify(Path(args.docx_path), args.tab_pos, args.line)


if __name__ == "__main__":
    raise SystemExit(main())
