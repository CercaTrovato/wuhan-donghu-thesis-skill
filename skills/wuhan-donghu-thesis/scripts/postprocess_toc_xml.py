# -*- coding: utf-8 -*-
from __future__ import annotations

import shutil
import sys
import tempfile
import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"
TABLE_WIDTH_DXA = "8720"
TABLE_STYLE_ID = "af"
TABLE_WIDTHS_BY_COLS = {
    3: ["1744", "2616", "4360"],
    5: ["1245", "2491", "2492", "1246", "1246"],
    6: ["1453", "1453", "1453", "1453", "1262", "1646"],
}


def qn(local_name: str) -> str:
    return W + local_name


def remove_children(parent, local_name: str) -> None:
    for child in list(parent.findall(f"w:{local_name}", NS)):
        parent.remove(child)


def first_child_index(parent, local_names: set[str]) -> int | None:
    for idx, child in enumerate(parent):
        if etree.QName(child).localname in local_names:
            return idx
    return None


def insert_before(parent, child, before: set[str]) -> None:
    index = first_child_index(parent, before)
    if index is None:
        parent.append(child)
    else:
        parent.insert(index, child)


def insert_after(parent, child, after: set[str], *, before: set[str] | None = None) -> None:
    last_after = -1
    for idx, existing in enumerate(parent):
        if etree.QName(existing).localname in after:
            last_after = idx
    if last_after >= 0:
        parent.insert(last_after + 1, child)
    elif before:
        insert_before(parent, child, before)
    else:
        parent.append(child)


def get_or_add(parent, local_name: str):
    child = parent.find(f"w:{local_name}", NS)
    if child is None:
        child = etree.SubElement(parent, qn(local_name))
    return child


def ensure_ppr(element):
    ppr = element.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        element.insert(0, ppr)
    return ppr


def ensure_style_ppr(style):
    ppr = style.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(qn("pPr"))
        insert_after(
            style,
            ppr,
            {
                "name",
                "aliases",
                "basedOn",
                "next",
                "link",
                "autoRedefine",
                "hidden",
                "uiPriority",
                "semiHidden",
                "unhideWhenUsed",
                "qFormat",
                "locked",
                "personal",
                "personalCompose",
                "personalReply",
                "rsid",
            },
            before={"rPr", "tblPr", "trPr", "tcPr", "tblStylePr"},
        )
    return ppr


def ensure_tcpr(element):
    tcpr = element.find("w:tcPr", NS)
    if tcpr is None:
        tcpr = etree.Element(qn("tcPr"))
        element.insert(0, tcpr)
    return tcpr


def ensure_rpr(style):
    rpr = style.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.SubElement(style, qn("rPr"))
    return rpr


def ensure_paragraph_rpr(paragraph):
    ppr = ensure_ppr(paragraph)
    rpr = ppr.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.Element(qn("rPr"))
        append_ppr_child(ppr, rpr, "rPr")
    return rpr


def set_rfonts(rpr, east_asia: str, west: str = "Times New Roman") -> None:
    remove_children(rpr, "rFonts")
    fonts = etree.SubElement(rpr, qn("rFonts"))
    fonts.set(qn("ascii"), west)
    fonts.set(qn("hAnsi"), west)
    fonts.set(qn("eastAsia"), east_asia)
    rpr.remove(fonts)
    insert_after(
        rpr,
        fonts,
        {"rStyle"},
        before={
            "b",
            "bCs",
            "i",
            "iCs",
            "caps",
            "smallCaps",
            "strike",
            "dstrike",
            "outline",
            "shadow",
            "emboss",
            "imprint",
            "noProof",
            "snapToGrid",
            "vanish",
            "webHidden",
            "color",
            "spacing",
            "w",
            "kern",
            "position",
            "sz",
            "szCs",
            "highlight",
            "u",
            "effect",
            "bdr",
            "shd",
            "fitText",
            "vertAlign",
            "rtl",
            "cs",
            "em",
            "lang",
            "eastAsianLayout",
            "specVanish",
            "oMath",
            "rPrChange",
        },
    )


def set_font_size(rpr, half_points: str) -> None:
    remove_children(rpr, "sz")
    remove_children(rpr, "szCs")
    size = etree.SubElement(rpr, qn("sz"))
    size.set(qn("val"), half_points)
    size_cs = etree.SubElement(rpr, qn("szCs"))
    size_cs.set(qn("val"), half_points)
    rpr.remove(size)
    rpr.remove(size_cs)
    insert_before(
        rpr,
        size,
        {
            "highlight",
            "u",
            "effect",
            "bdr",
            "shd",
            "fitText",
            "vertAlign",
            "rtl",
            "cs",
            "em",
            "lang",
            "eastAsianLayout",
            "specVanish",
            "oMath",
            "rPrChange",
        },
    )
    insert_after(rpr, size_cs, {"sz"}, before={"highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"})


def append_ppr_child(ppr, child, local_name: str) -> None:
    order_after = {
        "tabs": {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd"},
        "spacing": {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid"},
        "ind": {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing"},
        "jc": {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap"},
        "rPr": {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle"},
    }
    order_before = {
        "tabs": {"suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"},
        "spacing": {"ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"},
        "ind": {"contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"},
        "jc": {"textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"},
        "rPr": {"sectPr", "pPrChange"},
    }
    insert_after(ppr, child, order_after[local_name], before=order_before[local_name])


def apply_toc_paragraph_format(ppr, *, tab_pos: str = "8504", line: str = "460") -> None:
    remove_children(ppr, "tabs")
    remove_children(ppr, "spacing")
    remove_children(ppr, "ind")

    tabs = etree.Element(qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    tab.set(qn("pos"), tab_pos)
    append_ppr_child(ppr, tabs, "tabs")

    spacing = etree.Element(qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), line)
    spacing.set(qn("lineRule"), "exact")
    append_ppr_child(ppr, spacing, "spacing")

    ind = etree.Element(qn("ind"))
    ind.set(qn("left"), "0")
    ind.set(qn("firstLine"), "0")
    append_ppr_child(ppr, ind, "ind")


def apply_toc_run_format(rpr, east_asia: str) -> None:
    set_rfonts(rpr, east_asia)
    set_font_size(rpr, "24")


def set_border(parent, name: str, val: str, size: str | None = None, color: str = "000000") -> None:
    remove_children(parent, name)
    border = etree.SubElement(parent, qn(name))
    border.set(qn("val"), val)
    if size is not None:
        border.set(qn("color"), color)
        border.set(qn("sz"), size)
        border.set(qn("space"), "0")
    elif val in {"none", "nil"}:
        border.set(qn("color"), "auto")
        border.set(qn("sz"), "0")
        border.set(qn("space"), "0")


def patch_table_caption(paragraph) -> None:
    ppr = ensure_ppr(paragraph)
    remove_children(ppr, "jc")
    jc = etree.Element(qn("jc"))
    jc.set(qn("val"), "center")
    append_ppr_child(ppr, jc, "jc")
    remove_children(ppr, "spacing")
    spacing = etree.Element(qn("spacing"))
    spacing.set(qn("before"), "156")
    spacing.set(qn("beforeLines"), "50")
    spacing.set(qn("after"), "156")
    spacing.set(qn("afterLines"), "50")
    spacing.set(qn("line"), "240")
    spacing.set(qn("lineRule"), "auto")
    append_ppr_child(ppr, spacing, "spacing")
    p_rpr = ensure_paragraph_rpr(paragraph)
    set_rfonts(p_rpr, "\u9ed1\u4f53")
    set_font_size(p_rpr, "24")
    for run in paragraph.findall("w:r", NS):
        if not run.findall("w:t", NS):
            continue
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = etree.Element(qn("rPr"))
            run.insert(0, rpr)
        set_rfonts(rpr, "\u9ed1\u4f53")
        set_font_size(rpr, "24")


def patch_table(table) -> None:
    rows = table.findall("w:tr", NS)
    if not rows:
        return
    col_count = max(len(row.findall("w:tc", NS)) for row in rows)
    widths = TABLE_WIDTHS_BY_COLS.get(col_count)
    if not widths:
        base = int(int(TABLE_WIDTH_DXA) / col_count)
        widths = [str(base)] * col_count
        widths[-1] = str(int(TABLE_WIDTH_DXA) - base * (col_count - 1))

    tbl_pr = table.find("w:tblPr", NS)
    if tbl_pr is None:
        tbl_pr = etree.Element(qn("tblPr"))
        table.insert(0, tbl_pr)
    remove_children(tbl_pr, "tblStyle")
    style = etree.Element(qn("tblStyle"))
    style.set(qn("val"), TABLE_STYLE_ID)
    tbl_pr.insert(0, style)
    remove_children(tbl_pr, "tblW")
    tbl_w = etree.Element(qn("tblW"))
    tbl_w.set(qn("w"), TABLE_WIDTH_DXA)
    tbl_w.set(qn("type"), "dxa")
    insert_after(
        tbl_pr,
        tbl_w,
        {"tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize"},
        before={"jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"},
    )
    remove_children(tbl_pr, "tblLayout")
    layout = etree.Element(qn("tblLayout"))
    layout.set(qn("type"), "fixed")
    insert_after(
        tbl_pr,
        layout,
        {"tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd"},
        before={"tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"},
    )
    remove_children(tbl_pr, "tblBorders")
    borders = etree.Element(qn("tblBorders"))
    insert_after(
        tbl_pr,
        borders,
        {"tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd"},
        before={"shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"},
    )
    set_border(borders, "top", "single", "8")
    set_border(borders, "left", "none")
    set_border(borders, "bottom", "single", "8")
    set_border(borders, "right", "none")
    set_border(borders, "insideH", "none")
    set_border(borders, "insideV", "none")

    old_grid = table.find("w:tblGrid", NS)
    if old_grid is not None:
        table.remove(old_grid)
    grid = etree.Element(qn("tblGrid"))
    insert_at = 1 if table.find("w:tblPr", NS) is not None else 0
    table.insert(insert_at, grid)
    for width in widths:
        col = etree.SubElement(grid, qn("gridCol"))
        col.set(qn("w"), width)

    for row_index, row in enumerate(rows):
        cells = row.findall("w:tc", NS)
        for cell_index, cell in enumerate(cells):
            width = widths[min(cell_index, len(widths) - 1)]
            tcpr = ensure_tcpr(cell)
            remove_children(tcpr, "tcW")
            tcw = etree.Element(qn("tcW"))
            tcw.set(qn("w"), width)
            tcw.set(qn("type"), "dxa")
            insert_after(tcpr, tcw, {"cnfStyle"}, before={"gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"})
            remove_children(tcpr, "tcBorders")
            tc_borders = etree.Element(qn("tcBorders"))
            insert_after(tcpr, tc_borders, {"cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge"}, before={"shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"})
            if row_index == 0:
                set_border(tc_borders, "bottom", "single", "6")
            elif row_index == 1:
                set_border(tc_borders, "top", "single", "6")
            set_border(tc_borders, "tl2br", "nil")
            set_border(tc_borders, "tr2bl", "nil")
            for paragraph in cell.findall("w:p", NS):
                ppr = ensure_ppr(paragraph)
                remove_children(ppr, "spacing")
                spacing = etree.Element(qn("spacing"))
                spacing.set(qn("line"), "240")
                spacing.set(qn("lineRule"), "auto")
                append_ppr_child(ppr, spacing, "spacing")
                remove_children(ppr, "jc")
                jc = etree.Element(qn("jc"))
                jc.set(qn("val"), "center")
                append_ppr_child(ppr, jc, "jc")
                for run in paragraph.findall("w:r", NS):
                    if not run.findall("w:t", NS):
                        continue
                    rpr = run.find("w:rPr", NS)
                    if rpr is None:
                        rpr = etree.Element(qn("rPr"))
                        run.insert(0, rpr)
                    set_rfonts(rpr, "\u5b8b\u4f53", "\u5b8b\u4f53")
                    set_font_size(rpr, "21")


def patch_tables(document_root) -> None:
    body = document_root.find(".//w:body", NS)
    if body is None:
        return
    children = list(body)
    for idx, node in enumerate(children):
        if node.tag != qn("tbl"):
            continue
        for previous in reversed(children[:idx]):
            if previous.tag == qn("p") and "".join(t.text or "" for t in previous.findall(".//w:t", NS)).strip():
                patch_table_caption(previous)
                break
        patch_table(node)


def patch_styles(styles_root, *, tab_pos: str = "8504", line: str = "460") -> None:
    for style_id in ("TOC1", "TOC2", "TOC3"):
        matches = styles_root.xpath(f'.//w:style[@w:styleId="{style_id}"]', namespaces=NS)
        if not matches:
            continue
        style = matches[0]
        apply_toc_paragraph_format(ensure_style_ppr(style), tab_pos=tab_pos, line=line)
        east_asia = "\u9ed1\u4f53" if style_id == "TOC1" else "\u5b8b\u4f53"
        apply_toc_run_format(ensure_rpr(style), east_asia)


def patch_document(document_root, *, tab_pos: str = "8504", line: str = "460") -> None:
    for paragraph in document_root.findall(".//w:p", NS):
        text = "".join(t.text or "" for t in paragraph.findall(".//w:t", NS))
        if text == "\u76ee\u3000\u5f55":
            ppr = ensure_ppr(paragraph)
            remove_children(ppr, "spacing")
            spacing = etree.Element(qn("spacing"))
            spacing.set(qn("before"), "400")
            spacing.set(qn("after"), "230")
            spacing.set(qn("line"), line)
            spacing.set(qn("lineRule"), "exact")
            append_ppr_child(ppr, spacing, "spacing")
            continue

        p_style = paragraph.find("w:pPr/w:pStyle", NS)
        if p_style is None:
            continue
        style_id = p_style.get(qn("val"))
        if style_id not in {"TOC1", "TOC2", "TOC3"}:
            continue
        ppr = ensure_ppr(paragraph)
        apply_toc_paragraph_format(ppr, tab_pos=tab_pos, line=line)
    patch_tables(document_root)


def rewrite_docx(path: Path, *, tab_pos: str = "8504", line: str = "460") -> None:
    with ZipFile(path, "r") as zin:
        document_xml = etree.fromstring(zin.read("word/document.xml"))
        styles_xml = etree.fromstring(zin.read("word/styles.xml"))

        patch_document(document_xml, tab_pos=tab_pos, line=line)
        patch_styles(styles_xml, tab_pos=tab_pos, line=line)

        document_bytes = etree.tostring(
            document_xml, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        styles_bytes = etree.tostring(
            styles_xml, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp_path = Path(tmp.name)

        with ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, document_bytes)
                elif item.filename == "word/styles.xml":
                    zout.writestr(item, styles_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))

    shutil.move(str(tmp_path), path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Freeze Wuhan Donghu thesis TOC/table OpenXML formatting.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--tab-pos", default="8504", help="TOC right tab stop in twips, default: 8504")
    parser.add_argument("--line", default="460", help="TOC line spacing in twips, default: 460")
    args = parser.parse_args()

    rewrite_docx(args.docx, tab_pos=str(args.tab_pos), line=str(args.line))
    print(args.docx)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
