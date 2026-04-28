#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import re
import sys
import textwrap
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
EXPECTED_TABLE_WIDTH_DXA = "8720"
EXPECTED_TABLE_LINE = "240"
EXPECTED_TABLE_LINE_RULE = "auto"
EXPECTED_FONT_SIZE = "21"
EXPECTED_OUTER_BORDER = "8"
EXPECTED_HEADER_BORDER = "6"
EXPECTED_TABLE_STYLE = "af"
EXPECTED_CAPTION_BEFORE = "156"
EXPECTED_CAPTION_AFTER = "156"
EXPECTED_FIGURE_CAPTION_STYLE = "ThesisCaption"
EXPECTED_IMAGE_STYLE = "ImageParagraph"
EXPECTED_GRID_BY_COLS = {
    3: ["1744", "2616", "4360"],
    5: ["1245", "2491", "2492", "1246", "1246"],
    6: ["1700", "1300", "850", "1200", "1250", "2420"],
}


def attr(el: ET.Element | None, name: str) -> str | None:
    if el is None:
        return None
    return el.attrib.get(W + name)


def children(el: ET.Element | None, tag: str) -> list[ET.Element]:
    if el is None:
        return []
    return list(el.findall(W + tag))


def fail(message: str, failures: list[str]) -> None:
    failures.append(message)


def check_build_script(
    build_script: Path,
    failures: list[str],
    *,
    require_er_function: bool = False,
    forbidden_er_tokens: list[str] | None = None,
    check_api_flow_blank_space: bool = False,
) -> None:
    if not build_script.exists():
        fail(f"build script not found: {build_script}", failures)
        return

    text = build_script.read_text(encoding="utf-8")
    er_match = re.search(r"def\s+er\s*\(draw\):(?P<body>.*?)(?:\n\s+def\s+\w+\s*\(draw\):|\n\s+diagrams\s*=)", text, re.S)
    if not er_match and require_er_function:
        fail("cannot find er(draw) function in build script", failures)
    elif er_match and forbidden_er_tokens:
        body = er_match.group("body")
        hits = [token for token in forbidden_er_tokens if re.search(rf"\b{re.escape(token)}\b", body)]
        if hits:
            fail("E-R diagram still contains English database labels: " + ", ".join(hits), failures)

    box_re = re.compile(
        r"box\s*\(\s*draw\s*,\s*\((?P<x1>\d+)\s*,\s*(?P<y1>\d+)\s*,\s*(?P<x2>\d+)\s*,\s*(?P<y2>\d+)\)\s*,\s*\"(?P<text>[^\"]+)\"(?P<args>[^\n]*)",
        re.S,
    )
    risky = []
    for match in box_re.finditer(text):
        x1, x2 = int(match.group("x1")), int(match.group("x2"))
        args = match.group("args")
        box_width = x2 - x1
        font_size = 28
        if "FONT_SMALL" in args:
            font_size = 22
        elif "FONT_BOLD" in args:
            font_size = 30
        max_chars_match = re.search(r"max_chars\s*=\s*(\d+)", args)
        max_chars = int(max_chars_match.group(1)) if max_chars_match else 10
        text_value = match.group("text").replace("\\n", "\n")
        wrapped_lines = [line for part in text_value.split("\n") for line in (textwrap.wrap(part, width=max_chars) or [""])]
        # Chinese glyphs are roughly one em wide; ASCII identifiers are narrower.
        # This catches Chinese label overflow without falsely rejecting IP/path text.
        estimated_width = max(
            (
                int(
                    sum((0.58 if ord(ch) < 128 else 1.05) * font_size for ch in line)
                )
                for line in wrapped_lines
            ),
            default=0,
        )
        if estimated_width > box_width - 32:
            risky.append(f"{text_value!r} estimated {estimated_width}px in {box_width}px box")
    if risky:
        fail("flow/chart text may overflow fixed boxes: " + "; ".join(risky), failures)

    api_match = re.search(r"def\s+api_flow\s*\(draw\):(?P<body>.*?)(?:\n\s+def\s+\w+\s*\(draw\):|\n\s+diagrams\s*=)", text, re.S)
    if api_match:
        api_body = api_match.group("body")
        y2_values = [
            int(m.group("y2"))
            for m in re.finditer(
                r"(?:box|rect_box)\s*\(\s*draw\s*,\s*\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*(?P<y2>\d+)\s*\)",
                api_body,
            )
        ]
        image_height = 900
        size_match = re.search(r"""["']api_flow\.png["']\s*:\s*\(\s*api_flow\s*,\s*\(\s*\d+\s*,\s*(?P<h>\d+)\s*\)""", text)
        if size_match:
            image_height = int(size_match.group("h"))
        if y2_values and image_height - max(y2_values) > 120:
            fail(
                f"api_flow image leaves too much blank space below content: height={image_height}, content_bottom={max(y2_values)}",
                failures,
            )
    elif check_api_flow_blank_space:
        fail("cannot find api_flow(draw) function in build script", failures)


def border_info(parent: ET.Element | None, name: str) -> tuple[str | None, str | None]:
    if parent is None:
        return None, None
    el = parent.find(W + name)
    return attr(el, "val"), attr(el, "sz")


def validate_spacing(p: ET.Element, context: str, failures: list[str]) -> None:
    spacing = p.find(W + "pPr/" + W + "spacing")
    line = attr(spacing, "line")
    rule = attr(spacing, "lineRule")
    if line != EXPECTED_TABLE_LINE or rule != EXPECTED_TABLE_LINE_RULE:
        fail(f"{context}: expected table paragraph spacing line={EXPECTED_TABLE_LINE} rule={EXPECTED_TABLE_LINE_RULE}, got line={line} rule={rule}", failures)


def validate_run(r: ET.Element, context: str, failures: list[str], header: bool) -> None:
    if not r.findall(W + "t"):
        return
    rpr = r.find(W + "rPr")
    size = attr(rpr.find(W + "sz") if rpr is not None else None, "val")
    if size != EXPECTED_FONT_SIZE:
        fail(f"{context}: expected 10.5pt font size w:sz={EXPECTED_FONT_SIZE}, got {size}", failures)
    if header and rpr is not None and rpr.find(W + "b") is not None:
        fail(f"{context}: table header should not be bold; sample uses regular Songti", failures)


def previous_nonempty_paragraph(body_children: list[ET.Element], table_index: int) -> ET.Element | None:
    for node in reversed(body_children[:table_index]):
        if node.tag == W + "p" and "".join(t.text or "" for t in node.findall(".//" + W + "t")).strip():
            return node
    return None


def validate_caption(paragraph: ET.Element | None, context: str, failures: list[str]) -> None:
    if paragraph is None:
        fail(f"{context}: table caption paragraph not found", failures)
        return
    p_pr = paragraph.find(W + "pPr")
    jc = p_pr.find(W + "jc") if p_pr is not None else None
    spacing = p_pr.find(W + "spacing") if p_pr is not None else None
    if attr(jc, "val") != "center":
        fail(f"{context}: caption should be centered, got {attr(jc, 'val')}", failures)
    if attr(spacing, "before") != EXPECTED_CAPTION_BEFORE or attr(spacing, "after") != EXPECTED_CAPTION_AFTER:
        fail(
            f"{context}: caption spacing should before/after={EXPECTED_CAPTION_BEFORE}, got before={attr(spacing, 'before')} after={attr(spacing, 'after')}",
            failures,
        )
    if attr(spacing, "line") != EXPECTED_TABLE_LINE or attr(spacing, "lineRule") != EXPECTED_TABLE_LINE_RULE:
        fail(
            f"{context}: caption line spacing should line={EXPECTED_TABLE_LINE} rule={EXPECTED_TABLE_LINE_RULE}, got line={attr(spacing, 'line')} rule={attr(spacing, 'lineRule')}",
            failures,
        )
    text_runs = [r for r in paragraph.findall(W + "r") if r.findall(W + "t")]
    if not text_runs:
        fail(f"{context}: caption has no text run", failures)
        return
    for run_index, run in enumerate(text_runs, 1):
        r_pr = run.find(W + "rPr")
        fonts = r_pr.find(W + "rFonts") if r_pr is not None else None
        if attr(fonts, "eastAsia") != "黑体":
            fail(f"{context}: caption run {run_index} should use 黑体 eastAsia font, got {attr(fonts, 'eastAsia')}", failures)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//" + W + "t")).strip()


def check_style_format(styles_root: ET.Element, failures: list[str]) -> None:
    style_map = {
        attr(style, "styleId"): style
        for style in styles_root.findall(W + "style")
        if attr(style, "styleId")
    }

    caption_style = style_map.get(EXPECTED_FIGURE_CAPTION_STYLE)
    if caption_style is None:
        fail(f"missing figure caption style {EXPECTED_FIGURE_CAPTION_STYLE}", failures)
    else:
        p_pr = caption_style.find(W + "pPr")
        r_pr = caption_style.find(W + "rPr")
        jc = p_pr.find(W + "jc") if p_pr is not None else None
        spacing = p_pr.find(W + "spacing") if p_pr is not None else None
        fonts = r_pr.find(W + "rFonts") if r_pr is not None else None
        size = r_pr.find(W + "sz") if r_pr is not None else None
        if attr(jc, "val") != "center":
            fail(f"{EXPECTED_FIGURE_CAPTION_STYLE}: figure captions should be centered", failures)
        if attr(spacing, "line") != "460" or attr(spacing, "lineRule") != "exact":
            fail(
                f"{EXPECTED_FIGURE_CAPTION_STYLE}: figure caption line spacing should be fixed 23pt, got line={attr(spacing, 'line')} rule={attr(spacing, 'lineRule')}",
                failures,
            )
        if attr(fonts, "eastAsia") != "黑体" or attr(size, "val") != "24":
            fail(
                f"{EXPECTED_FIGURE_CAPTION_STYLE}: figure captions should use 黑体 12pt, got eastAsia={attr(fonts, 'eastAsia')} size={attr(size, 'val')}",
                failures,
            )
        if r_pr is None or r_pr.find(W + "b") is None:
            fail(f"{EXPECTED_FIGURE_CAPTION_STYLE}: figure captions should be bold", failures)

    image_style = style_map.get(EXPECTED_IMAGE_STYLE)
    if image_style is None:
        fail(f"missing image paragraph style {EXPECTED_IMAGE_STYLE}", failures)
    else:
        p_pr = image_style.find(W + "pPr")
        jc = p_pr.find(W + "jc") if p_pr is not None else None
        spacing = p_pr.find(W + "spacing") if p_pr is not None else None
        if attr(jc, "val") != "center":
            fail(f"{EXPECTED_IMAGE_STYLE}: image paragraphs should be centered", failures)
        if attr(spacing, "line") != "240" or attr(spacing, "lineRule") != "auto":
            fail(
                f"{EXPECTED_IMAGE_STYLE}: image paragraphs should use auto/single line spacing, got line={attr(spacing, 'line')} rule={attr(spacing, 'lineRule')}",
                failures,
            )


def check_docx_figures(docx_path: Path, failures: list[str]) -> None:
    with ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        styles_root = ET.fromstring(zf.read("word/styles.xml"))

    check_style_format(styles_root, failures)

    body = root.find(".//" + W + "body")
    body_children = list(body) if body is not None else []
    figure_caption_re = re.compile(r"^图\d+-\d+\s+")
    figure_count = 0
    for idx, node in enumerate(body_children):
        if node.tag != W + "p":
            continue
        text = paragraph_text(node)
        if not figure_caption_re.match(text):
            continue
        figure_count += 1
        context = f"figure caption {figure_count} ({text})"
        p_style = node.find(W + "pPr/" + W + "pStyle")
        if attr(p_style, "val") != EXPECTED_FIGURE_CAPTION_STYLE:
            fail(f"{context}: expected paragraph style {EXPECTED_FIGURE_CAPTION_STYLE}, got {attr(p_style, 'val')}", failures)
        previous = body_children[idx - 1] if idx > 0 else None
        if previous is None or previous.tag != W + "p" or not previous.findall(".//" + W + "drawing"):
            fail(f"{context}: previous body paragraph should contain the figure image", failures)
            continue
        image_style = previous.find(W + "pPr/" + W + "pStyle")
        if attr(image_style, "val") != EXPECTED_IMAGE_STYLE:
            fail(f"{context}: preceding image paragraph should use {EXPECTED_IMAGE_STYLE}, got {attr(image_style, 'val')}", failures)

    if figure_count == 0:
        fail("document contains no figure captions matching 图N-N", failures)


def check_docx_tables(docx_path: Path, failures: list[str]) -> None:
    if not docx_path.exists():
        fail(f"docx not found: {docx_path}", failures)
        return

    with ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(".//" + W + "body")
    body_children = list(body) if body is not None else []
    tables = root.findall(".//" + W + "tbl")
    if not tables:
        fail("document contains no tables", failures)
        return

    for table_index, tbl in enumerate(tables, 1):
        context = f"table {table_index}"
        body_index = body_children.index(tbl) if tbl in body_children else -1
        validate_caption(previous_nonempty_paragraph(body_children, body_index), context, failures)
        tbl_pr = tbl.find(W + "tblPr")
        tbl_style = tbl_pr.find(W + "tblStyle") if tbl_pr is not None else None
        if attr(tbl_style, "val") != EXPECTED_TABLE_STYLE:
            fail(f"{context}: expected table style {EXPECTED_TABLE_STYLE}, got {attr(tbl_style, 'val')}", failures)
        tbl_w = tbl_pr.find(W + "tblW") if tbl_pr is not None else None
        if attr(tbl_w, "w") != EXPECTED_TABLE_WIDTH_DXA or attr(tbl_w, "type") != "dxa":
            fail(f"{context}: expected fixed width {EXPECTED_TABLE_WIDTH_DXA} dxa, got w={attr(tbl_w, 'w')} type={attr(tbl_w, 'type')}", failures)
        layout = tbl_pr.find(W + "tblLayout") if tbl_pr is not None else None
        if attr(layout, "type") != "fixed":
            fail(f"{context}: expected fixed table layout", failures)

        borders = tbl_pr.find(W + "tblBorders") if tbl_pr is not None else None
        for name in ("top", "bottom"):
            val, sz = border_info(borders, name)
            if val != "single" or sz != EXPECTED_OUTER_BORDER:
                fail(f"{context}: expected {name} border single size {EXPECTED_OUTER_BORDER}, got val={val} sz={sz}", failures)
        for name in ("left", "right", "insideH", "insideV"):
            val, _ = border_info(borders, name)
            if val != "none":
                fail(f"{context}: expected {name} border explicit none, got {val}", failures)

        grid = tbl.find(W + "tblGrid")
        grid_widths = [attr(gc, "w") for gc in children(grid, "gridCol")]
        if grid_widths and str(sum(int(w or "0") for w in grid_widths)) != EXPECTED_TABLE_WIDTH_DXA:
            fail(f"{context}: grid widths should sum to {EXPECTED_TABLE_WIDTH_DXA}, got {grid_widths}", failures)

        rows = children(tbl, "tr")
        if not rows:
            fail(f"{context}: table has no rows", failures)
            continue
        expected_grid = EXPECTED_GRID_BY_COLS.get(max(len(children(row, "tc")) for row in rows))
        if expected_grid and grid_widths != expected_grid:
            fail(f"{context}: expected {len(expected_grid)}-column sample grid {expected_grid}, got {grid_widths}", failures)
        for row_index, row in enumerate(rows, 1):
            cells = children(row, "tc")
            for cell_index, tc in enumerate(cells, 1):
                cell_context = f"{context} row {row_index} cell {cell_index}"
                tc_pr = tc.find(W + "tcPr")
                tc_w = tc_pr.find(W + "tcW") if tc_pr is not None else None
                if grid_widths and cell_index <= len(grid_widths) and attr(tc_w, "w") != grid_widths[cell_index - 1]:
                    fail(f"{cell_context}: tcW should match grid width {grid_widths[cell_index - 1]}, got {attr(tc_w, 'w')}", failures)
                tc_borders = tc_pr.find(W + "tcBorders") if tc_pr is not None else None
                for name in ("left", "right"):
                    val, _ = border_info(tc_borders, name)
                    if val not in (None, "none", "nil"):
                        fail(f"{cell_context}: expected no vertical cell border {name}, got {val}", failures)
                if row_index == 1:
                    val, sz = border_info(tc_borders, "top")
                    if val == "nil":
                        fail(f"{cell_context}: first-row top border must not be nil; sample relies on table top border", failures)
                    val, sz = border_info(tc_borders, "bottom")
                    if val != "single" or sz != EXPECTED_HEADER_BORDER:
                        fail(f"{cell_context}: expected header bottom border single size {EXPECTED_HEADER_BORDER}, got val={val} sz={sz}", failures)
                if row_index == 2:
                    val, sz = border_info(tc_borders, "top")
                    if val != "single" or sz != EXPECTED_HEADER_BORDER:
                        fail(f"{cell_context}: expected first body row top border single size {EXPECTED_HEADER_BORDER}, got val={val} sz={sz}", failures)
                if row_index == len(rows):
                    val, sz = border_info(tc_borders, "bottom")
                    if val == "nil":
                        fail(f"{cell_context}: last-row bottom border must not be nil; sample relies on table bottom border", failures)
                paragraphs = children(tc, "p")
                for p_index, p in enumerate(paragraphs, 1):
                    validate_spacing(p, f"{cell_context} paragraph {p_index}", failures)
                    for r_index, r in enumerate(children(p, "r"), 1):
                        validate_run(r, f"{cell_context} paragraph {p_index} run {r_index}", failures, header=(row_index == 1))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Verify Wuhan Donghu thesis figure/table layout constraints.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--build-script", type=Path)
    parser.add_argument("--require-er-function", action="store_true")
    parser.add_argument("--forbidden-er-token", action="append", default=[])
    parser.add_argument("--check-api-flow-blank-space", action="store_true")
    args = parser.parse_args(argv)

    failures: list[str] = []
    if args.build_script is not None:
        check_build_script(
            args.build_script,
            failures,
            require_er_function=args.require_er_function,
            forbidden_er_tokens=args.forbidden_er_token,
            check_api_flow_blank_space=args.check_api_flow_blank_space,
        )
    elif args.require_er_function or args.forbidden_er_token or args.check_api_flow_blank_space:
        fail("project-specific build-script checks requested but --build-script was not provided", failures)
    check_docx_figures(args.docx, failures)
    check_docx_tables(args.docx, failures)

    if failures:
        print("FAIL")
        for item in failures:
            print("- " + item)
        return 1
    print("PASS: figure/table layout checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
