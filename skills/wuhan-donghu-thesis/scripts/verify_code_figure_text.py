#!/usr/bin/env python3
"""Validate code figure source text and font coverage.

This check prevents screenshots/code figures from silently rendering Chinese
strings as missing-glyph boxes or question marks. It accepts either a generated
manifest JSON or a Python builder script containing a ``code_snippets`` dict.
"""

from __future__ import annotations

import argparse
import ast
import json
from pathlib import Path

from fontTools.ttLib import TTFont


BAD_MARKERS = ("\ufffd", "????", "□□", "■■")


def load_font_cmap(font_path: Path) -> set[int]:
    font = TTFont(str(font_path), fontNumber=0)
    cmap: set[int] = set()
    for table in font["cmap"].tables:
        cmap.update(table.cmap.keys())
    font.close()
    return cmap


def extract_from_python(path: Path) -> dict[str, str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    snippets: dict[str, str] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(target, ast.Name) and target.id == "code_snippets" for target in node.targets):
            continue
        if not isinstance(node.value, ast.Dict):
            continue
        for key, value in zip(node.value.keys, node.value.values):
            if isinstance(key, ast.Constant) and isinstance(value, ast.Constant):
                if isinstance(key.value, str) and isinstance(value.value, str):
                    snippets[key.value] = value.value
    return snippets


def extract_from_manifest(path: Path) -> tuple[dict[str, str], list[Path]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    snippets: dict[str, str] = {}
    fonts: list[Path] = []
    for item in data.get("code_figures", []):
        name = str(item.get("name", "code_figure"))
        text = str(item.get("source_text", ""))
        snippets[name] = text
    for raw in data.get("font_paths", []):
        fonts.append(Path(raw))
    return snippets, fonts


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate code figure text and font glyph coverage.")
    parser.add_argument("source", type=Path, help="code_figures_manifest.json or Python builder script")
    parser.add_argument("--font", action="append", default=[], help="Font path to include in glyph coverage")
    args = parser.parse_args()

    if args.source.suffix.lower() == ".json":
        snippets, manifest_fonts = extract_from_manifest(args.source)
    else:
        snippets = extract_from_python(args.source)
        manifest_fonts = []

    font_paths = [Path(value) for value in args.font] + manifest_fonts
    errors: list[str] = []
    if not snippets:
        errors.append("no code snippets found")
    if not font_paths:
        errors.append("no font paths provided")

    cmaps: list[tuple[Path, set[int]]] = []
    for font_path in font_paths:
        if not font_path.exists():
            errors.append(f"font not found: {font_path}")
            continue
        try:
            cmaps.append((font_path, load_font_cmap(font_path)))
        except Exception as exc:  # pragma: no cover - depends on local font files
            errors.append(f"font unreadable: {font_path}: {exc}")

    for name, text in snippets.items():
        for marker in BAD_MARKERS:
            if marker in text:
                errors.append(f"{name} source text contains mojibake marker {marker!r}")
        for char in sorted({c for c in text if ord(c) > 127}):
            if not any(ord(char) in cmap for _, cmap in cmaps):
                errors.append(f"{name} character {char!r} U+{ord(char):04X} is not covered by configured fonts")

    if errors:
        print("CODE_FIGURE_TEXT_CHECK=FAIL")
        for error in errors:
            print(f"- {error}")
        return 1

    print("CODE_FIGURE_TEXT_CHECK=PASS")
    print(f"CODE_FIGURES={len(snippets)}")
    print(f"FONTS={len(cmaps)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
