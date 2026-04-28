# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from PIL import Image


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

DEFAULT_CAPTION_RE = re.compile(r"^图\d+-\d+.*(架构|结构|模块|流程|E-R|ER|用例|部署|业务|接口|预测)")
DEFAULT_EXCLUDE_RE = re.compile(r"(界面|截图|结果|核心代码|代码)")


@dataclass
class DiagramImage:
    label: str
    name: str
    data: bytes


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def paragraph_blips(p: ET.Element) -> list[str]:
    return [blip.attrib.get(f"{{{NS['r']}}}embed", "") for blip in p.findall(".//a:blip", NS)]


def read_relationships(zf: ZipFile) -> dict[str, str]:
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    result: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid and target.startswith("media/"):
            result[rid] = "word/" + target
    return result


def extract_docx_diagrams(
    docx_path: Path,
    caption_re: re.Pattern[str],
    exclude_re: re.Pattern[str] | None,
) -> list[DiagramImage]:
    diagrams: list[DiagramImage] = []
    with ZipFile(docx_path) as zf:
        rels = read_relationships(zf)
        root = ET.fromstring(zf.read("word/document.xml"))
        last_rids: list[str] = []
        for child in root.findall(".//w:body/*", NS):
            if child.tag.endswith("}p"):
                blips = paragraph_blips(child)
                if blips:
                    last_rids = blips
                    continue

                text = paragraph_text(child)
                if not text or not caption_re.search(text):
                    continue
                if exclude_re and exclude_re.search(text):
                    continue
                if not last_rids:
                    diagrams.append(DiagramImage(text, "(missing-image-before-caption)", b""))
                    continue
                for rid in last_rids:
                    target = rels.get(rid)
                    if not target or target not in zf.namelist():
                        diagrams.append(DiagramImage(text, f"(missing-target:{rid})", b""))
                        continue
                    diagrams.append(DiagramImage(text, target, zf.read(target)))
                last_rids = []
    return diagrams


def extract_path_images(path: Path) -> list[DiagramImage]:
    suffixes = {".png", ".jpg", ".jpeg", ".bmp"}
    if path.is_dir():
        files = sorted(p for p in path.iterdir() if p.suffix.lower() in suffixes)
    else:
        files = [path]
    return [DiagramImage(p.stem, str(p), p.read_bytes()) for p in files]


def analyze_image(data: bytes, *, light_gray_threshold: float, chroma_threshold: float) -> tuple[bool, str]:
    if not data:
        return False, "image data missing"
    try:
        image = Image.open(BytesIO(data)).convert("RGB")
    except Exception as exc:  # pragma: no cover - CLI diagnostic
        return False, f"cannot open image: {exc}"

    total = image.width * image.height
    if total <= 0:
        return False, "empty image"

    white = 0
    dark = 0
    light_gray = 0
    chromatic = 0
    for r, g, b in image.getdata():
        mx = max(r, g, b)
        mn = min(r, g, b)
        spread = mx - mn
        if mn >= 250:
            white += 1
            continue
        if mx <= 45 and spread <= 18:
            dark += 1
            continue
        if spread > 18:
            chromatic += 1
            continue
        if 200 <= mn <= 249:
            light_gray += 1

    light_gray_ratio = light_gray / total
    chroma_ratio = chromatic / total
    white_ratio = white / total
    dark_ratio = dark / total
    failures: list[str] = []
    if light_gray_ratio > light_gray_threshold:
        failures.append(f"light_gray_ratio={light_gray_ratio:.4f}>{light_gray_threshold:.4f}")
    if chroma_ratio > chroma_threshold:
        failures.append(f"chroma_ratio={chroma_ratio:.4f}>{chroma_threshold:.4f}")
    if white_ratio < 0.45:
        failures.append(f"white_ratio={white_ratio:.4f}<0.4500")
    if dark_ratio <= 0:
        failures.append("no dark line/text pixels detected")
    detail = (
        f"size={image.width}x{image.height} "
        f"white={white_ratio:.4f} dark={dark_ratio:.4f} "
        f"light_gray={light_gray_ratio:.4f} chroma={chroma_ratio:.4f}"
    )
    if failures:
        return False, detail + " | " + "; ".join(failures)
    return True, detail


def compile_optional(pattern: str | None) -> re.Pattern[str] | None:
    return re.compile(pattern) if pattern else None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Verify that generated thesis diagrams use white background, black text, and black lines only."
    )
    parser.add_argument("path", help="DOCX file, image file, or directory of images")
    parser.add_argument("--caption-pattern", default=DEFAULT_CAPTION_RE.pattern)
    parser.add_argument("--exclude-caption-pattern", default=DEFAULT_EXCLUDE_RE.pattern)
    parser.add_argument("--light-gray-threshold", type=float, default=0.050)
    parser.add_argument("--chroma-threshold", type=float, default=0.002)
    args = parser.parse_args(argv)

    path = Path(args.path)
    caption_re = re.compile(args.caption_pattern)
    exclude_re = compile_optional(args.exclude_caption_pattern)
    if path.suffix.lower() == ".docx":
        diagrams = extract_docx_diagrams(path, caption_re, exclude_re)
    else:
        diagrams = extract_path_images(path)

    if not diagrams:
        print("MONOCHROME_DIAGRAM_CHECK=FAIL")
        print("- no matching structural/process/ER diagrams found")
        return 1

    failures: list[str] = []
    print(f"CHECKED_DIAGRAMS={len(diagrams)}")
    for diagram in diagrams:
        ok, detail = analyze_image(
            diagram.data,
            light_gray_threshold=args.light_gray_threshold,
            chroma_threshold=args.chroma_threshold,
        )
        status = "PASS" if ok else "FAIL"
        print(f"{status}: {diagram.label} [{diagram.name}] {detail}")
        if not ok:
            failures.append(diagram.label)

    if failures:
        print("MONOCHROME_DIAGRAM_CHECK=FAIL")
        for label in failures:
            print(f"- {label}")
        return 1

    print("MONOCHROME_DIAGRAM_CHECK=PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
