#!/usr/bin/env python3
"""Validate ER diagram geometry manifests.

The expected manifest is emitted by thesis generation scripts next to the
rendered ER image. It contains entity/theme boxes, relation diamonds, connector
segments, ports, and optional cardinality label boxes.
"""

from __future__ import annotations

import argparse
import json
import math
from collections import Counter
from itertools import combinations
from pathlib import Path


Point = tuple[float, float]
Rect = tuple[float, float, float, float]
Segment = tuple[Point, Point]


def point(value) -> Point:
    return (float(value[0]), float(value[1]))


def rect(value) -> Rect:
    return (float(value[0]), float(value[1]), float(value[2]), float(value[3]))


def segment(value) -> Segment:
    return (point(value[0]), point(value[1]))


def segment_length(seg: Segment) -> float:
    (x1, y1), (x2, y2) = seg
    return math.hypot(x2 - x1, y2 - y1)


def rects_overlap(a: Rect, b: Rect, margin: float = 0.0) -> bool:
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    return not (
        ax2 <= bx1 + margin
        or ax1 >= bx2 - margin
        or ay2 <= by1 + margin
        or ay1 >= by2 - margin
    )


def segment_intersects_rect_interior(seg: Segment, box: Rect, margin: float = 2.0) -> bool:
    """Liang-Barsky clipping against the rectangle interior."""

    (x1, y1), (x2, y2) = seg
    xmin, ymin, xmax, ymax = box
    xmin += margin
    ymin += margin
    xmax -= margin
    ymax -= margin
    if xmin >= xmax or ymin >= ymax:
        return False

    dx = x2 - x1
    dy = y2 - y1
    p = [-dx, dx, -dy, dy]
    q = [x1 - xmin, xmax - x1, y1 - ymin, ymax - y1]
    u1 = 0.0
    u2 = 1.0
    for pi, qi in zip(p, q):
        if abs(pi) < 1e-9:
            if qi < 0:
                return False
            continue
        t = qi / pi
        if pi < 0:
            if t > u2:
                return False
            u1 = max(u1, t)
        else:
            if t < u1:
                return False
            u2 = min(u2, t)
    return u2 - u1 > 1e-6


def orientation(a: Point, b: Point, c: Point) -> int:
    value = (b[1] - a[1]) * (c[0] - b[0]) - (b[0] - a[0]) * (c[1] - b[1])
    if abs(value) < 1e-9:
        return 0
    return 1 if value > 0 else 2


def on_segment(a: Point, b: Point, c: Point) -> bool:
    return (
        min(a[0], c[0]) - 1e-9 <= b[0] <= max(a[0], c[0]) + 1e-9
        and min(a[1], c[1]) - 1e-9 <= b[1] <= max(a[1], c[1]) + 1e-9
    )


def segments_intersect(first: Segment, second: Segment) -> bool:
    p1, q1 = first
    p2, q2 = second
    o1 = orientation(p1, q1, p2)
    o2 = orientation(p1, q1, q2)
    o3 = orientation(p2, q2, p1)
    o4 = orientation(p2, q2, q1)
    if o1 != o2 and o3 != o4:
        return True
    if o1 == 0 and on_segment(p1, p2, q1):
        return True
    if o2 == 0 and on_segment(p1, q2, q1):
        return True
    if o3 == 0 and on_segment(p2, p1, q2):
        return True
    if o4 == 0 and on_segment(p2, q1, q2):
        return True
    return False


def diamond_box(relation: dict) -> Rect:
    if "diamond_box" in relation:
        return rect(relation["diamond_box"])
    cx, cy = point(relation["center"])
    return (cx - 60, cy - 36, cx + 60, cy + 36)


def collect_segments(relations: list[dict]) -> list[tuple[str, Segment]]:
    items: list[tuple[str, Segment]] = []
    for relation in relations:
        label = str(relation.get("label", "relation"))
        for raw in relation.get("segments", []):
            items.append((label, segment(raw)))
    return items


def validate(data: dict, min_segment_length: float = 24.0) -> list[str]:
    errors: list[str] = []
    boxes = {str(name): rect(value) for name, value in data.get("boxes", {}).items()}
    relations = list(data.get("relations", []))
    ports: list[tuple[float, float]] = []
    diamonds: list[tuple[str, Rect]] = []
    label_boxes: list[tuple[str, Rect]] = []

    for relation in relations:
        label = str(relation.get("label", "relation"))
        if "start" in relation:
            ports.append(point(relation["start"]))
        if "end" in relation:
            ports.append(point(relation["end"]))

        dbox = diamond_box(relation)
        diamonds.append((label, dbox))
        for box_name, box_rect in boxes.items():
            if rects_overlap(dbox, box_rect, margin=1.0):
                errors.append(f"relation diamond '{label}' overlaps box '{box_name}'")

        for raw in relation.get("segments", []):
            seg = segment(raw)
            if segment_length(seg) < min_segment_length:
                errors.append(f"relation '{label}' has too-short connector segment {raw}")
            for box_name, box_rect in boxes.items():
                if segment_intersects_rect_interior(seg, box_rect):
                    errors.append(f"relation '{label}' connector crosses box '{box_name}'")

        for item in relation.get("cardinality_boxes", []):
            lbox = rect(item.get("box", item))
            label_boxes.append((f"{label}:{item.get('text', '')}", lbox))
            for box_name, box_rect in boxes.items():
                if rects_overlap(lbox, box_rect, margin=1.0):
                    errors.append(f"cardinality label '{label}' overlaps box '{box_name}'")
            if rects_overlap(lbox, dbox, margin=1.0):
                errors.append(f"cardinality label '{label}' overlaps its relation diamond")

    for port, count in Counter(ports).items():
        if count > 1:
            errors.append(f"shared connection port {port}")

    all_segments = collect_segments(relations)
    for (first_label, first_seg), (second_label, second_seg) in combinations(all_segments, 2):
        if first_label == second_label:
            continue
        if segments_intersect(first_seg, second_seg):
            errors.append(f"connector for '{first_label}' intersects connector for '{second_label}'")

    for (first_label, first_box), (second_label, second_box) in combinations(label_boxes, 2):
        if rects_overlap(first_box, second_box, margin=1.0):
            errors.append(f"cardinality label '{first_label}' overlaps '{second_label}'")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate ER diagram layout geometry JSON.")
    parser.add_argument("layout_json", type=Path)
    parser.add_argument("--min-segment-length", type=float, default=24.0)
    args = parser.parse_args()

    data = json.loads(args.layout_json.read_text(encoding="utf-8"))
    errors = validate(data, min_segment_length=args.min_segment_length)
    if errors:
        print("ER_LAYOUT_GEOMETRY_CHECK=FAIL")
        for error in errors:
            print(f"- {error}")
        return 1

    print("ER_LAYOUT_GEOMETRY_CHECK=PASS")
    print(f"BOXES={len(data.get('boxes', {}))}")
    print(f"RELATIONS={len(data.get('relations', []))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
