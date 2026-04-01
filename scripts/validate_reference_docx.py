#!/usr/bin/env python3
"""Validate and inspect custom styles in a DOCX reference template."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
DEFAULT_REQUIRED_STYLES = [
    "Note",
    "Tip",
    "Important",
    "Warning",
    "Risk",
    "Caution",
    "InfoBox",
    "Decision",
    "Open Question",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="List or validate custom paragraph styles in a reference.docx file."
    )
    parser.add_argument(
        "--reference",
        default="skills/markdown-to-docx/scripts/reference.docx",
        help="Path to the reference.docx file to inspect.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Print the discovered custom paragraph styles.",
    )
    parser.add_argument(
        "--require",
        action="append",
        default=[],
        help="Require an additional custom paragraph style. Can be repeated.",
    )
    parser.add_argument(
        "--skill-doc",
        help="Optional SKILL.md path to compare against the documented bundled styles.",
    )
    parser.add_argument(
        "--callout-doc",
        help="Optional callout-styles.md path to compare against the documented bundled styles.",
    )
    return parser.parse_args()


def load_styles_xml(reference_doc: Path) -> ET.Element:
    with zipfile.ZipFile(reference_doc) as archive:
        with archive.open("word/styles.xml") as handle:
            return ET.parse(handle).getroot()


def collect_custom_paragraph_styles(root: ET.Element) -> list[str]:
    styles: list[str] = []
    for style in root.findall("w:style", WORD_NS):
        if style.get(f"{{{WORD_NS['w']}}}type") != "paragraph":
            continue
        if style.get(f"{{{WORD_NS['w']}}}customStyle") != "1":
            continue
        name = style.find("w:name", WORD_NS)
        if name is None:
            continue
        value = name.get(f"{{{WORD_NS['w']}}}val")
        if value:
            styles.append(value)
    return sorted(styles)


def extract_bullets_after_heading(doc_path: Path, heading: str) -> list[str]:
    text = doc_path.read_text(encoding="utf-8")
    marker = f"## {heading}"
    if marker not in text:
        raise ValueError(f"Heading not found in {doc_path}: {heading}")

    section = text.split(marker, 1)[1]
    lines = section.splitlines()[1:]
    bullets: list[str] = []
    started = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("## "):
            break
        if stripped.startswith("- `") and stripped.endswith("`"):
            bullets.append(stripped[3:-1])
            started = True
            continue
        if stripped.startswith("- "):
            bullets.append(stripped[2:].strip("`"))
            started = True
            continue
        if started and not stripped:
            break

    if not bullets:
        raise ValueError(f"No bullet list found under {heading!r} in {doc_path}")
    return bullets


def extract_skill_callouts(doc_path: Path) -> list[str]:
    text = doc_path.read_text(encoding="utf-8")
    marker = "Current bundled callout styles include:"
    if marker not in text:
        raise ValueError(f"Marker not found in {doc_path}: {marker}")
    section = text.split(marker, 1)[1]
    bullets = re.findall(r"^- `(.+?)`$", section, re.MULTILINE)
    if not bullets:
        raise ValueError(f"No callout bullets found in {doc_path}")
    return bullets


def compare_documented_styles(source_name: str, documented: list[str], actual: set[str]) -> list[str]:
    errors: list[str] = []
    missing = sorted(set(documented) - actual)
    if missing:
        errors.append(
            f"{source_name} documents styles missing from reference.docx: {', '.join(missing)}"
        )
    return errors


def main() -> int:
    args = parse_args()
    reference_doc = Path(args.reference)
    if not reference_doc.is_file():
        print(f"Error: reference DOCX not found: {reference_doc}", file=sys.stderr)
        return 1

    root = load_styles_xml(reference_doc)
    styles = collect_custom_paragraph_styles(root)
    style_set = set(styles)

    if args.list:
        print("Custom paragraph styles:")
        for style in styles:
            print(f"- {style}")

    required = sorted(set(DEFAULT_REQUIRED_STYLES + args.require))
    missing_required = [style for style in required if style not in style_set]
    if missing_required:
        print(
            "Missing required custom paragraph styles: "
            + ", ".join(missing_required),
            file=sys.stderr,
        )
        return 1

    errors: list[str] = []

    if args.skill_doc:
        documented = extract_skill_callouts(Path(args.skill_doc))
        errors.extend(compare_documented_styles("SKILL.md", documented, style_set))

    if args.callout_doc:
        documented = extract_bullets_after_heading(Path(args.callout_doc), "Bundled Styles")
        errors.extend(
            compare_documented_styles("callout-styles.md", documented, style_set)
        )

    if errors:
        for error in errors:
            print(error, file=sys.stderr)
        return 1

    print(
        f"Validated {reference_doc} with {len(styles)} custom paragraph styles."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
