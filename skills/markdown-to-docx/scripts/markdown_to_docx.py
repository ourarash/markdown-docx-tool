#!/usr/bin/env python3
"""Cross-platform launcher for the bundled Markdown-to-DOCX converters."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Markdown to DOCX with the bundled markdown-to-docx skill."
    )
    parser.add_argument(
        "--reference-doc",
        help="Optional custom reference DOCX to use instead of the bundled template.",
    )
    parser.add_argument(
        "--toc",
        action="store_true",
        help="Include a table of contents in the generated DOCX.",
    )
    parser.add_argument(
        "--metadata-file",
        help="Optional Pandoc metadata file to include during conversion.",
    )
    parser.add_argument(
        "--output-dir",
        help="Optional output directory for the generated DOCX when output_path is omitted.",
    )
    parser.add_argument("input_path", help="Path to the source Markdown file")
    parser.add_argument(
        "output_path",
        nargs="?",
        help="Optional output DOCX path. Defaults next to the input file.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent

    if sys.platform.startswith("win"):
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if not powershell:
            print(
                "Error: PowerShell was not found. Install PowerShell or run the bundled "
                "pandoc_md_to_docx.ps1 manually.",
                file=sys.stderr,
            )
            return 1

        cmd = [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_dir / "pandoc_md_to_docx.ps1"),
            "-InputPath",
            args.input_path,
        ]
        if args.reference_doc:
            cmd.extend(["-ReferenceDoc", args.reference_doc])
        if args.toc:
            cmd.append("-TableOfContents")
        if args.metadata_file:
            cmd.extend(["-MetadataFile", args.metadata_file])
        if args.output_dir:
            cmd.extend(["-OutputDir", args.output_dir])
        if args.output_path:
            cmd.extend(["-OutputPath", args.output_path])
    else:
        bash = shutil.which("bash")
        if not bash:
            print(
                "Error: bash was not found. Install bash or run the bundled shell script manually.",
                file=sys.stderr,
            )
            return 1

        cmd = [bash, str(script_dir / "pandoc_md_to_docx.sh")]
        if args.reference_doc:
            cmd.extend(["--reference-doc", args.reference_doc])
        if args.toc:
            cmd.append("--toc")
        if args.metadata_file:
            cmd.extend(["--metadata-file", args.metadata_file])
        if args.output_dir:
            cmd.extend(["--output-dir", args.output_dir])
        cmd.append(args.input_path)
        if args.output_path:
            cmd.append(args.output_path)

    completed = subprocess.run(cmd)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
