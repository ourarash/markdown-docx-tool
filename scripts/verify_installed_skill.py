#!/usr/bin/env python3
"""Verify that the markdown-to-docx skill works after being installed to a Codex-style path."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Install the local skill into a temp Codex-style directory and run a conversion."
    )
    parser.add_argument(
        "--skill-src",
        default="skills/markdown-to-docx",
        help="Path to the source skill directory.",
    )
    parser.add_argument(
        "--input",
        default="samples/showcase.md",
        help="Markdown file to convert during verification.",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output DOCX path that should be produced by the installed skill.",
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help="Python executable to use for invoking the installed launcher.",
    )
    parser.add_argument(
        "--toc",
        action="store_true",
        help="Pass --toc through the installed skill launcher.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    skill_src = Path(args.skill_src).resolve()
    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()

    if not skill_src.is_dir():
        print(f"Error: skill source directory not found: {skill_src}", file=sys.stderr)
        return 1
    if not input_path.is_file():
        print(f"Error: input markdown file not found: {input_path}", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory(prefix="codex-skill-verify-") as tmp_dir:
        codex_home = Path(tmp_dir) / "codex-home"
        install_dir = codex_home / "skills" / "markdown-to-docx"
        shutil.copytree(skill_src, install_dir)

        expected_files = [
            install_dir / "SKILL.md",
            install_dir / "scripts" / "markdown_to_docx.py",
            install_dir / "scripts" / "reference.docx",
        ]
        missing = [str(path) for path in expected_files if not path.exists()]
        if missing:
            print(
                "Error: installed skill is missing expected files: "
                + ", ".join(missing),
                file=sys.stderr,
            )
            return 1

        output_path.parent.mkdir(parents=True, exist_ok=True)
        launcher = install_dir / "scripts" / "markdown_to_docx.py"
        cmd = [args.python, str(launcher)]
        if args.toc:
            cmd.append("--toc")
        cmd.extend([str(input_path), str(output_path)])

        completed = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if completed.returncode != 0:
            sys.stdout.write(completed.stdout)
            sys.stderr.write(completed.stderr)
            return completed.returncode

        if not output_path.is_file():
            print(f"Error: installed skill did not produce {output_path}", file=sys.stderr)
            return 1

        print(f"Installed skill verification passed: {output_path}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
