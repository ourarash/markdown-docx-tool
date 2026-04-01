# Quick Usage

## Core Commands

macOS/Linux:

```bash
python3 scripts/markdown_to_docx.py path/to/file.md
python3 scripts/markdown_to_docx.py path/to/file.md path/to/output.docx
```

Windows:

```powershell
py -3 .\scripts\markdown_to_docx.py .\path\to\file.md
py -3 .\scripts\markdown_to_docx.py .\path\to\file.md .\path\to\output.docx
```

Fallback on macOS/Linux if the Python launcher is not appropriate:

```bash
bash scripts/pandoc_md_to_docx.sh path/to/file.md
```

## Behavior

- The input file can be absolute, repo-relative, or relative to the current working directory.
- If output is omitted, the script writes a `.docx` file next to the Markdown file.
- Relative images are resolved from the input file's directory.
- Styling comes from the bundled `scripts/reference.docx`.
- The scripts apply one DOCX XML fix after Pandoc runs so Word tables auto-fit better.
- The Python launcher avoids relying on the shell script's executable bit after skill installation.
