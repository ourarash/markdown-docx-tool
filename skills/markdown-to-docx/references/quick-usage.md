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
- Mermaid fenced code blocks are rendered to images when `mmdc` is available.
- Mermaid diagrams are centered by default in the generated DOCX.
- Mermaid diagrams get automatic captions in the generated DOCX.
- Mermaid diagrams default to `5.5in` wide unless a block-level `width` attribute overrides that.
- The scripts apply DOCX XML cleanup after Pandoc runs so Word tables auto-fit better, paragraphs after tables have clearer spacing, and generated heading bookmarks are removed.
- The Python launcher avoids relying on the shell script's executable bit after skill installation.
