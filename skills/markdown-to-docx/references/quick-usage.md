# Quick Usage

## Core Commands

macOS/Linux:

```bash
./scripts/pandoc_md_to_docx.sh path/to/file.md
./scripts/pandoc_md_to_docx.sh path/to/file.md path/to/output.docx
```

Windows:

```powershell
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\path\to\file.md
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\path\to\file.md -OutputPath .\path\to\output.docx
```

## Behavior

- The input file can be absolute, repo-relative, or relative to the current working directory.
- If output is omitted, the script writes a `.docx` file next to the Markdown file.
- Relative images are resolved from the input file's directory.
- Styling comes from the bundled `scripts/reference.docx`.
- The scripts apply one DOCX XML fix after Pandoc runs so Word tables auto-fit better.
