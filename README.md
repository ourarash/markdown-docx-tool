# Markdown to DOCX

[![CI](https://github.com/ourarash/markdown-docx-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/ourarash/markdown-docx-tool/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/ourarash/markdown-docx-tool/blob/main/LICENSE)

`markdown-to-docx` is a small Pandoc wrapper for turning Markdown files into Microsoft Word documents with a bundled `reference.docx`, cross-platform helper scripts, and one post-processing fix so wide tables auto-fit better in Word.

## Preview

![Preview of markdown-to-docx Word output](assets/showcase-preview.svg)

## Quick Start

Requirements:

- `pandoc` on every platform
- `perl`, `zip`, and `unzip` on macOS/Linux
- Python only if you want to use the installed skill launcher directly

Run a sample conversion from the repo root:

macOS/Linux:

```bash
./scripts/pandoc_md_to_docx.sh samples/showcase.md
```

Windows:

```powershell
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md
```

This writes `samples/showcase.docx`.

Useful options:

- `--toc` / `-TableOfContents` adds a table of contents
- `--metadata-file` / `-MetadataFile` passes Pandoc metadata
- `--reference-doc` / `-ReferenceDoc` overrides the bundled template
- `--output-dir` / `-OutputDir` keeps the default filename but writes elsewhere

Examples:

macOS/Linux:

```bash
./scripts/pandoc_md_to_docx.sh path/to/file.md
./scripts/pandoc_md_to_docx.sh --toc --output-dir output samples/showcase.md
```

Windows:

```powershell
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\path\to\file.md
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md -OutputDir .\output -TableOfContents
```

## Codex Skill

This repo also ships a self-contained Codex skill at `skills/markdown-to-docx`.

Install it from GitHub:

```bash
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py --repo ourarash/markdown-docx-tool --path skills/markdown-to-docx
```

Then use the bundled Python launcher:

macOS/Linux:

```bash
python3 ~/.codex/skills/markdown-to-docx/scripts/markdown_to_docx.py path/to/file.md
```

Windows:

```powershell
py -3 ~\.codex\skills\markdown-to-docx\scripts\markdown_to_docx.py .\path\to\file.md
```

Example prompts:

- `Convert ./samples/showcase.md to a Word document using the markdown-to-docx skill.`
- `Make a DOCX version of ./samples/meeting-notes.md and put it in ./output/meeting-notes.docx.`
- `Update the markdown-to-docx template so Note and Warning callouts stand out more in Word.`

## Template Customization

Edit `skills/markdown-to-docx/scripts/reference.docx` in Microsoft Word to change headings, paragraph spacing, code blocks, table styling, or fonts. Keep the filename the same so the scripts continue to find it automatically.

The sample files under `samples/` are good starting points for testing template changes:

- `samples/showcase.md` covers headings, lists, tables, code blocks, footnotes, and callout styles
- `samples/meeting-notes.md` is a smaller notes-style example

## Troubleshooting

- If `pandoc` is not found, install it and reopen your terminal.
- If the macOS/Linux shell script reports a missing command, install the missing tool and rerun.
- If PowerShell blocks script execution, run `Set-ExecutionPolicy -Scope Process Bypass` in that terminal window.
- If a Markdown file references images, keep the image paths relative to the Markdown file or use absolute paths.

## Development Checks

Validate the bundled Word template styles:

```bash
python3 scripts/validate_reference_docx.py --list --skill-doc skills/markdown-to-docx/SKILL.md --callout-doc skills/markdown-to-docx/references/callout-styles.md
```

Verify the skill after installing it into a temporary Codex-style directory:

```bash
python3 scripts/verify_installed_skill.py --output /tmp/installed-skill-showcase.docx
```

## License

MIT
