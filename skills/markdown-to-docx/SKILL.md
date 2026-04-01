---
name: markdown-to-docx
description: Use when the user wants to convert Markdown to DOCX, make a Microsoft Word version of a Markdown file, style Word output with a reference template, fix DOCX callouts, or customize the bundled Word template for notes, reports, and review docs.
---

# Markdown to DOCX

Converts Markdown files into Microsoft Word documents with Pandoc, a bundled Word reference template, and one post-processing fix that makes wide tables auto-fit more naturally in Word.

This skill supports two kinds of work:

- convert Markdown to DOCX with the bundled scripts
- customize Word styling and `custom-style` callouts in the bundled `reference.docx`

## Quick Start

Use the bundled launcher from this skill directory so installation does not depend on shell execute bits:

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

Fallback on macOS/Linux if needed:

```bash
bash scripts/pandoc_md_to_docx.sh path/to/file.md
```

## Workflow

1. Detect the operating system and choose:
   - `scripts/markdown_to_docx.py` as the preferred launcher
   - `scripts/pandoc_md_to_docx.sh` on macOS/Linux when direct script invocation is needed
   - `scripts/pandoc_md_to_docx.ps1` on Windows when direct PowerShell invocation is needed
2. Verify that `pandoc` is installed before attempting conversion.
3. If the user did not provide an output path, let the script write `<input>.docx` next to the source Markdown file.
4. Use the bundled `scripts/reference.docx` as the Word template.
5. If the user wants callouts or special formatting, confirm that the corresponding custom style exists in the template.
6. If the user wants the callout look or typography changed, edit `scripts/reference.docx` rather than trying to fake the styling in Markdown.

## Callout Guidance

Pandoc `custom-style` blocks only render with special Word styling when the style exists in `scripts/reference.docx`.

Current bundled callout styles include:

- `Note`
- `Tip`
- `Important`
- `Warning`
- `Risk`
- `Caution`
- `InfoBox`
- `Decision`
- `Open Question`

When the user asks to add or fix callouts:

- keep the Markdown block in the form `::: {custom-style="Style Name"}`
- make sure the style name matches Word exactly
- edit the bundled `reference.docx` if the style is missing or needs different colors, borders, or spacing

## Dependencies

- `pandoc` is required on every platform
- macOS/Linux also need `perl`, `zip`, and `unzip`
- Windows relies on built-in PowerShell archive commands

## References

Open only what you need:

- `references/quick-usage.md` for direct conversion commands and expected behavior
- `references/callout-styles.md` for style names and Markdown examples
- `references/troubleshooting.md` for missing-tool and template issues
