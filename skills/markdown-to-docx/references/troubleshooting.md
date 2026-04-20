# Troubleshooting

## Pandoc Not Found

If the script cannot find `pandoc`, install it first and rerun the conversion.

- macOS: `brew install pandoc`
- Windows: install Pandoc and make sure it is on `PATH`

## Mermaid Diagrams Do Not Render

Mermaid fenced blocks need Mermaid CLI so the wrapper can turn them into images before Pandoc writes the DOCX.

- Install it with `npm install -g @mermaid-js/mermaid-cli`
- If the executable is not on `PATH`, set `MERMAID_MMDC` to the full executable path
- Use `MARKDOWN_DOCX_MERMAID_SCALE=2` or higher if the PNG output looks soft in Word
- If diagrams look too large in Word, lower `MARKDOWN_DOCX_MERMAID_WIDTH` or set `width=...` on the individual Mermaid block
- If Chrome fails to launch in a sandboxed environment, set `MARKDOWN_DOCX_MERMAID_PUPPETEER_CONFIG` to a Puppeteer config file such as `scripts/puppeteer-no-sandbox.json`

## Missing Helper Tools On macOS/Linux

The shell script also expects:

- `perl`
- `zip`
- `unzip`

Install any missing tool and rerun the command.

## Shell Script Is Not Executable After Skill Install

Some installation paths may not preserve the shell script's executable bit.

Use the Python launcher instead:

- macOS/Linux: `python3 scripts/markdown_to_docx.py path/to/file.md`
- Windows: `py -3 .\scripts\markdown_to_docx.py .\path\to\file.md`

Fallback on macOS/Linux:

- `bash scripts/pandoc_md_to_docx.sh path/to/file.md`

## Callout Does Not Style Correctly

If a `custom-style` block converts but looks like normal body text in Word:

- confirm the style name matches Word exactly
- confirm the style exists in `scripts/reference.docx`
- edit the reference document if you need a new callout style

## Layout Looks Off In Word

Pandoc handles structure well, but it is still worth visually checking:

- wide tables
- large images
- long code blocks

The scripts already switch tables from fixed layout to auto-fit, but final visual review in Word is still recommended.
