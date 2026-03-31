# Markdown to DOCX

`markdown-to-docx` is a small wrapper repo around Pandoc for turning Markdown files into Microsoft Word documents.

It does not replace Pandoc. Instead, it packages a reusable Word style template, cross-platform helper scripts, and one small post-processing fix for Word table layout.

## Run The Showcase Example

If you want the fastest way to see the repo working, run the included showcase sample from the repo root:

macOS/Linux:

```bash
./scripts/pandoc_md_to_docx.sh samples/showcase.md
```

Windows:

```powershell
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md
```

This writes `samples/showcase.docx`.

## What This Adds

- `📝` a reusable Word reference template in `scripts/reference.docx`
- `🖥️` helper scripts for macOS/Linux and Windows
- `📎` predictable handling for relative input paths and linked assets
- `📊` one DOCX XML fix so wide tables auto-fit better in Word

If you already prefer running Pandoc directly and do not need these defaults, you may not need this repo.

## What's Included

It ships with:

- a macOS/Linux shell script: `scripts/pandoc_md_to_docx.sh`
- a Windows PowerShell script: `scripts/pandoc_md_to_docx.ps1`
- a Word reference template: `scripts/reference.docx`
- sample Markdown files under `samples/`

The conversion flow stays intentionally close to Pandoc and keeps one post-processing step that switches Word tables from fixed layout to auto-fit, which helps wide Markdown tables render more naturally.

## Repo Layout

```text
markdown-to-docx/
├── README.md
├── scripts/
│   ├── pandoc_md_to_docx.ps1
│   ├── pandoc_md_to_docx.sh
│   └── reference.docx
└── samples/
    ├── meeting-notes.md
    └── showcase.md
```

## Install On macOS

1. Install Pandoc.

   ```bash
   brew install pandoc
   ```

2. Make sure the helper tools are available.

   macOS usually already includes `bash`, `perl`, `zip`, and `unzip`. You can verify with:

   ```bash
   bash --version
   perl -v
   zip -v
   unzip -v
   ```

3. Make the shell script executable.

   ```bash
   chmod +x scripts/pandoc_md_to_docx.sh
   ```

4. Run a sample conversion from the repo root.

   ```bash
   ./scripts/pandoc_md_to_docx.sh samples/showcase.md
   ```

   That writes `samples/showcase.docx`.

5. Optional: write the output somewhere else.

   ```bash
   ./scripts/pandoc_md_to_docx.sh samples/showcase.md output/showcase.docx
   ```

## Install On Windows

1. Install Pandoc and confirm it is on your `PATH`.

   In PowerShell:

   ```powershell
   pandoc --version
   ```

2. Open PowerShell in the repo root.

3. If your system blocks local scripts, allow this session to run them.

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   ```

4. Run a sample conversion.

   ```powershell
   .\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md
   ```

   That writes `samples\showcase.docx`.

5. Optional: write the output somewhere else.

   ```powershell
   .\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md -OutputPath .\output\showcase.docx
   ```

## How To Use It

From the repo root, point either script at a Markdown file:

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

Behavior:

- The input file can be absolute, repo-relative, or relative to your current working directory.
- If you do not pass an output path, the script writes a `.docx` file next to the Markdown file.
- Relative images are resolved from the input file's directory.
- Styling comes from `scripts/reference.docx`.
- The scripts apply one DOCX XML fix after Pandoc runs so Word tables auto-fit better.

## Customizing The Word Style

Edit `scripts/reference.docx` in Microsoft Word when you want to change:

- heading styles
- normal paragraph spacing
- code block appearance
- table styling
- fonts

Keep the filename the same so the scripts can continue to find it automatically.

## Sample Markdown

Use the included files to see what the converter handles well:

- `samples/showcase.md` demonstrates headings, emphasis, lists, tables, quotes, footnotes, fenced code blocks, and `custom-style` callout sections such as `Note`, `Tip`, `Important`, `Warning`, `Risk`, and `Caution`.
- `samples/meeting-notes.md` is a smaller real-world example for notes and action items.

Try them directly:

macOS/Linux:

```bash
./scripts/pandoc_md_to_docx.sh samples/showcase.md
./scripts/pandoc_md_to_docx.sh samples/meeting-notes.md output/meeting-notes.docx
```

Windows:

```powershell
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\showcase.md
.\scripts\pandoc_md_to_docx.ps1 -InputPath .\samples\meeting-notes.md -OutputPath .\output\meeting-notes.docx
```

## Troubleshooting

- If `pandoc` is not found, install it first and reopen your terminal.
- If the shell script says a command is missing on macOS, install the missing tool and rerun the conversion.
- If PowerShell blocks script execution, rerun `Set-ExecutionPolicy -Scope Process Bypass` in that terminal window.
- If a Markdown file references images, keep the image paths relative to the Markdown file or use absolute paths.
