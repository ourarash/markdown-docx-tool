# Troubleshooting

## Pandoc Not Found

If the script cannot find `pandoc`, install it first and rerun the conversion.

- macOS: `brew install pandoc`
- Windows: install Pandoc and make sure it is on `PATH`

## Missing Helper Tools On macOS/Linux

The shell script also expects:

- `perl`
- `zip`
- `unzip`

Install any missing tool and rerun the command.

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
