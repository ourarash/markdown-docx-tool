# Callout Styles

Use Pandoc fenced div syntax when you want a paragraph to map to a Word custom style:

```md
::: {custom-style="Risk"}
**Risk:** Add the callout text here.
This line becomes part of the same styled block in Word.
Use the bundled `reference.docx` when you want to change the visual treatment.
:::
```

## Bundled Styles

These styles currently exist in the bundled `scripts/reference.docx`:

- `Note`
- `Tip`
- `Important`
- `Warning`
- `Risk`
- `Caution`
- `InfoBox`
- `Decision`
- `Open Question`

## Important Detail

The `custom-style` value must match the Word style name exactly. For example:

- `Risk` works
- `Open Question` works
- `OpenQuestion` does not work

If a style is missing, add it to the bundled `reference.docx` instead of trying to approximate it with plain Markdown formatting.
