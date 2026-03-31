# Markdown to Word Showcase

This sample demonstrates the kinds of Markdown features that convert cleanly into a Word document with the included scripts and `reference.docx`.

## Highlights

- Bold text for emphasis
- Italic text for nuance
- Inline code like `pandoc`
- [Links](https://pandoc.org)
- Footnotes for supporting detail[^1]
- Custom callout blocks with Word styles

## Status Table

| Deliverable | Owner | Status | Notes |
| --- | --- | --- | --- |
| Draft outline | Maya | Done | Shared with the team |
| Technical review | Omar | In progress | Waiting on GPU metrics |
| Final polish | Priya | Planned | Add screenshots and appendix |

## Ordered Steps

1. Write the Markdown draft.
2. Convert it to DOCX.
3. Review the generated document in Word.
4. Adjust `scripts/reference.docx` if you want different styling.

## Quote

> The fastest way to improve a repeatable document workflow is to make the template and the conversion command boring and reliable.

## Callout Styles

::: {custom-style="Note"}
**Note:** This showcase includes a few callout blocks that use Pandoc's `custom-style` attribute.
Define matching styles in `scripts/reference.docx` if you want them to stand out visually in Word.
Without those styles, the content still converts, but the callout may look like normal body text.
:::

---

::: {custom-style="Tip"}
**Tip:** Keep each callout short enough to scan quickly in Word after conversion.
A bold lead-in works well when you want the first line to communicate the purpose immediately.
This pattern is useful for onboarding docs, handoff notes, and review checklists.
:::

---

::: {custom-style="Important"}
**Important:** Custom styles are only as good as the reference document behind them.
If you rename a style in Word, update the Markdown examples so the `custom-style` value still matches.
That keeps the generated DOCX consistent across machines and team members.
:::

---

::: {custom-style="Warning"}
**Warning:** Wide tables, images, and long code blocks can still need a quick visual review in Word.
Pandoc handles most structure well, but layout-sensitive content is always worth checking after export.
Treat the generated DOCX as reliable output, not as something you should never proofread.
:::

---

::: {custom-style="Risk"}
**Risk:** A custom callout style that exists only on one person's machine can create confusing results for everyone else.
Store the shared look in `scripts/reference.docx` so the repo stays the single source of truth.
That reduces drift between drafts and avoids last-minute formatting cleanup.
:::

---

::: {custom-style="Caution"}
**Caution:** If you copy callouts from another Markdown flavor, verify that the syntax still works with Pandoc.
Some tools support admonitions differently, and not every fenced block format maps to DOCX the same way.
Using one consistent pattern in this repo makes the conversion workflow easier to trust.
:::

## Code Block

```bash
./scripts/pandoc_md_to_docx.sh samples/showcase.md output/showcase.docx
```

## Action Checklist

- [x] Include headings
- [x] Include a table
- [x] Include a quote
- [x] Include code
- [ ] Add your own project-specific template styles

## Small Appendix

Here is a short paragraph that gives Word a little more structure to work with. Multi-paragraph sections, lists, and tables usually look best once the reference document has the fonts and spacing you want.

[^1]: Pandoc carries footnotes into the generated Word document.
