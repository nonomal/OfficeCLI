# Run / Character Formatting Showcase

Exercises the docx **run** (character-level) property surface. Three files:

- **run-showcase.sh** — builds the document with `officecli`.
- **run-showcase.docx** — generated output.
- **run-showcase.md** — this file.

## Regenerate

```bash
cd examples/word
bash run-showcase.sh
# → run-showcase.docx
```

## Families demonstrated

| Family | Properties |
|---|---|
| Weight & style | `bold`, `italic` |
| Underline | `underline=single\|double\|thick\|dotted\|wave`, `underline.color` |
| Strikethrough | `strike` (single), `dstrike` (double) |
| Case | `caps`, `smallcaps` |
| Vertical align | `superscript`, `subscript` (set on explicit `--type run` children) |
| Color / size / highlight | `color`, `size`, `highlight` |
| Per-script fonts | `font.latin`, `font.eastAsia` (Latin + CJK in one run) |
| Text effects | `emboss`, `imprint`, `outline`, `shadow` |
| Character spacing | `charSpacing`, `position` |
| Language | `lang`, `lang.latin`, `lang.ea`, `lang.cs` (BCP-47 per script) |
| Complex script | `bold.cs`, `italic.cs`, `size.cs`, `rtl`, `direction` |
| Theme fonts | `font.asciiTheme`, `font.hAnsiTheme`, `font.eaTheme`, `font.csTheme` |
| Per-script fonts | `font` (all), `font.latin`, `font.ea`, `font.cs` |
| Run shading / hidden | `shading`, `vanish`, `noproof` |
| w14 text effects | `textFill`, `textOutline`, `w14glow`, `w14reflection`, `w14shadow` |

This trio exercises **all 43 settable run properties** (verified by a coverage check).

## Mixed runs (super/subscript)

Most lines set run formatting on the paragraph's implicit run. For `E = mc²`
and `H₂O`, separate runs are appended so only part of the line is raised/lowered:

```bash
officecli add file.docx /body --type paragraph --prop text="E = mc"
officecli add file.docx "/body/p[last()]" --type run --prop text=2 --prop superscript=true
```

> The paragraph path `/body/p[last()]` must be quoted in the shell — `[` / `(`
> are shell metacharacters.

> `kern` is **not** a docx run property (only `charSpacing` is); use
> `charSpacing` for inter-character spacing.

> w14 text effects (`textFill`, `textOutline`, `w14*`) now apply via
> `add paragraph --prop ...` too (routed to the implicit run), matching how
> `bold`/`color` already behaved.
