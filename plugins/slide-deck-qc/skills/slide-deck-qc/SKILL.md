---
name: slide-deck-qc
description: Cleans up the formatting of a Syndeio PowerPoint deck without changing any of its wording. Audits a .pptx against the Syndeio standard (footers, page numbers, fonts, title alignment, slide margins, palette), shows the user what it found, applies the safe fixes, and hands back a numbered list of recommendations the user can pick from. Use this whenever someone asks to clean up, tidy, polish, QC, quality-check, fix the formatting of, or "make consistent" a slide deck or PowerPoint, and whenever someone asks what is wrong with a deck or whether a deck is ready to send. Also use it when a deck is about to go to a board, a partner, or an investor and someone wants it checked first.
---

# Slide deck QC

Cleans up the formatting of a Syndeio deck. Never changes what it says.

## The one rule that matters

**Do not change the user's words.** Not a rephrase, not a shortened
bullet, not a fixed typo, not a capitalisation change. If a bullet is
clumsy, leave it clumsy. Formatting only. The moment this skill starts
editing content, people stop trusting it with their decks, and they are
right to.

Removing a trailing space is formatting. Rewriting a heading is not.

## Workflow

### 1. Ask which footer

Before anything else, ask one question:

> Plain footer (SYNDEIO BIOSCIENCES) or confidential
> (SYNDEIO BIOSCIENCES - CONFIDENTIAL)?

It affects every content slide, so it has to be settled first. Plain is
usual for capabilities and external decks; confidential for board and
internal material. If the deck already uses one consistently, say so and
offer to keep it.

If the user asked for an audit only, skip this — nothing is being written.

### 2. Audit

```bash
python scripts/audit.py deck.pptx --json /tmp/audit.json
```

Read the JSON. It returns two lists: `fix` (safe, mechanical) and
`recommend` (needs a human decision).

### 3. Show what was found, numbered

Print two numbered lists in the chat. Number them in one continuous
sequence so "skip 4" and "do 11" are never ambiguous.

Group the fixes by kind rather than printing one line per instance —
"Spacing cleaned on 21 paragraphs across slides 6, 7, 11…" not
twenty-one lines. Keep each recommendation to one line that names the
slide and what to do.

Then stop and ask. The user replies "go", or "go but skip 4 and 7".

**If the deck is clean, say so in a sentence and stop.** Do not pad the
list to look useful.

### 4. Apply

```bash
python scripts/apply_fixes.py deck.pptx -o "DeckName-QC.pptx" \
    --footer confidential --skip whitespace --log changes.md
```

- The original is never touched.
- Output is `<OriginalName>-QC.pptx` in the same folder.
- `--skip` takes a fix kind, repeatable, for anything the user waved off.

Then validate and confirm it is idempotent:

```bash
python scripts/office/validate.py DeckName-QC.pptx --original deck.pptx
python scripts/audit.py DeckName-QC.pptx
```

The second audit should report **0 fixes**. If it does not, a fix did not
apply and you should say so rather than claim success.

### 5. Hand back the recommendations

Present the recommendation list again with its numbers unchanged. The
user says "do 2, 5 and 9". Apply those to the **same** `-QC.pptx` file —
never create `-QC2`. Update the change log, then reshow whatever is left,
still with the original numbers.

Repeat as long as the user wants.

## Audit-only mode

If the user says "just tell me what's wrong" or "don't change anything",
run step 2, print step 3, and stop. No file is written, no footer
question is asked.

## What gets fixed automatically

| Kind | What it does |
|---|---|
| `footer_missing` | Adds the footer to a content slide that lacks one |
| `footer_pos` | Moves a footer to 7.97" / 6.74" |
| `footer_mixed` | Makes every footer use the chosen wording |
| `pagenum_missing` | Adds the auto page-number field at 12.46" / 6.74" |
| `pagenum_pos` | Moves a page number to the standard spot |
| `pagenum_extra` | Removes page numbers from title, divider and back slides |
| `font` | Off-brand fonts to Aptos |
| `title_drift` | Snaps titles to the deck's own dominant position |
| `text_overhang` | Pulls a text box back inside the slide edge |
| `empty_box` | Deletes genuinely empty text boxes |
| `whitespace` | Trailing spaces and doubled spaces between words |

## What only ever gets recommended

Colors, chart and table overhangs, text below 12pt, stretched images,
picture bleeds, speaker notes, hidden slides, comments, file metadata.

These need a judgment call. An off-palette color can be deliberate, a
bleed is usually design, and resizing a chart can break its labels. Offer
to do them, do not do them unasked.

## Traps worth knowing

**Never merge a paragraph's runs.** The obvious way to clean up spacing —
join the paragraph, fix the string, write it back to the first run — will
silently destroy superscripts, inline bold, and symbol runs. A real deck
lost its `™` this way in testing and it was only caught by rendering the
slide. Edit each run separately; strip trailing space only on a
paragraph's last run.

**Only top-level shapes have slide coordinates.** Shapes inside a group
are positioned in the group's own coordinate space, so checking them
against the slide edge produces nonsense — an early version reported five
overhangs on a slide that had none, and then "fixed" them. Geometry checks
and geometry fixes run at depth 0 only.

**Pictures that hang off the slide are usually meant to.** Full-bleed
cover art reads as an error to a script. Text boxes that overhang are
genuine mistakes. The audit already splits these — keep it that way.

**Blank-layout slides are ambiguous.** A slide on a "Blank" layout may be
a statement slide that should carry chrome, or a divider that should not.
If one turns up mid-deck without a footer, ask rather than guess.

**Slide size is not fixable.** A deck that is not 13.33 × 7.5 needs
rebuilding from the template. Say so; do not resize.

## Verify before declaring success

Always run `validate.py`, always re-run the audit, and render any slide
where something was added or resized:

```bash
python scripts/office/soffice.py --headless --convert-to pdf DeckName-QC.pptx
pdftoppm -jpeg -r 80 -f 31 -l 31 DeckName-QC.pdf slide
```

Look at the image. Compare against the same slide in the original if
anything was added. This is how the run-merging bug above was found.

## The standard

Measured from three reference decks. Full detail in
`references/syndeio-standard.md`.

- Slide size 13.33 × 7.5 in
- Footer at left 7.97", top 6.74"
- Page number field at left 12.46", top 6.74"
- No page number on the title slide, section dividers, or the back cover.
  Numbers are auto-fields, so the remaining slides stay correct on their own
- Aptos for body, Aptos Display for headings
- Palette: navy `090446`, purple `884CE0`, orange `FC3702`,
  grey `6A6A72`, white `FFFFFF`
- 12pt floor for body text; smaller is accepted inside charts and tables
