# The Syndeio deck standard

Measured from three decks held up as good examples:

- `Syndeio_Biosciences_August_2026_Capabilities.pptx` (7 slides)
- `Syndeio_Biosciences_Presentation_April_2026_streamlined.pptx` (18 slides)
- `Syndeio_July_2026_Innoviva_Board_v18.pptx` (33 slides)

Everything below held across all three unless noted.

## Canvas

Slide size is **13.333 × 7.5 in** (16:9 widescreen) in all three decks.
A deck at any other size cannot be fixed by moving shapes — it needs
rebuilding from the template.

## Footer

Sits at **left 7.97", top 6.74"** with no variation at all across the
three decks. Two approved wordings:

| Wording | Used by |
|---|---|
| `SYNDEIO BIOSCIENCES` | Capabilities, April investor deck |
| `SYNDEIO BIOSCIENCES - CONFIDENTIAL` | Innoviva board deck |

A deck should use one or the other throughout, never both.

## Page numbers

At **left 12.46", top 6.74"** — same baseline as the footer, right side.

They are **auto-fields** (`<a:fld type="slidenum">`), not typed digits.
PowerPoint recalculates them, so inserting or reordering slides keeps
every number correct. All three decks had printed numbers matching actual
slide position.

Slides that carry no number:

- the title slide
- section dividers (layout names containing "Transition")
- the back cover

The board deck omits numbers on 1, 3, 19, 24, 31, 33. Slides 3, 19 and 24
are dividers; 1 and 33 are the covers. Slide 31 is on a "Blank" layout and
is genuinely ambiguous — a statement slide rather than a divider.

## Type

| Role | Font |
|---|---|
| Body | Aptos |
| Headings | Aptos Display |

Theme-inherited runs (`+mn-lt`, `+mj-lt`, or no explicit name) are correct
and should be left alone — they resolve to the theme fonts.

The board deck carried 5 runs of **Inter** on its back cover. That is the
only off-brand font found across 58 slides.

Sizes in use: 28, 24, 20, 18, 16, 14, 12, 11, 10.5, 10, 9, 8 pt. The
12pt floor is a target, not a hard rule — dense tables and chart labels
legitimately go smaller. The board deck has 5pt text, which is worth
raising.

## Title block

Titles sit at **left 0.68", top 0.28"** on most content slides, but drift
across a range of 0.60–0.69" left and 0.22–0.45" top. That drift is the
defect: snap to whichever position the deck itself uses most.

## Margins

Common left edges: 0.60–0.69" for titles and body, 0.79–0.80" for
secondary blocks. Common right margins: 0.44" and 0.96".

Four title placeholders in the board deck are 12.74" wide starting at
0.69", which puts the right edge at 13.43" on a 13.33" slide — a 0.1"
overhang on slides 15, 21, 22 and 23.

## Palette

| Hex | Name |
|---|---|
| `090446` | navy (primary) |
| `884CE0` | purple |
| `FC3702` | orange (accent) |
| `6A6A72` | grey |
| `FFFFFF` | white |
| `000000` | black |

Near-misses found in real decks, all of which should be consolidated:

- navies: `0D1B4B`, `151B47`, `241B45`, `23232E`
- greys: `2D3748`, `3C4756`, `535353`, `393B39`, `7C7E8E`
- reds and oranges: `F71735`, `C00000`, `E80142`, `C70E28`
- greens: `519A6B`, `00B050`, `7EE69C`
- purple: `6F5BC4`

The green is not in the palette at all. If greens are a real part of the
Syndeio scheme, one should be picked and added here.

## Known defects in the reference decks

Useful as test cases:

| Deck | Issue |
|---|---|
| Board | Slide 31 has no footer or page number |
| Board | Titles overhang the right edge on 15, 21, 22, 23 |
| Board | Back cover uses Inter |
| Board | Chart runs off the bottom of slide 8 |
| Board | 5pt text |
| Board | Footer on slide 30 is too wide for its box and wraps |
| April | 5 text boxes overhang; one page number on a slide that should not have one |
| All | Author and last-modified-by are personal names, not Syndeio |
