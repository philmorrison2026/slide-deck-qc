# slide-deck-qc

A Claude Code plugin marketplace hosting the `slide-deck-qc` skill: a
PowerPoint quality-control skill for Syndeio.

## What it does

The `slide-deck-qc` skill cleans up the formatting of a Syndeio deck.
It never changes any wording — no rephrasing, no shortened bullets, no
typo fixes. Formatting only.

It asks whether the deck should use the plain or CONFIDENTIAL footer,
audits the deck against the Syndeio standard, shows you what it found,
and applies the safe fixes on your go-ahead. The output is
`DeckName-QC.pptx` next to the untouched original, along with a change
log and a numbered list of recommendations you can pick from for
anything that needs a judgment call.

It checks: footers, page numbers, fonts, title alignment, slide
margins, palette, image aspect ratios, speaker notes, hidden slides,
and file metadata.

Requires [python-pptx](https://python-pptx.readthedocs.io/).

## Installing (for Syndeio colleagues)

Most colleagues will use the **Claude app** route below.

### Claude app (claude.ai, desktop, mobile)

1. Go to **Customize → Plugins**.
2. Click **+ → Add marketplace → Add from a repository**.
3. Paste this repo's URL: `https://github.com/philmorrison2026/slide-deck-qc`
4. Install the **slide-deck-qc** plugin.

> **Before installing:** code execution must be enabled first, under
> **Settings → Capabilities**.

### Claude Code

```
/plugin marketplace add philmorrison2026/slide-deck-qc
/plugin install slide-deck-qc@syndeio
```
