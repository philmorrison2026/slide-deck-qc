# slide-deck-qc

A Claude Code plugin marketplace hosting the `slide-deck-qc` skill: a
PowerPoint quality-control skill for Syndeio.

## What it does

The `slide-deck-qc` skill checks Syndeio slide decks (`.pptx`) for
quality issues before they go out — things like formatting
consistency, broken layouts, leftover placeholder text, and other
common review-worthy problems. (This is a scaffold; the actual QC
checks are still being written.)

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
