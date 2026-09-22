#!/usr/bin/env python3
"""
Read-only audit of a Syndeio slide deck.

Never modifies the input file. Writes a JSON report of every issue found,
split into FIX (safe to apply automatically) and RECOMMEND (needs a human
decision).

Usage:
    python audit.py deck.pptx [--json out.json]
"""

import argparse
import collections
import json
import math
import os
import re
import sys
import zipfile

from pptx import Presentation

EMU = 914400.0

# ---------------------------------------------------------------- standard
SLIDE_W = 13.333
SLIDE_H = 7.5
SIZE_TOL = 0.05

FOOTER_LEFT, FOOTER_TOP = 7.97, 6.74
PAGENUM_LEFT, PAGENUM_TOP = 12.46, 6.74
POS_TOL = 0.05

FOOTER_PLAIN = "SYNDEIO BIOSCIENCES"
FOOTER_CONF = "SYNDEIO BIOSCIENCES - CONFIDENTIAL"

ALLOWED_FONTS = {"Aptos", "Aptos Display"}
INHERITED = {None, "", "+mn-lt", "+mj-lt"}

MIN_PT = 12.0

PALETTE = {
    "090446": "navy",
    "884CE0": "purple",
    "FC3702": "orange",
    "6A6A72": "grey",
    "FFFFFF": "white",
    "000000": "black",
}
NEAR_MISS_DIST = 55.0

DIVIDER_HINTS = ("transition", "divider", "section")


# ---------------------------------------------------------------- helpers
def inches(v):
    return None if v is None else v / EMU


def hex_to_rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def color_distance(a, b):
    ra, ga, ba = hex_to_rgb(a)
    rb, gb, bb = hex_to_rgb(b)
    return math.sqrt((ra - rb) ** 2 + (ga - gb) ** 2 + (ba - bb) ** 2)


def nearest_brand(h):
    best, bestd = None, 1e9
    for p in PALETTE:
        d = color_distance(h, p)
        if d < bestd:
            best, bestd = p, d
    return best, bestd


def slide_role(idx, total, layout_name):
    ln = (layout_name or "").lower()
    if idx == 1:
        return "title"
    if idx == total:
        return "back"
    if any(k in ln for k in DIVIDER_HINTS):
        return "divider"
    return "content"


def iter_runs(shape):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            yield para, run


def walk(shapes, depth=0):
    """Yield (shape, depth). Depth 0 shapes are positioned in slide
    coordinates; group children are not, so geometry checks must ignore
    anything below depth 0."""
    for sh in shapes:
        yield sh, depth
        if sh.shape_type == 6:  # GROUP
            try:
                for inner in walk(sh.shapes, depth + 1):
                    yield inner
            except Exception:
                pass


def in_chart_or_table(shape):
    try:
        return shape.has_chart or shape.has_table
    except Exception:
        return False


# ---------------------------------------------------------------- audit
def audit(path):
    prs = Presentation(path)
    total = len(prs.slides)
    W = inches(prs.slide_width)
    H = inches(prs.slide_height)

    fix, rec, info = [], [], {}

    # -- deck level ------------------------------------------------------
    info["slides"] = total
    info["slide_size"] = [round(W, 2), round(H, 2)]
    if abs(W - SLIDE_W) > SIZE_TOL or abs(H - SLIDE_H) > SIZE_TOL:
        rec.append({
            "kind": "slide_size",
            "slide": None,
            "detail": "Deck is %.2f x %.2f in; Syndeio standard is 13.33 x 7.5."
                      % (W, H),
            "action": "Resizing changes every layout. Rebuild from the "
                      "template instead.",
        })

    # metadata
    cp = prs.core_properties
    meta = {
        "author": cp.author or "",
        "last_modified_by": cp.last_modified_by or "",
        "title": cp.title or "",
    }
    info["metadata"] = meta
    leaked = [k for k, v in meta.items()
              if v and "syndeio" not in v.lower() and k != "title"]
    if leaked:
        rec.append({
            "kind": "metadata",
            "slide": None,
            "detail": "File properties carry: "
                      + "; ".join("%s = %s" % (k, meta[k]) for k in leaked),
            "action": "Clear author / last-modified-by before sending "
                      "outside Syndeio.",
        })

    # hidden slides + comments + notes (package level)
    hidden, commented = [], []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for i in range(1, total + 1):
            nm = "ppt/slides/slide%d.xml" % i
            if nm in names:
                xml = z.read(nm).decode("utf8", "ignore")
                if 'show="0"' in xml:
                    hidden.append(i)
        if any(n.startswith("ppt/comments/") for n in names):
            commented = [n for n in names if n.startswith("ppt/comments/")]

    if hidden:
        rec.append({
            "kind": "hidden_slides",
            "slide": None,
            "detail": "Hidden slides still travel with the file: %s"
                      % ", ".join(map(str, hidden)),
            "action": "Delete them if this deck leaves Syndeio.",
        })
    if commented:
        rec.append({
            "kind": "comments",
            "slide": None,
            "detail": "%d comment part(s) are still in the file."
                      % len(commented),
            "action": "Resolve or delete comments before sending.",
        })

    # -- pass 1: learn the deck's own dominant title position -------------
    title_pos = collections.Counter()
    for s in prs.slides:
        for sh in s.shapes:
            try:
                if sh.is_placeholder and "TITLE" in str(
                        sh.placeholder_format.type):
                    title_pos[(round(inches(sh.left), 2),
                               round(inches(sh.top), 2))] += 1
            except Exception:
                pass
    dom_title = title_pos.most_common(1)[0][0] if title_pos else None
    info["dominant_title_pos"] = list(dom_title) if dom_title else None

    # -- pass 2: per slide -----------------------------------------------
    notes_on = []
    off_palette = collections.Counter()
    small_text = collections.defaultdict(set)
    title_drift = []

    for idx, slide in enumerate(prs.slides, 1):
        role = slide_role(idx, total, slide.slide_layout.name)
        wants_chrome = role == "content"

        has_footer = False
        footer_text = None
        has_pagenum = False

        for sh, depth in walk(slide.shapes):
            try:
                l, t = inches(sh.left), inches(sh.top)
                w, h = inches(sh.width), inches(sh.height)
            except Exception:
                l = t = w = h = None

            # ---- footer / page number -------------------------------
            if sh.has_text_frame and t is not None and t > H - 1.0:
                txt = sh.text_frame.text.strip()
                if "SYNDEIO" in txt.upper():
                    has_footer = True
                    footer_text = txt
                    if (abs(l - FOOTER_LEFT) > POS_TOL
                            or abs(t - FOOTER_TOP) > POS_TOL):
                        fix.append({
                            "kind": "footer_pos",
                            "slide": idx,
                            "detail": "Footer at %.2f/%.2f, standard is "
                                      "%.2f/%.2f." % (l, t, FOOTER_LEFT,
                                                      FOOTER_TOP),
                            "action": "Move to the standard position.",
                        })

            if sh.has_text_frame:
                xml = sh.text_frame._txBody.xml
                if "slidenum" in xml:
                    has_pagenum = True
                    if l is not None and (
                            abs(l - PAGENUM_LEFT) > POS_TOL
                            or abs(t - PAGENUM_TOP) > POS_TOL):
                        fix.append({
                            "kind": "pagenum_pos",
                            "slide": idx,
                            "detail": "Page number at %.2f/%.2f, standard "
                                      "is %.2f/%.2f." % (l, t, PAGENUM_LEFT,
                                                         PAGENUM_TOP),
                            "action": "Move to the standard position.",
                        })

            # ---- overhang -------------------------------------------
            if depth == 0 and None not in (l, t, w, h) and w > 0 and h > 0:
                over = (l < -0.02 or t < -0.02
                        or l + w > W + 0.02 or t + h > H + 0.02)
                if over:
                    is_pic = sh.shape_type == 13
                    is_chart = in_chart_or_table(sh)
                    if is_pic:
                        rec.append({
                            "kind": "picture_bleed",
                            "slide": idx,
                            "detail": "Image extends past the slide edge "
                                      "(%.2f/%.2f, %.2f x %.2f)."
                                      % (l, t, w, h),
                            "action": "Left alone — this is usually a "
                                      "deliberate bleed. Say so if not.",
                        })
                    elif is_chart:
                        rec.append({
                            "kind": "chart_overhang",
                            "slide": idx,
                            "detail": "Chart or table runs off the slide "
                                      "(%.2f/%.2f, %.2f x %.2f)."
                                      % (l, t, w, h),
                            "action": "Can be resized to fit inside the "
                                      "margins on request.",
                        })
                    elif sh.has_text_frame:
                        fix.append({
                            "kind": "text_overhang",
                            "slide": idx,
                            "detail": "Text box runs off the slide "
                                      "(right edge at %.2f on a %.2f slide)."
                                      % (l + w, W),
                            "action": "Pull back inside the slide.",
                        })

            # ---- title drift ----------------------------------------
            try:
                if (dom_title and role == "content" and sh.is_placeholder
                        and "TITLE" in str(sh.placeholder_format.type)):
                    if (abs(l - dom_title[0]) > POS_TOL
                            or abs(t - dom_title[1]) > POS_TOL):
                        title_drift.append((idx, round(l, 2), round(t, 2)))
            except Exception:
                pass

            # ---- stretched image ------------------------------------
            if depth == 0 and sh.shape_type == 13 and None not in (w, h):
                try:
                    px_w, px_h = sh.image.size
                    keep_w = 1.0 - (sh.crop_left + sh.crop_right)
                    keep_h = 1.0 - (sh.crop_top + sh.crop_bottom)
                    native = (px_w * keep_w) / float(px_h * keep_h)
                    shown = w / float(h)
                    if native > 0 and abs(shown - native) / native > 0.10:
                        rec.append({
                            "kind": "stretched_image",
                            "slide": idx,
                            "detail": "Image aspect ratio is off by %.0f%% "
                                      "(native %.2f, shown %.2f)."
                                      % (abs(shown - native) / native * 100,
                                         native, shown),
                            "action": "Resize proportionally or re-crop.",
                        })
                except Exception:
                    pass

            # ---- runs: fonts, sizes, colors, whitespace -------------
            small_ok = in_chart_or_table(sh)
            for para, run in iter_runs(sh):
                txt = run.text
                if not txt.strip():
                    continue

                name = run.font.name
                if name not in INHERITED and name not in ALLOWED_FONTS:
                    fix.append({
                        "kind": "font",
                        "slide": idx,
                        "detail": "Font %s on \"%s\"" % (name, txt[:35]),
                        "action": "Change to Aptos.",
                    })

                sz = run.font.size.pt if run.font.size else None
                if sz is not None and sz < MIN_PT and not small_ok:
                    small_text[sz].add(idx)

                try:
                    if run.font.color and run.font.color.rgb:
                        h = str(run.font.color.rgb).upper()
                        if h not in PALETTE:
                            near, d = nearest_brand(h)
                            off_palette[(h, near, round(d))] += 1
                except Exception:
                    pass


            # ---- whitespace, at paragraph level ---------------------
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    ptxt = "".join(r.text for r in para.runs)
                    if not ptxt.strip():
                        continue
                    bad = ptxt != ptxt.rstrip()
                    if len(ptxt) > 25 and re.search(r"\w  +\w", ptxt):
                        bad = True
                    if bad:
                        fix.append({
                            "kind": "whitespace",
                            "slide": idx,
                            "detail": "Stray spacing in \"%s\"" % ptxt[:40],
                            "action": "Clean up spacing.",
                        })

            # ---- empty text box -------------------------------------
            if (sh.shape_type == 17 and sh.has_text_frame
                    and not sh.text_frame.text.strip()
                    and not sh.is_placeholder):
                fix.append({
                    "kind": "empty_box",
                    "slide": idx,
                    "detail": "Empty text box at %.2f/%.2f."
                              % (l or 0, t or 0),
                    "action": "Delete it.",
                })

        # ---- missing chrome ------------------------------------------
        if wants_chrome and not has_footer:
            fix.append({
                "kind": "footer_missing",
                "slide": idx,
                "detail": "No Syndeio footer.",
                "action": "Add the footer at the standard position.",
            })
        if wants_chrome and not has_pagenum:
            fix.append({
                "kind": "pagenum_missing",
                "slide": idx,
                "detail": "No page number.",
                "action": "Add the auto page-number field.",
            })
        if not wants_chrome and has_pagenum:
            fix.append({
                "kind": "pagenum_extra",
                "slide": idx,
                "detail": "%s slide has a page number." % role.title(),
                "action": "Remove it — %s slides are unnumbered." % role,
            })

        # ---- speaker notes -------------------------------------------
        try:
            if (slide.has_notes_slide
                    and slide.notes_slide.notes_text_frame.text.strip()):
                notes_on.append(idx)
        except Exception:
            pass

    if title_drift:
        fix.append({
            "kind": "title_drift",
            "slide": None,
            "detail": "Titles off the deck standard (%.2f/%.2f) on slides: %s"
                      % (dom_title[0], dom_title[1],
                         ", ".join(str(s) for s, _, _ in title_drift)),
            "action": "Snap each to the standard title position.",
            "targets": [s for s, _, _ in title_drift],
        })

    if small_text:
        smallest = min(small_text)
        allsl = sorted(set().union(*small_text.values()))
        rec.append({
            "kind": "small_text",
            "slide": None,
            "detail": "Text below 12pt on %d slide(s); smallest is %gpt. "
                      "Sizes used: %s"
                      % (len(allsl), smallest,
                         ", ".join("%gpt" % s for s in sorted(small_text))),
            "action": "Raise toward 12pt where it still fits. Slides: %s"
                      % ", ".join(map(str, allsl)),
        })

    if notes_on:
        rec.append({
            "kind": "speaker_notes",
            "slide": None,
            "detail": "Speaker notes on slides: %s"
                      % ", ".join(map(str, notes_on)),
            "action": "Review before sending externally. Not touched.",
        })

    # footer text consistency
    variants = collections.Counter()
    for slide in prs.slides:
        for sh, depth in walk(slide.shapes):
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if "SYNDEIO" in t.upper() and len(t) < 70:
                    try:
                        if inches(sh.top) > H - 1.0:
                            variants[t] += 1
                    except Exception:
                        pass
    info["footer_variants"] = dict(variants)
    if len(variants) > 1:
        fix.append({
            "kind": "footer_mixed",
            "slide": None,
            "detail": "Deck mixes footer texts: %s"
                      % "; ".join("%r x%d" % (k, v)
                                  for k, v in variants.items()),
            "action": "Standardise on the chosen variant.",
        })

    # off-palette summary
    for (h, near, d), n in off_palette.most_common():
        if d <= NEAR_MISS_DIST:
            rec.append({
                "kind": "color_near",
                "slide": None,
                "detail": "#%s used %d time(s) — very close to brand "
                          "%s (#%s)." % (h, n, PALETTE[near], near),
                "action": "Snap to #%s on request." % near,
            })
        else:
            rec.append({
                "kind": "color_off",
                "slide": None,
                "detail": "#%s used %d time(s) — not in the palette."
                          % (h, n),
                "action": "Replace with a brand color, or confirm it is "
                          "deliberate.",
            })

    return {"file": os.path.basename(path), "info": info,
            "fix": fix, "recommend": rec}


def summarise(report):
    by = collections.Counter(i["kind"] for i in report["fix"])
    print("FIX (%d)" % len(report["fix"]))
    for k, v in by.most_common():
        print("   %-18s %d" % (k, v))
    by = collections.Counter(i["kind"] for i in report["recommend"])
    print("RECOMMEND (%d)" % len(report["recommend"]))
    for k, v in by.most_common():
        print("   %-18s %d" % (k, v))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--json")
    a = ap.parse_args()
    rep = audit(a.deck)
    if a.json:
        with open(a.json, "w") as f:
            json.dump(rep, f, indent=2)
    summarise(rep)


if __name__ == "__main__":
    main()
