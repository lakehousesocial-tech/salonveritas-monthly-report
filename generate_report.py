#!/usr/bin/env python3
"""
Generates the monthly marketing report .pptx.

Design system: Salon Veritas brand (per the client's Brand Kit -- sage green,
dusty blush, warm brass, ivory/cream, espresso walnut).
  - Ivory/Cream (#F5F1EA) background on every slide; Pure White card/table fills
  - Gill Sans MT throughout (titles bold, body regular) -- the brand kit reserves
    its hand-lettered signature script for the logotype/hero headlines/signage
    only, never body copy or small-size UI, so it is not used in this deck
  - Round bullet (●) lists for body text
  - Layout grid: title at (0.56in, 0.58in), content area starting at (0.56in, 1.24in)
  - Recurring 1px Rule-colored line under every slide title (signature framing
    motif); Dusty Blush reserved for sparing emphasis -- currently only the cover
    slide's accent rule, month chip, and each Insights slide's headline stat

  ink_mid/ink_light/rule below are tints of the brand kit's Espresso Walnut/Sage
  Green (blended toward Ivory/Cream) for text-hierarchy and divider use -- not
  literal brand-kit swatches, since the kit only specifies five hues and a UI
  needs more steps than that for secondary/tertiary text.

  NOTE: fonts are set by name only (run.font.name), not embedded as binary
  font data in the .pptx -- python-pptx has no support for OOXML font
  embedding. Gill Sans MT ships with Windows/Office and macOS; install it
  locally on any machine used to review/export the deck so PowerPoint/Keynote
  render it instead of silently substituting a fallback font.

Data-driven slides (Instagram/Facebook/TikTok) are built from a Buffer metrics
JSON file shaped like the (verified, real) output of buffer-metrics.js:

  {
    "channels": {
      "instagram": [ {"channelId", "name",
          "current": {"range": {...}, "posts": [ {postId, text, sentAt,
              metrics: {type: value, ...}}, ... ]},
          "prior":   {"range": {...}, "posts": [ ... same shape ... ]} } ],
      "facebook": [ ...same shape... ],
      "tiktok": [ ...same shape... ]
    }
  }

A channel entry may instead carry an "error" string if that channel's fetch
failed -- such entries are treated as having no data. Both periods are
derived identically (summed from each post's own `metrics`), never from
Buffer's `aggregatedPostMetrics` field -- that field is capped by Buffer's
Free plan to roughly the last 30 days (confirmed live: querying further back
errors with "Free-plan Insights are limited to the last 31 days of
history"), which would make it unusable for any prior period here anyway.

Current period is always the last 30 days. Prior period width is PER
CHANNEL (locked in after live diagnostics on 2026-07-18): Instagram uses a
60-day-wide prior window (a straight 30-day window only had 1 post -- too
thin to compare against; 60 days gave 10, comparable to the current
period's 7). TikTok and Facebook use a 30-day-wide prior window since
posting volume there was already comparable between periods.

Metric field naming is NOT uniform across platforms: Facebook reports
"impressions", Instagram/TikTok report "views"; only Instagram posts ever
carry "saves". None of the three ever carry a follower-count field via
Buffer's API -- confirmed via full schema introspection of Channel and
every per-service ChannelMetadata type.

Follower counts come from a SEPARATE source: fetch_followers.js queries the
Meta Graph API directly (GET /{FB_PAGE_ID}?fields=followers_count,
instagram_business_account{followers_count}) and appends a dated snapshot
to follower-history.json in this same folder. That log is read here via
--follower-log (default: follower-history.json next to this script) to add
a "Current Followers" hero card to the Instagram/Facebook Insights slides
-- TikTok has no follower data source at all, from Buffer or Meta, so it
never gets this card. Meta's API only exposes a current snapshot, not
history, so month-over-month growth only appears once a second monthly
snapshot exists; the first month shows the count alone with no invented
delta.

Misc Marketing Items / WINS slides are still filled in by hand each month --
generated here as templated placeholder slides, not from data.

Goals slides are now auto-generated from client-context.md (see
--client-context, default: client-context.md next to this script) plus this
run's real metrics:
  - "Last Month" goals: re-reads the most recent goal entries logged in
    client-context.md's Goal Log and checks each one's embedded
    "<!-- check: platform.metric op baseline -->" condition against this
    period's actual totals, flipping it to done (✅) only if the data
    confirms it -- never assumed. A goal with no check condition (e.g. the
    Milbon Gold Line content tie-in) stays unchecked until marked by hand.
    First run ever (empty log): falls back to the original manual
    placeholder bullets, since there's no prior automated goal to check.
  - "Next Month" goals: 3-5 new goals generated from the most actionable
    finding per platform (posting-frequency drop > content-concentration
    risk > engagement-rate decline > "maintain," in that priority order),
    plus a cross-platform engagement-gap goal when one platform trails the
    others by a meaningful margin, plus a static goal tying content to the
    Milbon Gold Line package spotlight from client-context.md.
  - After generating, this run's goals (and a short "what changed" summary)
    are appended to client-context.md's Goal Log for next month's checkoff,
    and printed to stdout separately from the .pptx so they can be reviewed
    before the deck goes anywhere.
"""

import argparse
import json
import os
import re
from datetime import datetime, timedelta, timezone

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# Design system
# --------------------------------------------------------------------------

COLORS = {
    "background": RGBColor(0xF5, 0xF1, 0xEA),   # Ivory/Cream -- slide background
    "cobalt_deep": RGBColor(0x70, 0x72, 0x52),  # Sage Green -- headers / primary accent
    "cobalt_mid": RGBColor(0xB0, 0x8D, 0x57),   # Warm Brass -- chart lines / active data elements
    "gilded": RGBColor(0xD9, 0x9F, 0x95),       # Dusty Blush -- sparing, emphasis only
    "ink": RGBColor(0x2B, 0x26, 0x20),          # Espresso Walnut -- bullets, table/data-row copy
    "ink_mid": RGBColor(0x86, 0x81, 0x7B),      # Espresso tinted toward Ivory -- labels, prompts
    "ink_light": RGBColor(0xB8, 0xB4, 0xAD),    # Espresso tinted further -- decline callouts
    "card_bg": RGBColor(0xFF, 0xFF, 0xFF),      # Pure White -- card/table backgrounds
    "white": RGBColor(0xFF, 0xFF, 0xFF),        # header text on the dark Sage Green fill
    "rule": RGBColor(0xE1, 0xDE, 0xD3),         # Sage tinted toward Ivory -- dividers / gridlines
}

TITLE_FONT = "Gill Sans MT"
BODY_FONT = "Gill Sans MT"

SLIDE_W = Inches(10)
SLIDE_H = Inches(5.62)

TITLE_POS = (Inches(0.56), Inches(0.58), Inches(8.88), Inches(0.66))  # full content width --
                                                                       # the old 5.23in box was
                                                                       # sized for short titles
                                                                       # like "Instagram Insights"
                                                                       # and wrapped on the longer
                                                                       # "[Platform] -- Top Posts
                                                                       # This Period" titles
CONTENT_TOP = Inches(1.24)
CONTENT_LEFT = Inches(0.56)
CONTENT_WIDTH = Inches(8.88)
CONTENT_BOTTOM = Inches(5.35)

# Insights-slide layout budget (hero cards + table + optional callout + dual chart)
GAP = Inches(0.03)
SECTION_LABEL_OFFSET = Inches(0.24)  # space a section label reserves above its content --
                                      # must match add_section_label's own box height exactly
HERO_CARD_HEIGHT = Inches(0.70)
COMPARISON_ROW_HEIGHT = Inches(0.18)
CALLOUT_HEIGHT = Inches(0.20)
COMPARISON_COL_WIDTHS = [Inches(2.6), Inches(1.7), Inches(1.7), Inches(2.88)]
COMPARISON_FONT_SIZE = Pt(8)


# --------------------------------------------------------------------------
# Low-level slide-building helpers
# --------------------------------------------------------------------------

def add_blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = COLORS["background"]
    return slide


def add_rule(slide, y, color, left=CONTENT_LEFT, width=CONTENT_WIDTH, weight=Pt(0.75)):
    """The brand's recurring thin single-rule framing motif."""
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left, y, Emu(int(left + width)), y)
    connector.line.color.rgb = color
    connector.line.width = weight
    return connector


def add_title(slide, text):
    box = slide.shapes.add_textbox(*TITLE_POS)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.name = TITLE_FONT
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = COLORS["cobalt_deep"]
    add_rule(slide, Inches(1.16), COLORS["rule"])
    return box


def set_bullet(paragraph, char="●", marL=228600, indent=228600):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(marL))
    pPr.set("indent", str(-indent))
    buFont = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    buChar = pPr.makeelement(qn("a:buChar"), {"char": char})
    pPr.append(buFont)
    pPr.append(buChar)


def add_bullets(slide, items, bold=False, top=CONTENT_TOP, height=Inches(3.9)):
    box = slide.shapes.add_textbox(CONTENT_LEFT, top, CONTENT_WIDTH, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        run = p.add_run()
        run.text = item
        run.font.name = BODY_FONT
        run.font.size = Pt(14)
        run.font.bold = bold
        run.font.color.rgb = COLORS["ink"]
        set_bullet(p)
    return box


def add_subtitle_prompt(slide, text):
    box = slide.shapes.add_textbox(Inches(5.88), Inches(2.83), Inches(3.4), Inches(0.55))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = BODY_FONT
    run.font.size = Pt(11)
    run.font.italic = True
    run.font.color.rgb = COLORS["ink_mid"]
    return box


def add_section_label(slide, text, top):
    """Small secondary-tone heading used above a chart/table section."""
    box = slide.shapes.add_textbox(CONTENT_LEFT, top, CONTENT_WIDTH, Inches(0.24))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    r.font.name = BODY_FONT
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COLORS["ink_mid"]
    return box


def add_callout(slide, text, top, height=CALLOUT_HEIGHT):
    """One-line factual context note (e.g. a decline tied to fewer posts)."""
    box = slide.shapes.add_textbox(CONTENT_LEFT, top, CONTENT_WIDTH, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.name = BODY_FONT
    r.font.size = Pt(10)
    r.font.italic = True
    r.font.color.rgb = COLORS["ink_light"]
    return box


def add_dual_line_chart(slide, chart_title, categories, current_values, prior_values,
                         top=CONTENT_TOP, height=Inches(2.7)):
    """Current period vs. prior period, both plotted by day-in-period so the
    two lines are directly shape-comparable regardless of calendar dates."""
    add_section_label(slide, chart_title, Emu(int(top - SECTION_LABEL_OFFSET)))

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("This Period", current_values)
    chart_data.add_series("Last Period", prior_values)

    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE, CONTENT_LEFT, top,
                                     CONTENT_WIDTH, height, chart_data)
    chart = gframe.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(9)
    chart.legend.font.name = BODY_FONT
    chart.legend.font.color.rgb = COLORS["ink_mid"]

    plot = chart.plots[0]
    line_colors = (COLORS["cobalt_mid"], COLORS["ink_light"])
    for series, color in zip(plot.series, line_colors):
        series.format.line.color.rgb = color
        series.format.line.width = Pt(2)
        series.smooth = False

    value_axis = chart.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = COLORS["rule"]
    value_axis.major_gridlines.format.line.width = Pt(0.5)
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = Pt(9)
    value_axis.tick_labels.font.name = BODY_FONT
    value_axis.tick_labels.font.color.rgb = COLORS["ink_mid"]

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(9)
    cat_axis.tick_labels.font.name = BODY_FONT
    cat_axis.tick_labels.font.color.rgb = COLORS["ink_mid"]
    cat_axis.format.line.color.rgb = COLORS["rule"]

    return gframe


def add_empty_state(slide, headline, subtext, top=CONTENT_TOP, height=Inches(2.7)):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, CONTENT_LEFT, top,
                                   CONTENT_WIDTH, height)
    card.adjustments[0] = 0.04
    card.fill.solid()
    card.fill.fore_color.rgb = COLORS["card_bg"]
    card.line.color.rgb = COLORS["rule"]
    card.line.width = Pt(0.75)
    card.shadow.inherit = False

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = headline
    r1.font.name = BODY_FONT
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = COLORS["ink"]

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = subtext
    r2.font.name = BODY_FONT
    r2.font.size = Pt(11)
    r2.font.color.rgb = COLORS["ink_mid"]

    return card


def add_hero_cards(slide, cards, top, height=HERO_CARD_HEIGHT):
    """cards: list of dicts with label, value_str, delta_text, delta_color,
    number_color (Dusty Blush for the headline metric, Sage Green for the rest --
    see build_insights_slide)."""
    n = len(cards)
    gap = Inches(0.1)
    card_w = Emu(int((CONTENT_WIDTH - gap * (n - 1)) / n))
    x = CONTENT_LEFT
    for card in cards:
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top, card_w, height)
        shape.adjustments[0] = 0.08
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLORS["card_bg"]
        shape.line.color.rgb = COLORS["rule"]
        shape.line.width = Pt(0.75)
        shape.shadow.inherit = False

        tf = shape.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)

        p_label = tf.paragraphs[0]
        p_label.alignment = PP_ALIGN.CENTER
        r_label = p_label.add_run()
        r_label.text = card["label"]
        r_label.font.name = BODY_FONT
        r_label.font.size = Pt(9)
        r_label.font.color.rgb = COLORS["ink_mid"]

        p_value = tf.add_paragraph()
        p_value.alignment = PP_ALIGN.CENTER
        r_value = p_value.add_run()
        r_value.text = card["value_str"]
        r_value.font.name = BODY_FONT
        r_value.font.size = Pt(22)
        r_value.font.bold = True
        r_value.font.color.rgb = card["number_color"]

        p_delta = tf.add_paragraph()
        p_delta.alignment = PP_ALIGN.CENTER
        r_delta = p_delta.add_run()
        r_delta.text = card["delta_text"]
        r_delta.font.name = BODY_FONT
        r_delta.font.size = Pt(9)
        r_delta.font.bold = True
        r_delta.font.color.rgb = card["delta_color"]

        x = Emu(int(x + card_w + gap))


def add_table(slide, headers, rows, top, height, col_widths=None, row_height=None,
              font_size=Pt(10)):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gframe = slide.shapes.add_table(n_rows, n_cols, CONTENT_LEFT, top, CONTENT_WIDTH, height)
    table = gframe.table

    if col_widths:
        for c, w in enumerate(col_widths):
            table.columns[c].width = w
    if row_height:
        for r in range(n_rows):
            table.rows[r].height = row_height

    for c, header in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLORS["cobalt_deep"]
        cell.margin_top = Pt(2)
        cell.margin_bottom = Pt(2)
        p = cell.text_frame.paragraphs[0]
        p.font.name = BODY_FONT
        p.font.size = font_size
        p.font.bold = True
        p.font.color.rgb = COLORS["white"]

    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            # A cell value may be a plain string/number, or a (text, color)
            # tuple to override the default ink color -- used for the
            # color-coded Change column.
            text, color = value if isinstance(value, tuple) else (value, COLORS["ink"])
            cell = table.cell(r, c)
            cell.text = str(text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = (
                COLORS["card_bg"] if r % 2 == 0 else COLORS["background"]
            )
            cell.margin_top = Pt(2)
            cell.margin_bottom = Pt(2)
            p = cell.text_frame.paragraphs[0]
            p.font.name = BODY_FONT
            p.font.size = font_size
            p.font.color.rgb = color

    return gframe


# --------------------------------------------------------------------------
# Buffer data aggregation
# --------------------------------------------------------------------------

def parse_iso(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def flatten_posts(period_profiles):
    """period_profiles: list of channel entries; extracts .posts from a given period dict."""
    posts = []
    for profile in period_profiles or []:
        posts.extend(profile.get("posts") or [])
    return posts


def primary_metric(metrics, keys=("views", "impressions")):
    """Views/impressions naming differs by platform -- use whichever exists."""
    for k in keys:
        if k in metrics:
            return metrics[k] or 0
    return 0


def daily_totals_by_day_index(posts, period_start, n_buckets, metric_keys=("views", "impressions"),
                               period_days=None):
    """Buckets posts by day-in-period (bucket 0 = period_start), not
    calendar date -- so a current-period line and a prior-period line land
    on the same Day-1..Day-N x-axis and are shape-comparable regardless of
    calendar dates.

    period_days lets a period WIDER than n_buckets compress onto the same
    axis instead of being truncated -- e.g. Instagram's 60-day-wide prior
    window (kept wide for a fair post-count sample) still plots on the same
    30-point axis as the 30-day current period, 2 real days per bucket,
    rather than discarding 9 of its 10 real posts to fit a literal 30-day
    slice. Defaults to n_buckets (i.e. no compression, 1 real day/bucket)."""
    period_days = period_days or n_buckets
    buckets = [0] * n_buckets
    for post in posts:
        sent = post.get("sentAt")
        if not sent:
            continue
        try:
            day_offset = (parse_iso(sent).date() - period_start).days
        except ValueError:
            continue
        if 0 <= day_offset < period_days:
            bucket_index = min(n_buckets - 1, int(day_offset * n_buckets / period_days))
            buckets[bucket_index] += primary_metric(post.get("metrics") or {}, metric_keys)
    return buckets


def sum_posts_metrics(posts):
    """Sum each post's own metrics -- the only source used for both periods
    (aggregatedPostMetrics is off the table; see module docstring)."""
    totals = {}
    for p in posts:
        for k, v in (p.get("metrics") or {}).items():
            if k == "engagementRate":
                continue  # a percentage, not summable across posts
            totals[k] = totals.get(k, 0) + (v or 0)
    totals["postCount"] = len(posts)
    return totals


LOW_BASE_PCT_THRESHOLD = 300  # % change at/above this, from a nonzero prior, reads as
                              # "new activity" instead of a raw (misleadingly huge) percentage


def classify_change(current, prior):
    """Returns (pct_or_None, display_label) per the low-base labeling rule."""
    if prior == 0:
        return None, ("new activity" if current else "flat (0/0)")
    pct = (current - prior) / prior * 100
    if pct >= LOW_BASE_PCT_THRESHOLD:
        return pct, f"new activity (low base: {fmt_number(prior)}→{fmt_number(current)})"
    return pct, f"{pct:+.1f}%"


def weighted_engagement_rate(totals):
    """interactions / reach, falling back to impressions only where reach
    isn't tracked at all (Facebook). Returns (rate, interactions, denom_key, denom)."""
    interactions = sum(totals.get(k, 0) for k in ("reactions", "comments", "shares", "saves"))
    for denom_key in ("reach", "impressions"):
        if denom_key in totals:
            denom = totals[denom_key]
            rate = (interactions / denom * 100) if denom else 0.0
            return rate, interactions, denom_key, denom
    return None, interactions, None, None


def change_arrow_and_color(current, prior):
    """Growth -> up-arrow in Cobalt Mid; decline -> down-arrow in the muted
    Ink Light tone (deliberately not stoplight red/green); flat -> neutral,
    no arrow. "New activity" counts as growth. Shared by the hero-card
    deltas and the table's Change column so both read the same way."""
    pct, label = classify_change(current, prior)
    if label == "flat (0/0)":
        return label, COLORS["ink_mid"]
    if pct is not None and pct < 0:
        return f"↓ {label}", COLORS["ink_light"]
    return f"↑ {label}", COLORS["cobalt_mid"]


def build_comparison_table_rows(metric_labels, metric_keys, current_totals, prior_totals):
    rows = []
    for key in metric_keys:
        cv, pv = current_totals.get(key, 0), prior_totals.get(key, 0)
        change_text, change_color = change_arrow_and_color(cv, pv)
        rows.append([metric_labels[key], fmt_number(cv), fmt_number(pv),
                     (change_text, change_color)])

    cur_rate, _, denom_key, _ = weighted_engagement_rate(current_totals)
    pri_rate, _, _, _ = weighted_engagement_rate(prior_totals)
    if cur_rate is not None and pri_rate is not None:
        change_text, change_color = change_arrow_and_color(cur_rate, pri_rate)
        rows.append([f"Weighted Eng. Rate ({denom_key})", f"{cur_rate:.2f}%", f"{pri_rate:.2f}%",
                     (change_text, change_color)])

    return rows


def build_decline_callout(metric_labels, metric_keys, current_totals, prior_totals):
    """A short, factual, data-supported note -- only when metrics are actually
    down, and only connecting them to a cause (posting frequency) the data
    can actually support. Returns None when there's nothing to explain."""
    declining = []
    for key in metric_keys:
        if key == "postCount":
            continue
        cv, pv = current_totals.get(key, 0), prior_totals.get(key, 0)
        if pv > 0 and cv < pv:
            declining.append(metric_labels[key])

    cur_rate, _, _, _ = weighted_engagement_rate(current_totals)
    pri_rate, _, _, _ = weighted_engagement_rate(prior_totals)
    rate_declining = cur_rate is not None and pri_rate is not None and cur_rate < pri_rate
    if rate_declining:
        declining = declining + ["Engagement Rate"]

    if not declining:
        return None

    cur_posts, pri_posts = current_totals.get("postCount", 0), prior_totals.get("postCount", 0)
    posts_declined = pri_posts > 0 and cur_posts < pri_posts
    posts_grew = pri_posts > 0 and cur_posts > pri_posts

    tracked_count = len([k for k in metric_keys if k != "postCount"])
    majority_declining = len(declining) >= max(1, (tracked_count + 1) // 2)

    if majority_declining and posts_declined:
        pct = (cur_posts - pri_posts) / pri_posts * 100
        return (f"Most metrics are down this period alongside a {abs(pct):.0f}% drop in "
                f"posting frequency ({cur_posts} vs. {pri_posts} posts).")
    if posts_grew:
        return (f"{', '.join(declining)} softened despite posting volume increasing "
                f"({cur_posts} vs. {pri_posts} posts).")
    return f"{', '.join(declining)} declined versus the prior period."


def fmt_number(n):
    if n >= 1000:
        return f"{n / 1000:.1f}K"
    return str(n)


def clean_caption(text, max_chars):
    """Collapses embedded newlines/repeated whitespace (real captions often
    have blank lines before hashtags) and truncates to max_chars."""
    text = " ".join((text or "").split())
    if len(text) <= max_chars:
        return text
    return text[:max_chars - 1].rstrip() + "…"


TOP_POSTS_TABLE_COL_WIDTHS = [Inches(4.93), Inches(0.70), Inches(0.75),
                              Inches(0.85), Inches(0.85), Inches(0.80)]


# --------------------------------------------------------------------------
# Automatic goal generation
#
# Reads client-context.md's Goal Log for the goals set last period, checks
# each one's embedded machine-readable condition against this period's real
# totals, and generates 3-5 new goals for next period from the most
# actionable finding per platform. See the module docstring for the overall
# design; this section is self-contained and only touches client-context.md
# and the totals already computed elsewhere in this file.
# --------------------------------------------------------------------------

DEFAULT_CLIENT_CONTEXT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                            "client-context.md")

GOAL_PLATFORM_CONFIG = {
    "instagram": {"label": "Instagram", "chart_metric_keys": ("views", "impressions")},
    "tiktok": {"label": "TikTok", "chart_metric_keys": ("views", "impressions")},
    "facebook": {"label": "Facebook", "chart_metric_keys": ("impressions", "views")},
}

GOAL_LINE_RE = re.compile(
    r"^-\s*\[([ xX])\]\s*(.+?)\s*<!--\s*check:\s*(\w+)\.(\w+)\s*(>=|<=|>|<)\s*([\d.]+)\s*-->$"
)
GOAL_LINE_NO_CHECK_RE = re.compile(r"^-\s*\[([ xX])\]\s*(.+)$")


def load_client_context(path):
    try:
        with open(path) as f:
            return f.read()
    except FileNotFoundError:
        return ""


def platform_totals_snapshot(data, platform_key, chart_metric_keys):
    """Same computation build_instagram_slide/build_facebook_slide/build_tiktok_slide
    already do -- recomputed here independently for goal generation rather
    than threaded through the slide builders, to keep this section self-contained."""
    profiles = data.get("channels", {}).get(platform_key, [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    prior_posts = flatten_posts([p.get("prior", {}) for p in profiles])
    current_totals = sum_posts_metrics(current_posts)
    prior_totals = sum_posts_metrics(prior_posts)
    return current_posts, current_totals, prior_totals


def compute_top_post_share(posts, totals, chart_metric_keys):
    """Mirrors build_top_posts_slide's share-of-total calculation."""
    if not posts:
        return None
    metric_key = next((k for k in chart_metric_keys if k in totals), chart_metric_keys[0])
    period_total = totals.get(metric_key, 0)
    if not period_total:
        return None
    top_value = max(primary_metric(p.get("metrics") or {}, chart_metric_keys) for p in posts)
    return top_value / period_total * 100


def platform_finding_goal(platform_key, label, current_totals, prior_totals, share_pct):
    """The single most actionable finding for this platform, in priority
    order: posting-frequency drop, content-concentration risk, engagement-
    rate decline, else a low-key "maintain" goal. Returns (goal_text, check_spec)."""
    cur_posts, pri_posts = current_totals.get("postCount", 0), prior_totals.get("postCount", 0)
    cur_rate, _, _, _ = weighted_engagement_rate(current_totals)
    pri_rate, _, _, _ = weighted_engagement_rate(prior_totals)

    if pri_posts > 0 and cur_posts < pri_posts * 0.8:
        drop_pct = (pri_posts - cur_posts) / pri_posts * 100
        text = (f"Restore {label} posting cadence — down {drop_pct:.0f}% this period "
                f"({cur_posts} vs. {pri_posts} posts). Aim for {pri_posts}+ posts next period.")
        return text, (platform_key, "postCount", ">", pri_posts)

    if share_pct is not None and share_pct >= 40:
        text = (f"Diversify {label} content — one post drove {share_pct:.0f}% of this period's "
                f"total; publish more posts capable of carrying reach on their own rather than "
                f"relying on a single standout.")
        return text, (platform_key, "topPostShare", "<", round(share_pct, 1))

    if cur_rate is not None and pri_rate is not None and cur_rate < pri_rate:
        text = (f"Improve {label} engagement quality — rate fell from {pri_rate:.2f}% to "
                f"{cur_rate:.2f}% this period.")
        return text, (platform_key, "engagementRate", ">", round(pri_rate, 2))

    text = (f"Maintain {label}'s current posting cadence and engagement level — no concerning "
            f"trend this period; keep up the consistency next period.")
    return text, (platform_key, "postCount", ">=", cur_posts)


def cross_platform_goal(rate_by_platform):
    """A goal only when one platform trails the others by a meaningful
    engagement-rate margin (>=3 points) -- otherwise returns None rather
    than manufacturing a comparison that isn't actually notable."""
    valid = {k: v for k, v in rate_by_platform.items() if v is not None}
    if len(valid) < 2:
        return None
    best = max(valid, key=valid.get)
    worst = min(valid, key=valid.get)
    gap = valid[best] - valid[worst]
    if worst == best or gap < 3:
        return None
    text = (f"Close the engagement gap on {GOAL_PLATFORM_CONFIG[worst]['label']} — trailing "
            f"{GOAL_PLATFORM_CONFIG[best]['label']} by {gap:.1f} points "
            f"({valid[worst]:.2f}% vs. {valid[best]:.2f}%). Borrow whatever content approach is "
            f"working on {GOAL_PLATFORM_CONFIG[best]['label']}.")
    return text, (worst, "engagementRate", ">", round(valid[worst], 2))


def context_tie_in_goal():
    """A static, evergreen goal type tying content to the Milbon Gold Line
    package initiative noted in client-context.md. Not data-checkable (no
    metric captures "did a piece of content spotlight this package"), so it
    has no check condition and stays unchecked until marked by hand."""
    text = ("Publish at least one piece of content spotlighting a Milbon Gold Line "
            "package with a clear booking CTA — e.g. a treatment before/after paired "
            "with the five-week result window — to drive add-on revenue.")
    return text, None


def generate_new_goals(data):
    """Returns (goals, current_totals_by_platform, prior_totals_by_platform).
    goals is a list of (goal_text, check_spec_or_None) tuples, 3-5 total:
    up to 3 platform-specific findings + the context tie-in (always) + the
    cross-platform gap goal (only when the gap is real)."""
    current_by_platform, prior_by_platform, rate_by_platform = {}, {}, {}
    platform_goals = []

    for platform_key, cfg in GOAL_PLATFORM_CONFIG.items():
        posts, current_totals, prior_totals = platform_totals_snapshot(
            data, platform_key, cfg["chart_metric_keys"])
        share_pct = compute_top_post_share(posts, current_totals, cfg["chart_metric_keys"])
        current_totals["_topPostShare"] = share_pct  # stashed for later checkoff evaluation

        current_by_platform[platform_key] = current_totals
        prior_by_platform[platform_key] = prior_totals
        cur_rate, _, _, _ = weighted_engagement_rate(current_totals)
        rate_by_platform[platform_key] = cur_rate

        if current_totals.get("postCount", 0) == 0:
            continue  # nothing meaningful to say about a platform with zero posts this period
        platform_goals.append(
            platform_finding_goal(platform_key, cfg["label"], current_totals, prior_totals, share_pct)
        )

    goals = platform_goals + [context_tie_in_goal()]
    cross = cross_platform_goal(rate_by_platform)
    if cross:
        goals.append(cross)

    return goals[:5], current_by_platform, prior_by_platform


def evaluate_goal_check(check, totals_by_platform):
    if check is None:
        return False
    platform, metric, op, baseline = check
    totals = totals_by_platform.get(platform)
    if not totals:
        return False
    value = totals.get("_topPostShare") if metric == "topPostShare" else totals.get(metric)
    if metric == "engagementRate":
        value, _, _, _ = weighted_engagement_rate(totals)
    if value is None:
        return False
    return {">": value > baseline, "<": value < baseline,
            ">=": value >= baseline, "<=": value <= baseline}[op]


def parse_last_goal_entries(context_text):
    """Finds the LAST '### <month>' section in the file (Goal Log entries
    are always appended at the end, so this is always the most recent one)
    and extracts its goal lines. Returns (full_original_line, goal_text,
    check_spec_or_None) tuples -- the original line is kept so checkoff
    results can be written back via a simple string replace."""
    if "\n### " not in context_text:
        return []
    last_section = context_text.rsplit("\n### ", 1)[-1]
    entries = []
    for line in last_section.splitlines():
        stripped = line.strip()
        if not stripped.startswith("- ["):
            continue
        m = GOAL_LINE_RE.match(stripped)
        if m:
            _, text, platform, metric, op, baseline = m.groups()
            entries.append((line, text, (platform, metric, op, float(baseline))))
            continue
        m2 = GOAL_LINE_NO_CHECK_RE.match(stripped)
        if m2:
            entries.append((line, m2.group(2), None))
    return entries


def apply_goal_checkoffs(context_text, entries_with_done):
    """entries_with_done: list of (full_original_line, done_bool). Flips
    '- [ ]' to '- [x]' in the file text for goals the data confirmed."""
    for original_line, done in entries_with_done:
        if done and "- [ ]" in original_line:
            context_text = context_text.replace(original_line, original_line.replace("- [ ]", "- [x]", 1), 1)
    return context_text


def summarize_period_changes(current_by_platform, prior_by_platform):
    """Short factual "what changed" lines for the Goal Log's monthly entry --
    the same underlying data as the goals themselves, in prose."""
    lines = []
    for platform_key, cfg in GOAL_PLATFORM_CONFIG.items():
        cur, pri = current_by_platform[platform_key], prior_by_platform[platform_key]
        if cur.get("postCount", 0) == 0:
            continue
        cur_rate, _, _, _ = weighted_engagement_rate(cur)
        pri_rate, _, _, _ = weighted_engagement_rate(pri)
        note = f"{cfg['label']}: {cur.get('postCount', 0)} vs. {pri.get('postCount', 0)} posts"
        if cur_rate is not None and pri_rate is not None:
            direction = "up" if cur_rate > pri_rate else "down" if cur_rate < pri_rate else "flat"
            note += f", engagement {direction} ({pri_rate:.2f}% → {cur_rate:.2f}%)"
        lines.append(note + ".")
    return lines


def append_goal_log_entry(context_text, month_label, what_changed_lines, goals):
    lines = [f"### {month_label}", "", "**What changed:** " + " ".join(what_changed_lines),
             "", "**Goals set:**"]
    for text, check in goals:
        if check:
            platform, metric, op, baseline = check
            lines.append(f"- [ ] {text} <!-- check: {platform}.{metric} {op} {baseline} -->")
        else:
            lines.append(f"- [ ] {text}")
    return context_text.rstrip() + "\n\n" + "\n".join(lines) + "\n"


# --------------------------------------------------------------------------
# Slide builders
# --------------------------------------------------------------------------

def build_cover_slide(prs, client, month, content_link):
    slide = add_blank_slide(prs)

    title_box = slide.shapes.add_textbox(Inches(1.18), Inches(1.6), Inches(7.64), Inches(1.85))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Monthly Marketing Report"
    r.font.name = TITLE_FONT
    r.font.size = Pt(48)
    r.font.bold = True
    r.font.color.rgb = COLORS["cobalt_deep"]

    # Cover-only signature accent rule -- the one deliberately blush placement
    # of the recurring frame motif, distinct from the rule-colored line used
    # under every other slide's title.
    add_rule(slide, Inches(3.5), COLORS["gilded"], left=Inches(4.0), width=Inches(2.0))

    subtitle_box = slide.shapes.add_textbox(Inches(3.79), Inches(1.32), Inches(2.42), Inches(0.48))
    p = subtitle_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = client
    r.font.name = BODY_FONT
    r.font.size = Pt(13)
    r.font.color.rgb = COLORS["ink_mid"]

    month_box = slide.shapes.add_textbox(Inches(4.32), Inches(3.6), Inches(1.35), Inches(0.5))
    p = month_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = month
    r.font.name = BODY_FONT
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = COLORS["gilded"]

    label_box = slide.shapes.add_textbox(Inches(0.27), Inches(4.29), Inches(2.42), Inches(0.48))
    p = label_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Content Link:"
    r.font.name = BODY_FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = COLORS["ink_mid"]

    link_box = slide.shapes.add_textbox(Inches(1.25), Inches(4.85), Inches(5.67), Inches(0.4))
    p = link_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = content_link
    r.font.name = BODY_FONT
    r.font.size = Pt(12)
    r.font.color.rgb = COLORS["ink"]

    return slide


METRIC_LABELS = {
    "postCount": "Posts", "views": "Views", "impressions": "Impressions", "reach": "Reach",
    "reactions": "Reactions", "comments": "Comments", "shares": "Shares", "clicks": "Clicks",
    "saves": "Saves",
}

# The finalized metric set per channel (order = display order). Each list
# excludes "Weighted Eng. Rate" -- build_comparison_table_rows appends that
# row itself, since its denominator (reach vs. impressions) is derived per
# channel rather than hardcoded here.
INSTAGRAM_METRICS = ["postCount", "views", "reach", "reactions", "comments", "shares", "saves"]
TIKTOK_METRICS = ["postCount", "views", "reach", "reactions", "comments", "shares"]
FACEBOOK_METRICS = ["postCount", "impressions", "reactions", "comments", "shares", "clicks"]


def comparison_table_height(n_metric_keys):
    """+1 for the Weighted Eng. Rate row, +1 for the header row."""
    return Emu(int(COMPARISON_ROW_HEIGHT * (n_metric_keys + 2)))


def build_hero_cards_config(current_totals, prior_totals, headline_key, include_reach):
    """2-3 hero cards: the platform's headline volume metric (Dusty Blush --
    the one sparing accent use per Insights slide), Reach if this platform
    tracks it, and Weighted Eng. Rate (both Sage Green)."""
    def card(label, current_value, prior_value, number_color, is_rate=False):
        delta_text, delta_color = change_arrow_and_color(current_value, prior_value)
        value_str = f"{current_value:.2f}%" if is_rate else fmt_number(current_value)
        return {"label": label, "value_str": value_str, "delta_text": delta_text,
                "delta_color": delta_color, "number_color": number_color}

    cards = [card(METRIC_LABELS[headline_key], current_totals.get(headline_key, 0),
                  prior_totals.get(headline_key, 0), COLORS["gilded"])]

    if include_reach:
        cards.append(card("Reach", current_totals.get("reach", 0), prior_totals.get("reach", 0),
                           COLORS["cobalt_deep"]))

    cur_rate, _, _, _ = weighted_engagement_rate(current_totals)
    pri_rate, _, _, _ = weighted_engagement_rate(prior_totals)
    if cur_rate is not None and pri_rate is not None:
        cards.append(card("Weighted Eng. Rate", cur_rate, pri_rate, COLORS["cobalt_deep"], is_rate=True))

    return cards


DEFAULT_FOLLOWER_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                          "follower-history.json")


def load_follower_history(path):
    """follower-history.json is written by fetch_followers.js (Meta Graph
    API snapshots, appended monthly). Missing/unreadable file just means no
    follower card is shown yet -- not an error, since this is a brand-new
    log that fills in over time."""
    try:
        with open(path) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def follower_hero_card(follower_log, metric_key):
    """metric_key: "facebookFollowers" or "instagramFollowers". Meta's Graph
    API only exposes a current snapshot, not history, so growth can only be
    shown once a second monthly snapshot exists -- the first month shows the
    count alone, no invented delta. Returns None if there's no snapshot for
    this platform at all yet."""
    entries = sorted(
        (e for e in follower_log if e.get(metric_key) is not None),
        key=lambda e: e["date"],
    )
    if not entries:
        return None

    current = entries[-1][metric_key]
    if len(entries) >= 2:
        delta_text, delta_color = change_arrow_and_color(current, entries[-2][metric_key])
    else:
        delta_text, delta_color = "First snapshot", COLORS["ink_mid"]

    return {"label": "Current Followers", "value_str": fmt_number(current),
            "delta_text": delta_text, "delta_color": delta_color,
            "number_color": COLORS["cobalt_deep"]}


def build_insights_slide(prs, title, current_posts, prior_posts, prior_period_days, metric_keys,
                          chart_metric_keys, include_reach, extra_hero_card=None):
    """title -> hero stat cards -> comparison table -> optional decline
    callout -> dual-line chart (current vs. prior, by day-in-period).

    extra_hero_card: an optional 4th (Instagram) / 3rd (Facebook) card --
    used for the "Current Followers" snapshot from fetch_followers.js.
    TikTok has no follower data source at all (see module docstring), so
    its wrapper never passes one."""
    slide = add_blank_slide(prs)
    add_title(slide, title)

    current_totals = sum_posts_metrics(current_posts)
    prior_totals = sum_posts_metrics(prior_posts)

    if current_totals["postCount"] == 0:
        add_empty_state(slide, "No published posts this period",
                         "Check back next month once new posts are published.",
                         top=CONTENT_TOP, height=Inches(3.9))
        return slide

    headline_key = next((k for k in chart_metric_keys if k in current_totals), chart_metric_keys[0])

    cards = build_hero_cards_config(current_totals, prior_totals, headline_key, include_reach)
    if extra_hero_card:
        cards.append(extra_hero_card)
    add_hero_cards(slide, cards, CONTENT_TOP, HERO_CARD_HEIGHT)
    hero_bottom = Emu(int(CONTENT_TOP + HERO_CARD_HEIGHT))

    table_top = Emu(int(hero_bottom + GAP))
    rows = build_comparison_table_rows(METRIC_LABELS, metric_keys, current_totals, prior_totals)
    table_height = comparison_table_height(len(metric_keys))
    add_table(slide, ["Metric", "Current", "Prior", "Change"], rows, table_top, table_height,
              col_widths=COMPARISON_COL_WIDTHS, row_height=COMPARISON_ROW_HEIGHT,
              font_size=COMPARISON_FONT_SIZE)
    table_bottom = Emu(int(table_top + table_height))

    callout_text = build_decline_callout(METRIC_LABELS, metric_keys, current_totals, prior_totals)
    if callout_text:
        callout_top = Emu(int(table_bottom + GAP))
        add_callout(slide, callout_text, callout_top)
        next_top = Emu(int(callout_top + CALLOUT_HEIGHT + GAP))
    else:
        next_top = Emu(int(table_bottom + GAP))

    chart_top = Emu(int(next_top + SECTION_LABEL_OFFSET))
    chart_height = Emu(int(CONTENT_BOTTOM - chart_top))

    current_period_start = datetime.now(timezone.utc).date() - timedelta(days=30)
    chart_prior_start = current_period_start - timedelta(days=prior_period_days)

    # Instagram's prior window is 60 days wide (kept wide for a fair
    # post-count sample -- a literal 30-day slice only had 1 post). Rather
    # than truncate to that thin slice for the chart too, compress the full
    # window onto the same 30-point axis (2 real days per bucket) so the
    # prior line still reflects all 10 real posts, not just 1.
    # Label only every 5th day (plus Day 1) -- 30 dense labels crowded the
    # axis; blank-string categories in between render cleanly in both
    # PowerPoint and Google Slides, unlike an OOXML tickLblSkip property
    # which Google Slides' PPTX import doesn't reliably honor.
    categories = [f"Day {i + 1}" if (i + 1) == 1 or (i + 1) % 5 == 0 else "" for i in range(30)]
    current_values = daily_totals_by_day_index(current_posts, current_period_start, 30, chart_metric_keys)
    prior_values = daily_totals_by_day_index(prior_posts, chart_prior_start, 30, chart_metric_keys,
                                              period_days=prior_period_days)
    add_dual_line_chart(slide, f"Daily {METRIC_LABELS[headline_key]} — Current vs. Prior Period",
                         categories, current_values, prior_values, top=chart_top, height=chart_height)

    return slide


def build_top_posts_slide(prs, platform_label, current_posts, current_totals, chart_metric_keys):
    """title -> "[top post] drove X% of total [metric]" callout stat ->
    the per-post table (moved here from the Insights slide, one dedicated
    slide per platform now instead of only TikTok having one)."""
    slide = add_blank_slide(prs)
    add_title(slide, f"{platform_label} — Top Posts This Period")

    if not current_posts:
        add_empty_state(slide, "No published posts this period",
                         "Check back next month once new posts are published.",
                         top=CONTENT_TOP, height=Inches(3.9))
        return slide

    metric_key = next((k for k in chart_metric_keys if k in current_totals), chart_metric_keys[0])
    metric_label = METRIC_LABELS[metric_key]
    period_total = current_totals.get(metric_key, 0)

    top_posts = sorted(current_posts,
                        key=lambda p: primary_metric(p.get("metrics") or {}, chart_metric_keys),
                        reverse=True)[:5]

    top_value = primary_metric(top_posts[0].get("metrics") or {}, chart_metric_keys)
    share_pct = (top_value / period_total * 100) if period_total else 0
    caption = (top_posts[0].get("text") or top_posts[0].get("postId") or "-")
    caption_short = clean_caption(caption, 45)

    stat_box = slide.shapes.add_textbox(CONTENT_LEFT, CONTENT_TOP, CONTENT_WIDTH, Inches(0.6))
    tf = stat_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    for text, bold, color in (
        (f"“{caption_short}” drove ", False, COLORS["ink"]),
        (f"{share_pct:.0f}%", True, COLORS["gilded"]),
        (f" of this period's total {metric_label.lower()}.", False, COLORS["ink"]),
    ):
        r = p.add_run()
        r.text = text
        r.font.name = BODY_FONT
        r.font.size = Pt(14)
        r.font.bold = bold
        r.font.color.rgb = color

    table_top = Emu(int(CONTENT_TOP + Inches(0.6) + GAP))
    table_height = Emu(int(CONTENT_BOTTOM - table_top))

    table_rows = []
    for post in top_posts:
        sent = post.get("sentAt")
        try:
            date_str = parse_iso(sent).strftime("%b %d")
        except (TypeError, ValueError):
            date_str = "-"
        cap = clean_caption(post.get("text") or post.get("postId") or "-", 180)
        metrics = post.get("metrics") or {}
        table_rows.append([
            cap, date_str,
            fmt_number(primary_metric(metrics, chart_metric_keys)),
            metrics.get("reactions") or 0,
            metrics.get("comments") or 0,
            metrics.get("shares") or 0,
        ])
    add_table(slide, ["Post", "Date", metric_label, "Reactions", "Comments", "Shares"],
              table_rows, table_top, table_height,
              col_widths=TOP_POSTS_TABLE_COL_WIDTHS, row_height=Inches(0.55), font_size=Pt(9))

    return slide


def prior_period_days_for(profiles):
    """Reads the actual prior-window width buffer-metrics.js used (60 days
    for Instagram, 30 for TikTok/Facebook) rather than assuming -- falls
    back to 30 if a channel entry is missing this for any reason."""
    for profile in profiles or []:
        days = (profile.get("prior") or {}).get("range", {}).get("days")
        if days:
            return days
    return 30


def build_instagram_slide(prs, data, follower_log):
    profiles = data.get("channels", {}).get("instagram", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    prior_posts = flatten_posts([p.get("prior", {}) for p in profiles])
    follower_card = follower_hero_card(follower_log, "instagramFollowers")
    return build_insights_slide(prs, "Instagram Insights", current_posts, prior_posts,
                                 prior_period_days_for(profiles),
                                 INSTAGRAM_METRICS, ("views", "impressions"), include_reach=True,
                                 extra_hero_card=follower_card)


def build_instagram_top_posts_slide(prs, data):
    profiles = data.get("channels", {}).get("instagram", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    current_totals = sum_posts_metrics(current_posts)
    return build_top_posts_slide(prs, "Instagram", current_posts, current_totals,
                                  ("views", "impressions"))


def build_facebook_slide(prs, data, follower_log):
    profiles = data.get("channels", {}).get("facebook", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    prior_posts = flatten_posts([p.get("prior", {}) for p in profiles])
    follower_card = follower_hero_card(follower_log, "facebookFollowers")
    return build_insights_slide(prs, "Facebook Insights", current_posts, prior_posts,
                                 prior_period_days_for(profiles),
                                 FACEBOOK_METRICS, ("impressions", "views"), include_reach=False,
                                 extra_hero_card=follower_card)


def build_facebook_top_posts_slide(prs, data):
    profiles = data.get("channels", {}).get("facebook", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    current_totals = sum_posts_metrics(current_posts)
    return build_top_posts_slide(prs, "Facebook", current_posts, current_totals,
                                  ("impressions", "views"))


def build_tiktok_slide(prs, data):
    profiles = data.get("channels", {}).get("tiktok", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    prior_posts = flatten_posts([p.get("prior", {}) for p in profiles])
    # TikTok has no follower data source at all (neither Buffer nor Meta's
    # Graph API cover it), so unlike Instagram/Facebook this card is a
    # manually-typed placeholder -- not pulled from any API or log file --
    # meant to be overtyped with the real count before each presentation.
    manual_follower_card = {
        "label": "Current Followers", "value_str": "[Count]",
        "delta_text": "Manual entry", "delta_color": COLORS["ink_mid"],
        "number_color": COLORS["cobalt_deep"],
    }
    return build_insights_slide(prs, "TikTok Insights", current_posts, prior_posts,
                                 prior_period_days_for(profiles),
                                 TIKTOK_METRICS, ("views", "impressions"), include_reach=True,
                                 extra_hero_card=manual_follower_card)


def build_tiktok_top_posts_slide(prs, data):
    profiles = data.get("channels", {}).get("tiktok", [])
    current_posts = flatten_posts([p.get("current", {}) for p in profiles])
    current_totals = sum_posts_metrics(current_posts)
    return build_top_posts_slide(prs, "TikTok", current_posts, current_totals,
                                  ("views", "impressions"))


def build_placeholder_bullet_slide(prs, title, bullets, bold=False, subtitle_prompt=None):
    slide = add_blank_slide(prs)
    add_title(slide, title)
    add_bullets(slide, bullets, bold=bold)
    if subtitle_prompt:
        add_subtitle_prompt(slide, subtitle_prompt)
    return slide


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


def adjacent_month(month, offset):
    idx = MONTHS.index(month)
    return MONTHS[(idx + offset) % 12]


def main():
    parser = argparse.ArgumentParser(description="Generate the monthly marketing report .pptx")
    parser.add_argument("--data", required=True, help="Path to Buffer metrics JSON")
    parser.add_argument("--output", default="monthly-report.pptx")
    parser.add_argument("--client", default="Salon Veritas")
    parser.add_argument("--month", required=True, help="Reporting month, e.g. 'July'")
    parser.add_argument("--prior-month", help="Defaults to the month before --month")
    parser.add_argument("--next-month", help="Defaults to the month after --month")
    parser.add_argument("--content-link", default="[Add content link]")
    parser.add_argument("--follower-log", default=DEFAULT_FOLLOWER_LOG_PATH,
                         help="Path to follower-history.json written by fetch_followers.js")
    parser.add_argument("--client-context", default=DEFAULT_CLIENT_CONTEXT_PATH,
                         help="Path to client-context.md (brand facts + auto-generated Goal Log)")
    args = parser.parse_args()

    with open(args.data) as f:
        data = json.load(f)
    follower_log = load_follower_history(args.follower_log)

    prior_month = args.prior_month or adjacent_month(args.month, -1)
    next_month = args.next_month or adjacent_month(args.month, 1)

    # --- Automatic goal generation -----------------------------------------
    context_text = load_client_context(args.client_context)
    new_goals, current_by_platform, prior_by_platform = generate_new_goals(data)

    prior_goal_entries = parse_last_goal_entries(context_text)
    checked_prior_goals = [(text, evaluate_goal_check(check, current_by_platform))
                            for _, text, check in prior_goal_entries]
    context_text = apply_goal_checkoffs(
        context_text, [(line, done) for (line, _, _), (_, done)
                        in zip(prior_goal_entries, checked_prior_goals)]
    )

    if checked_prior_goals:
        last_month_bullets = [f"{text} ✅" if done else text for text, done in checked_prior_goals]
    else:
        last_month_bullets = ["[Add goal - mark ✅ if complete]"] * 4

    next_month_bullets = [text for text, _ in new_goals]

    what_changed = summarize_period_changes(current_by_platform, prior_by_platform)
    context_text = append_goal_log_entry(context_text, args.month, what_changed, new_goals)
    with open(args.client_context, "w") as f:
        f.write(context_text)
    # -------------------------------------------------------------------------

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    build_cover_slide(prs, args.client, args.month, args.content_link)
    build_instagram_slide(prs, data, follower_log)
    build_instagram_top_posts_slide(prs, data)
    build_facebook_slide(prs, data, follower_log)
    build_facebook_top_posts_slide(prs, data)
    build_tiktok_slide(prs, data)
    build_tiktok_top_posts_slide(prs, data)
    build_placeholder_bullet_slide(prs, "Miscellaneous Marketing Items",
                                    ["[Add item]"] * 4, bold=False)
    build_placeholder_bullet_slide(prs, "Marketing WINS",
                                    ["[Add item]"] * 4, bold=False)
    build_placeholder_bullet_slide(prs, f"Last Month - Goals for {prior_month}",
                                    last_month_bullets,
                                    subtitle_prompt="What else? Goals for Month? Questions?")
    build_placeholder_bullet_slide(prs, f"Next Month - Goals for {next_month}",
                                    next_month_bullets,
                                    subtitle_prompt="What else? Goals for Month? Questions?")

    prs.save(args.output)
    print(f"Wrote {args.output}")

    print("\n" + "=" * 70)
    print(f"GENERATED GOALS -- Next Month slide ({next_month})")
    print("=" * 70)
    for text, _ in new_goals:
        print(f"  - {text}")

    print("\n" + "=" * 70)
    print(f"LAST MONTH'S GOALS -- checkoff results ({prior_month})")
    print("=" * 70)
    if checked_prior_goals:
        for text, done in checked_prior_goals:
            print(f"  [{'DONE ✅' if done else 'not yet'}] {text}")
    else:
        print("  No prior automated goals found in client-context.md -- "
              "Last Month slide used the manual placeholder bullets instead.")
    print(f"\nclient-context.md updated: {args.client_context}")


if __name__ == "__main__":
    main()
