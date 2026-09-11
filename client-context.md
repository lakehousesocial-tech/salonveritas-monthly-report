# Salon Veritas — Client Context

## Brand

- Business: Salon Veritas — a boutique hair color studio in Midtown Raleigh, NC
- Owner/stylist: Ashley
- Brand voice: Refined, warm, intentional, elevated, genuine — an unhurried, confident
  tone, never salesy. Speaks to a woman who wants her hair color to feel as intentional
  as everything else in her life. (Source: Salon Veritas Brand Kit, Google Drive.)
- Social channels (connected in Buffer, confirmed via the Buffer API):
  - Instagram: @salonveritas_ (business account)
  - TikTok: @ashleysalonveritas
  - Facebook: Salon Veritas (Page)

### Brand Voice Guardrails (for anyone writing Misc/Wins/Goals bullets by hand)

- No em dashes in any copy, captions, or scripts.
- No more than 5 hashtags on any single post, regardless of platform.
- No discount/cheap/"deal" language — Salon Veritas competes on craft, not price.
- No excessive emojis, all caps, or stacked exclamation points.
- Avoid generic salon phrases as a crutch ("glow up" is fine occasionally, not a tagline).

### Visual Identity

The physical-brand palette (from the salon's interior design and wordmark, per the
Salon Veritas Brand Kit):

| Role | Hex | Usage |
|---|---|---|
| Sage Green | `#707252` | Primary — grounding/botanical anchor, wordmark |
| Dusty Blush | `#D99F95` | Primary — warmth/femininity, script logotype |
| Warm Brass | `#B08D57` | Accent — fixtures, hardware, quiet luxury |
| Ivory/Cream | `#F5F1EA` | Neutral base — walls/backgrounds |
| Espresso Walnut | `#2B2620` | Grounding neutral — dark wood, contrast |

**For anything Claude generates as a report/deliverable (this monthly report deck
included), use the report design-system token set instead** — established by the
Salon Veritas Instagram Account Audit artifact
(`https://claude.ai/code/artifact/b300d567-1732-4ed5-aec5-7d20898e7d4f`), which refines
the same sage/blush/brass family into report-appropriate tokens and pairs them with
free, embeddable Google Fonts (the brand kit's Gill Sans MT isn't one). This is the
palette/fonts actually applied in `generate_report.py`:

| Token | Hex | Deck usage |
|---|---|---|
| sage (`--sage`) | `#707252` | Headers / primary accent |
| brass (`--brass`) | `#B8965A` | Chart lines / secondary accent |
| blush (`--blush`) | `#D99F95` | Sparing emphasis (headline stat only) |
| bg (`--bg`) | `#FAF7F4` | Slide background |
| text-primary | `#2A2118` | Body text |
| text-secondary | `#6B5F55` | Secondary labels |
| text-caption | `#8A7E74` | Captions / decline callouts |
| border | `#E2DBD4` | Dividers / rule motif / gridlines |

Typography: **Cormorant Garamond** (display/titles) + **DM Sans** (body/UI), per that
same artifact — both free Google Fonts, unlike the brand kit's licensed Gill Sans MT.
Do NOT source branding from the Wix site (`lakehousesocial.wixsite.com/salon-veritas`)
— it's a work-in-progress refresh, not the settled report design system.

### Content Pillars

1. **Color Transformation Reveals** — root touch-ups, balayage, full-color transformations
   as hook-first Reels/TikToks.
2. **Treatment Spotlights** — education/storytelling around professional treatments
   (e.g. the Milbon Illuminating Glow system), before/after pairs.
3. **Salon Atmosphere & Behind the Chair** — the space itself (brass fixtures, the
   flamingo powder room, the coffee nook), day-in-the-life/stylist personality content.
4. **Client Care & Education** — home care tips, product pairing, maintaining color
   between visits.
5. **Booking & Availability** — seasonal specials, opening availability, "comment to
   book" CTAs.

### Target Audience Profiles

- **Established Ellen (35–50)** — values expertise over price, wants low-maintenance
  high-impact color, responds to honest before/after reveals.
- **New to Raleigh Nicole (28–40)** — actively searching for a new stylist, compares
  on reviews/atmosphere, responds to warm welcoming-space content and testimonials.
- **Special Occasion Sophie (22–35)** — books around events, wants glam that
  photographs well, responds to glossy finished-look content and space tours.

## Active Initiatives

- **Wix site refresh**: in progress (`https://lakehousesocial.wixsite.com/salon-veritas`),
  built around the sage/blush/brass/ivory/espresso palette and the moodboard pulled from
  the salon's interior design.
- **Milbon Gold Line packages**: a premium treatment line and meaningful add-on revenue
  driver — deserves its own content spotlight (see "Treatment Spotlights" pillar).
- **Lake House Social content partnership**: on-site shoot days plus short-form video
  production across TikTok/Instagram/Facebook (see the SalonVeritas_LakeHouseSocial
  proposal in Drive for the full tiered scope).
- (Add other active initiatives — new service launches, seasonal campaigns,
  collaborations — here as they come up, so goal generation can tie content to them.)

## Goal Log

<!--
  Auto-generated and auto-checked by generate_report.py -- do not hand-edit the
  structure below without care. Each month's entry records:
    - What changed that period (a short data summary)
    - The goals set for the FOLLOWING period
  The following month's run re-evaluates each goal's <!-- check: ... --> condition
  against that period's real data and flips "- [ ]" to "- [x]" automatically if met.
  Format: "- [ ] <goal text> <!-- check: <platform>.<metric> <op> <baseline> -->"
  A goal with no check comment (e.g. the Milbon Gold Line content tie-in) isn't
  data-checkable and stays unchecked until marked by hand.

  This log starts empty -- the first run has no prior goals to check off, so the
  "Last Month" slide falls back to manual placeholder bullets for that one run only.
-->
