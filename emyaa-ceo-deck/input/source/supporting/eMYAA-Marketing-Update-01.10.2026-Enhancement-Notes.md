# eMYAA Marketing Update — Enhancement Notes

Rebuild of `eMYAA Marketing Update - 01.10.2026.pptx` (the prior 8-slide story deck) into a 10-slide, executive-grade presentation, `eMYAA-Marketing-Update-01.10.2026-Enhanced.pptx`, plus a matching HTML slide version.

## Content inventory and slide disposition

| Original slide | Disposition |
|---|---|
| 1. What We've Completed | Split: headline facts feed the new Executive Summary; the detailed action table moves into the new Campaign & Content Tracker |
| 2. Google Ads & LinkedIn: The Launch | Kept as its own slide, redesigned with a donut chart replacing the bar chart, KPI cards, and the GA screenshot repositioned |
| 3. New Video Content | Merged into the new Campaign & Content Tracker (one status table covering all campaigns and content, instead of a separate video-only slide) |
| 4. Best Performing Content | Moved to Appendix, kept its table and the content-grid screenshot |
| 5. Instagram: Standout Channel | Merged into new Channel Performance slide (paired with LinkedIn for direct comparison) |
| 6. LinkedIn: Building Credibility | Merged into new Channel Performance slide |
| 7. English Is Outperforming Arabic in Saudi | Kept as its own slide (unchanged substance, redesigned visually) |
| 8. What's Next | Redesigned into the new Next Steps slide |
| *(new)* | Cover slide — did not exist before |
| *(new)* | Executive Summary — six status-coded callout cards, did not exist before |
| *(new)* | KPI Snapshot — a dedicated dashboard slide, did not exist before |
| *(new)* | Key Issues, Risks & Dependencies — consolidates open items that were previously scattered across slide footnotes, now a single decision-useful slide |

Net effect: 8 slides in, 10 slides out, but with the marketing status/video content consolidated (was 2 slides, now 1 table) and three genuinely new slides added (cover, executive summary, risks) that the source deck lacked entirely.

## What changed structurally

1. **A cover slide** was added — the old deck opened directly on a data slide.
2. **An executive summary** was added as six status-coded cards (Launched / Performing / Growing / In Progress / Needs Attention / Open Item), giving a five-second read before any detail.
3. **A KPI dashboard slide** consolidates the six most important marketing numbers in one place, previously scattered across the deck's first two slides.
4. **Instagram and LinkedIn were merged into one comparison slide** instead of two separate slides, since the audience benefit is seeing both channels side by side, not paging between them.
5. **Campaign and content status were merged into one tracker table** instead of splitting "What's Completed" and "New Video Content" — a single status matrix is the correct format for this content, not two.
6. **Risks and dependencies got their own slide.** In the source deck these were buried as small italic footnotes under individual slides; they are now a single, explicit, decision-useful slide, which is standard executive-reporting practice for open items.
7. **Appendix** now holds the detailed best-performing-content table and screenshot, keeping the core 9-slide narrative uncluttered.

## Charts added

- Instagram views-by-content-type and interactions-by-content-type (two clustered bar charts, replacing the source deck's basic bars, now in a matched blue/gold palette)
- LinkedIn follower growth (two-point bar chart, 284 → 531)
- Website traffic-by-source **donut chart** (new chart type, replacing a bar chart — better suited to a share-of-total story)
- English vs Arabic grouped bar chart (sessions and engagement rate side by side)

No new data was introduced for any chart; every value traces to the same source already used in the prior deck (screenshots, invoices, Dashboard #32).

## Visual system applied

Per your explicit instruction, the visual system was pulled directly from `eMYAA Board of Directors Update - 17.09.2026.pptx`, not invented:

- **Header gradient**: linear, left to right, `#101B3B → #22376F` — the exact two-stop gradient and angle used in the board deck's own slide headers, extracted from its XML.
- **Cover slide**: same gradient, full-bleed, matching the board deck's title-slide treatment.
- **Gold accent** (`#C9A84C`): used sparingly, for kickers, dividers, and card borders only, never as a fill.
- **Body area**: kept white/light (`#F7F9FC` panels), matching the board deck's own KPI-tile slides, which are white-bodied with a dark gradient header band, not full-dark-body slides.
- **Chart palette**: restricted to the navy family (`#121A2F`, `#1B2845`) plus one gold accent color, no default Office chart colors, no 3D, no clutter.
- **Status color coding**: gold for "Live," green for "Completed/Posted," navy-slate for "In Production/Scheduled" — consistent, restrained, not decorative.

One deliberate deviation from a literal 1:1 copy: the board deck's header band is taller (≈2.6M EMU) than what was used here (≈1.75M EMU), because this deck's slides carry denser tabular content than the board deck's. The gradient, colors, and structural logic are identical; only the proportion was adjusted for content density. Flagging this rather than silently changing it.

## Source limitations that remain

- Instagram/LinkedIn/X/TikTok have no standing analytics export or update cadence — every number in this deck came from a screenshot supplied in-session, not a repeatable report.
- The LinkedIn follower drop (586 → 531 across two pulls) is carried into this deck unresolved, same as the audited CEO deck — no source explains it.
- Gross deposits ($12,206) is a verbal figure from Malek, not a spreadsheet export.
- No week-over-week or month-over-month trend line could be built for Instagram or website sessions beyond single before/after snapshots, since no time-series export exists for either.

## Content moved to appendix

- The best-performing-content table (Musheera, Ali Sabeel, Mention & Win, Saudi Market Size Infographic) and its supporting content-grid screenshot.

## QA performed

Every slide was rendered to image and visually inspected for overflow, overlap, and truncation. One real issue was found and fixed during QA: the taller board-deck-style header initially collided with body text on the two slides that had a lead-in paragraph before their chart (Paid Media, English vs Arabic) — both were corrected by shifting the affected elements down, then re-rendered clean.

## Deliverables

- `/Users/hatemisa/Downloads/eMYAA-Marketing-Update-01.10.2026-Enhanced.pptx`
- `/Users/hatemisa/Downloads/eMYAA-Marketing-Update-01.10.2026-Enhanced.slides.html`
- `/Users/hatemisa/Downloads/eMYAA-Marketing-Update-01.10.2026-Enhancement-Notes.md` (this file)

The original `eMYAA Marketing Update - 01.10.2026.pptx` was not modified and remains in Downloads.
