# eMYAA CEO Marketing Update - redesign workspace

Rebuild of the 1 Oct 2026 CEO marketing update in the Board of Directors Update (17.09.2026) visual language.

## Outputs

```
output/final-deck/eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pptx   8 slides, Exo 2 embedded
output/final-deck/eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pdf
output/audit/CEO-Audit.md             audit memo for the CEO
output/audit/Slide-Change-Log.md      slide-by-slide changes
output/audit/Percentage-Audit.csv     every percentage: numerator, denominator, period, recalculation, status
output/audit/Source-Traceability.csv  every claim and its evidence file
output/audit/Open-Items.md            what is still open
output/previews/slide-1..8.png      rendered slides
```

## Inputs (all read-only, never edited)

```
input/source/                primary deck, Enhanced deck, supporting decks, memos, monthly reports, content calendars
input/board-reference/       Board decks (17.09 is the style reference) and the candidate note
input/analytics/             GA4 exports, daily dashboards, LinkedIn/X exports, OneLink stats, KPI workbook
input/campaign-evidence/     $5,000 campaign T&Cs, landing pages, approval checklist, 24F briefing, screenshots
input/brand-assets/          Exo 2 fonts and logos taken from the Board deck
```

## Rebuild

From this folder (needs `pptxgenjs` on NODE_PATH):

```
python3 working/audit-data/build_audit.py
node working/build_deck.js
python3 working/embed_fonts.py
```

`working/charts/` holds the header, background, logo and screenshot crops used by the build.
