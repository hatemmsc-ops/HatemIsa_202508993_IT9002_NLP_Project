# eMYAA CEO Marketing Update - redesign workspace

Rebuild of the 1 Oct 2026 CEO marketing update. Current stage: **inputs check and number audit done; redesign paused pending missing sources.**
See `output/audit/Open-Items.md`.

```
input/source/            original decks (read-only, never edited)
input/board-reference/   Board deck or screenshots (missing)
input/analytics/         GA / social exports (only a GA screenshot so far)
input/campaign-evidence/ T&Cs, screenshots
input/brand-assets/      logo, fonts (missing)
working/extracted/       inventory.json, Enhanced-text.md, extract.py
working/rendered-before/ source deck rendered to PDF + PNG
working/audit-data/      build_audit.py (regenerates the audit CSVs)
output/audit/            Percentage-Audit.csv, Source-Traceability.csv, Open-Items.md
output/final-deck/       (after redesign)
output/previews/         (after redesign)
```

Regenerate: `python3 working/extracted/extract.py` and `python3 working/audit-data/build_audit.py` from this folder.
