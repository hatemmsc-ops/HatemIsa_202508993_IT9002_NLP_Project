# eMYAA Marketing & Reporting Workspace — Project Summary

Branch: `claude/emyaa-ceo-marketing-redesign-q8t96n`
Folder: `emyaa-ceo-deck/`
Owner: Hatem Isa Hatem, Marketing Officer, WealthTech, Ajyad Capital

This file is a running summary of everything built in this workspace, so a new
session (or a person) can pick up context quickly without reading the full chat
history. All source files referenced here are read-only under `input/`; all
deliverables are under `output/`.

---

## 1. What this workspace is for

Redesigning and maintaining eMYAA's (Ajyad Capital's Shari'ah-compliant trading
app) marketing and management reporting decks in the Board of Directors'
visual language (navy gradient header `#101B3B → #22376F`, gold rule
`#D4AF37`, Exo 2 font, navy/green/gold accents), plus related comms (emails,
meeting minutes, content calendars).

**Hard rules followed throughout:**
- Never edit or overwrite original uploaded files — all saved read-only under `input/`.
- Never invent numbers — unverified figures are labelled "reported" or "per team".
- Git: develop and push only on the branch above; no PRs unless asked.
- The client tracker spreadsheet (names, phone numbers) is kept **out of git**
  entirely — see `.gitignore` at the repo root.
- Plain punctuation, no em dashes, no banned marketing buzzwords (house style).

---

## 2. Key decks produced (all in `output/`)

### 2.1 CEO Marketing Update — redesigned (`output/final-deck/`)
`eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pptx` (+ PDF), 6 slides.
Built via `working/build_deck.js` (pptxgenjs). Canva-safe: every chart is
hand-drawn from shapes (bars, lines, dots), not native PowerPoint charts.
Slides: Title · What we did (status table + creator stills) · Results (70%
traffic, +53 downloads, +4 onboarded) · Google Ads & installs (line chart +
funnel) · English vs Arabic (diverging bar chart) · Decisions & open items.
Supporting audit files in `output/audit/`: `CEO-Audit.md`,
`Slide-Change-Log.md`, `Open-Items.md`, `Percentage-Audit.csv` (every % with
numerator/denominator/period), `Source-Traceability.csv`.

**Key corrections made during the audit:** LinkedIn "+6,075%" is growth in new
followers (247 vs 4), not total followers; referral traffic is 5.5% not 6.6%;
"English leads on every metric" is false — Arabic has more key events per
session (1.05 vs 0.94); 793% active-user rise stays labelled "reported" (no
export for the prior week).

### 2.2 Board of Directors Update (`output/board-updates/`, `output/board-17-sep/`)
- `eMYAA_Board_of_Directors_Update_-_Updated_30_Sep_2026.pptx` — full 14-slide
  deck rebuilt from the original via direct XML edits (python-pptx/lxml) after
  it was accidentally deleted from Canva. Page 3 date fixed; appendix pages 7–8
  updated to the 30 Sep dashboard (492 downloads, 55 trades, 72 incomplete KYC).
- `eMYAA_Board_Update_-_Updated_Slides.pptx` — a smaller file with just the
  changed slides (3, 7, 8), for faster re-import.
- `eMYAA_Board_of_Directors_Update_-_17.09.2026.pptx`/`.pdf` — the original
  17 Sep deck, saved as-is (PPTX) plus a rendered PDF, per user request.
- **Resolved finding:** $12,206 = total deposits recorded on the platform to
  date (confirmed by management 29 Sep); it matches the Board figures. The
  dashboard's $10,029.69 figure is AUM (assets under management), a different
  measure — not a conflict.

### 2.3 Daily Performance Dashboard, as PPT (`output/dashboard-pptx/`)
`eMYAA_-_Daily_Performance_Dashboard.pptx` (+ PDF), 63 slides: cover + trend
overview (shape-drawn line charts) + 61 daily report slides (30 Sep back to
20 May), each with 4 KPI columns and change-vs-previous indicators. Built via
`working/build_dashboard.js`; source data extracted from the dashboard PDFs
via `working/dashboard_extract.py` into `working/audit-data/dashboard-pages.json`.
Design improved over the literal PDF: dropped decorative icons, standardized
money formatting, added a "Total KYC applications" derived card.

### 2.4 CEO Update / Performance Analysis deck (`output/ceo-update/`)
Two related decks here:

**A. `eMYAA_Performance_Analysis_-_01.10.2026_-_Updated_29_Sep.pptx`** (24
slides) — the "Performance Analysis" deck updated to 29 Sep figures, then
progressively extended:
- Slide 2 Performance Update: revenue updated, Key Highlights rewritten as a
  whole-deck summary (Platform / Marketing / Client Experience & IT).
- New IT slides (5–8): IT Status at a Glance, Next App timeline
  (Exante/QATestLab), Android security-update results, Helpdesk backlog
  (moved to appendix).
- New Client Experience slides (10–11): Overview (21 called, 72 emailed, 4
  WhatsApp, 6 onboarded = 28.6%, split into onboarded/not-onboarded/funding-gap)
  and Feedback & Next Steps, built from `eMYAA_Client_Monthly_Tracker_-_2026.xlsx`
  — **kept out of git** (personal data); only aggregate counts committed via
  `working/cx_tracker_summary.py` → `working/audit-data/cx-tracker-weekly.csv`.
- Page numbers added; appendix grouped by section (IT, Client Experience,
  Marketing) to match the main-deck order.
- Build script: `working/build_ceo_it_update.py`.

**B. `eMYAA_Progress_Update_CEO_-_01.10.2026_-_Canva-ready.pptx`** — the
user's own Canva-exported "Progress Update" deck, fixed for Canva-safety and
design consistency via `working/canva_fix_progress_update.py`:
- Every SVG image (Canva renders SVG image-fills as a solid blue box) replaced
  with a rendered PNG (via `cairosvg`).
- Every gradient fill (Canva flattens gradients to light blue) replaced with a
  pre-rendered PNG gradient.
- One consistent navy-band header (gold stripe, gold section kicker, white
  title, Ajyad logo) applied to every content slide.
- Revenue card updated to $140.21 (+30.4%); every positive change arrow forced
  green (some were grey).

### 2.5 October 2026 Content Calendar (`output/content-calendar/`)
`Content_Calendar_OCT_2026_V3_-_eMYAA.pptx` — built from the user's V2 calendar
with Eman Almansoor's review comments applied in both English and Arabic
(`working/apply_october_calendar_edits.py`), e.g.:
- "Which one is right for me?" → "Which one fits my investment goals?"
- SPUS description corrected (it tracks its own Shari'ah index, not literally
  the S&P 500 index).
- "43.7% of global stock market cap" → dated and sourced ("in 2025", SIFMA
  Capital Markets Fact Book 2026) with a footer citation on every slide of
  that carousel.
- Arabic/English "Buy" button references corrected to match the actual app UI
  label; "in real time" removed (no real-time data available) in both
  languages.
- Final pass: all of Eman's review comments (and the comment-author metadata)
  stripped from the file before sending to the agency.

### 2.6 Vision Bank / Saudi market-entry decks (`output/vision-bank/`)
Two decks, both external-facing (no internal budgets beyond what's explicitly
approved, no client personal data):

**A. `eMYAA_Saudi_Market_Plan_-_Vision_Bank.pptx`** (8 slides) — Claude-built
from scratch (`working/build_vision_bank.js`) covering: who eMYAA is; the
Saudi channel/cost plan table (Google Ads + LinkedIn = USD 30,000 approved;
Jodel + outdoor = ~USD 48,700 proposed, not yet approved); Google Ads/LinkedIn
detail; outdoor media options from Nadher Media (elevator + LED screens,
USD 16,000–67,200 by duration); Jodel takeover-poll proposal with a BISB
(Bahrain Islamic Bank) peer case study; discussion topics. All prices
converted to and shown in USD (source quotes were in SAR/BHD). Jodel app logo
and two real phone-screenshot mockups (Riyadh poll + BISB ad) extracted from
the vendor PDF and composited into phone-frame graphics
(`working/charts/vb_jodel_*.png`).

**B. `eMYAA_KSA_Market_Discussion_-_v2.pptx`** (10 slides) — the user's own
Canva-fixed/trimmed version of the above, further extended per Yaqoob
AlBalooshi's review comments (`working/update_ksa_deck.py`):
- New slide: "Instagram, X and TikTok: What Worked and Why" — three-column
  comparison with Instagram reel thumbnails (Musheera, Ali Sabeel, Mention &
  Win, Abdulelah Al Harbi) pulled from `working/charts/tile9_*`.
- New slide: marketing activity table since launch (May–Oct 2026), sourced
  from the July/August monthly reports and the Board budget.
  table added to the channels slide (Dates column).
- New slide: a worked Jodel takeover-poll **copy example** on an emotional/
  personal-goals angle (debt payoff, car, education, marriage), in Saudi
  dialect Arabic with English translation and a follow-up ad line — flagged
  for compliance review (no promise of returns; debt-payoff framing may need
  removal).
- Channel table rebuilt with an explicit "Dates" column (bold) per Yaqoob's
  comment.

**Meeting minutes** for the Vision Bank discussion were drafted (not saved as
a file — delivered in chat) covering: current approach endorsed by Vision
Bank's Jarrah Al Sari; conversion-agency referral (SAR 250–300 per
conversion); outdoor media deprioritized as too expensive; focus
recommended on ads/social/digital presence. Attendees: Yaqoob Yusuf
AlBalooshi (organizer), Eman Almansoor, Hatem Isa Hatem (Ajyad); Jarrah Al
Sari (Vision Bank). Meeting held on Microsoft Teams.

---

## 3. Source data and inputs (`input/`, all read-only, chmod 444)

| Folder | Contents |
|---|---|
| `input/source/` | Primary + Enhanced CEO marketing decks, Performance Analysis (orig. + audited), Progress Update, IT Status Update, client tracker XLSX (gitignored), monthly reports (May–Aug), content calendars |
| `input/board-reference/` | 4 Board deck candidates (17.09 is the style reference), BEC appendix, funding draft, Board Members deck, BOD Update PDF (29.09) |
| `input/analytics/` | GA4 exports, 61 daily dashboard PDFs, LinkedIn/X exports, OneLink stats, KPI workbook |
| `input/campaign-evidence/` | $5,000 Prize Campaign T&Cs, landing pages, approval checklist, 24F briefing |
| `input/brand-assets/` | Exo 2 font files (`.fntdata`), Ajyad/eMYAA logos |
| `input/saudi-entry/` | Nadher Media rate cards & KSA campaign proposals, Jodel/BISB case study, the user's own KSA deck iterations, Yaqoob's review-comment screenshots |

---

## 4. Verified facts / numbers to reuse (avoid re-deriving)

- **Platform (as at 30 Sep 2026):** 492 downloads (Android 262 / iOS 230),
  61 clients onboarded, 15 funded, 55 total trades, $12,206.00 total deposits
  (gross, confirmed by management), $10,539.81 AUM, 72 incomplete KYC, 143
  total KYC applications.
- **As at 29 Sep:** 490 downloads, 61 onboarded, revenue $107.52 → later
  updated to $140.21 (+30.4%) in the Progress Update.
- **Google Ads/LinkedIn Saudi budget:** USD 30,000 total, Sep–Dec 2026
  (Google Search Saudi $9K, LinkedIn Saudi $9K, rest-of-GCC $6K each for
  Google and LinkedIn). Google Ads live since 22 Sep; LinkedIn live since
  6 Oct.
  Source: Board update 17 Sep 2026.
- **Marketing budget (full year):** USD 80,000 allocated; USD 34,482 spent
  14 May–6 Sep; influencer programme USD 9,037.
- **English vs Arabic ads (1–28 Sep GA4):** English 212 sessions / 97.2%
  engagement; Arabic 176 sessions / 86.9% engagement; but Arabic leads on key
  events per session (1.05 vs 0.94) — "English wins everything" is false.
- **Instagram:** 2,116 followers (Sep report); Musheera reel 5.3M views, Ali
  Sabeel 1.3M, both paid-boosted; Mention & Win pilot (11–18 Jun) $600 spend →
  266,218 views, 22,580 clicks, 6 new accounts.
- **X (Twitter):** 738 followers, 96.4% Bahrain / 3.6% Saudi; paid support
  paused since July.
- **TikTok:** 15 followers; bot/fraud traffic found May 2026 (706,445 Pangle
  clicks), paid TikTok removed 22 Jun 2026, clean since.
- **Nadher Media outdoor rates (converted to USD @ 3.75 SAR):** 2wk $16,000 /
  4wk $27,200 / 8wk $44,800 (latest offer, 30% off $64,000 original) / 12wk
  $67,200. 83 elevator-screen towers + 8 LED faces at The Zone, Riyadh.
- **Jodel (Nadher Media proposal):** USD 3,862.50 + VAT, 7 days, all of KSA,
  525,000 display/video impressions + 2,000,000-impression 24h takeover poll.
  Jodel KSA audience: 2.5M active users, ages 21–35, avg CTR 0.44%.

---

## 5. Open items / things still needing input from the user

1. **CEO Progress Update (Canva-ready):** still has stale slides 10, 25, 26
   (old Overview numbers, "Helpdesk under review") that duplicate/contradict
   newer slides — removal was proposed but not yet actioned pending user OK.
2. **"Top 10 companies = 38% of S&P 500 weight"** figure (Oct content
   calendar, Post 4) has no dated source — either confirm a source or the
   simplified wording (already applied) stays.
3. **17 Sep 2026 CEO update file** — user said one should exist; could not be
   located in uploads or Drive. Still outstanding if the user finds/sends it.
4. **Next App sanity-test date** — blank in the IT Status Update source file;
   shows as "to be confirmed" everywhere it's referenced.
5. **Vision Bank deck — Jodel poll debt-payoff option** — flagged for
   compliance review before use (Shari'ah/regulatory optics of linking
   investing to debt).
6. **Outdoor/Jodel spend vs Board contingency** — proposed Jodel+outdoor
   spend (~USD 48,700) exceeds the USD 32,477 contingency line in the Board
   budget; flagged, not resolved. (Per the later Vision Bank meeting, Jarrah
   advised against outdoor spend for now, which may resolve this.)
7. **Vision Bank meeting minutes** — meeting date and Vision Bank attendee's
   exact title still placeholders; email draft sent to user but not confirmed
   sent.

---

## 6. Useful scripts (all run from `emyaa-ceo-deck/`)

| Script | Purpose |
|---|---|
| `working/build_deck.js` (Node/pptxgenjs) | CEO Marketing Update redesign |
| `working/embed_fonts.py` | Embeds Exo 2 `.fntdata` into any output pptx |
| `working/build_dashboard.js` | Daily Performance Dashboard → 63-slide PPT |
| `working/dashboard_extract.py` | Parses dashboard PDFs → JSON |
| `working/build_ceo_it_update.py` | Adds IT + Client Experience sections to the Performance Analysis deck |
| `working/cx_tracker_summary.py` | Aggregates the (gitignored) client tracker into a safe CSV |
| `working/canva_fix_progress_update.py` | SVG→PNG, gradient→PNG, consistent header, for the user's Canva export |
| `working/apply_october_calendar_edits.py` | Applies Eman's EN/AR copy edits to the Oct calendar |
| `working/build_vision_bank.js` | Builds the from-scratch Saudi market-entry deck |
| `working/update_ksa_deck.py` | Extends the user's KSA deck per Yaqoob's comments |
| `skills pptx validate.py` (via `office/validate.py`) | OOXML validation check before delivering any pptx |

All builds are followed by: embed fonts (if a new file) → validate → render
to PDF via `soffice --headless --convert-to pdf` → spot-check PNG renders →
commit → push → `SendUserFile`.

---

## 7. Git / delivery conventions used in this session

- Every deliverable committed to `claude/emyaa-ceo-marketing-redesign-q8t96n`
  and pushed immediately after validation passes.
- Large multi-file deliveries zipped and split under the 30 MB upload limit
  when needed.
- Client-identifying data (tracker spreadsheet) never committed — see root
  `.gitignore`.
- Commit messages end with the Claude attribution trailer in effect at the
  time of the commit (model name varies by session — check the latest
  system reminder for the current one before committing).
