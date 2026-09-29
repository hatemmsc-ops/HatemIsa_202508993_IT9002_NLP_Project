# eMYAA CEO Deck - Readiness Report and Open Items

Status: **redesign paused.** Setup, source render, extraction and the number audit are done.
Per the setup instruction, the redesign has not started because key inputs are missing.

Prepared 29 Sep 2026 for the 1 Oct 2026 CEO update.

## 1. Environment check

| Tool | Status |
|---|---|
| Python 3.11.15 | OK |
| python-pptx 1.0.2, lxml 6.1.3 | OK |
| pandas 3.0.6, numpy 2.4.6, matplotlib 3.11.2, seaborn 0.13.2, openpyxl 3.1.5 | OK |
| Pillow 12.3.0, PyMuPDF 1.28.2, pdf2image, cairosvg 2.9.1, svglib 2.2.0 | OK |
| LibreOffice 24.2.7.2 with Impress (installed this session) | OK, PPTX to PDF works |
| Poppler pdftoppm 24.02.0 (installed this session) | OK, PDF to PNG works |
| Tesseract 5.3.4 + pytesseract | OK (not relied on; figures were read visually) |
| pptxgenjs (Node) for native charts | OK |

Render test: the Enhanced deck converted to PDF and 10 PNGs at 150 dpi in
`working/rendered-before/`.

## 2. Files found vs missing

| Required input | Status | Where looked |
|---|---|---|
| `eMYAA-Marketing-Update-01.10.2026-Enhanced.pptx` | **Found** (upload). Copied read-only to `input/source/`. SHA-256 `e01b155c...dfd74c` matches the upload | Upload |
| `eMYAA-Marketing-Update-01.10.2026.pptx` (primary CEO deck) | **Missing** | Upload, repo, Google Drive |
| `Content-Calendar-OCT-2026-V2-eMYAA.pptx` | **Missing** (Drive has only Mar/Apr 2026 calendars) | Upload, repo, Google Drive |
| Finalized Board presentation or screenshots | **Missing** | Upload, repo, Google Drive |
| eMYAA logo (SVG / hi-res PNG), brand fonts, brand guide | **Missing** (logo only appears inside social screenshots) | Upload, repo, Google Drive |
| GA export (sessions by source, active users, EN vs AR) | **Missing**. Only a GA home-card screenshot inside the Enhanced deck | Upload, Drive |
| Instagram, LinkedIn, X, TikTok exports (slides 15-18) | **Missing** | Upload, Drive |
| Google Ads and LinkedIn campaign reports | **Missing** | Upload, Drive |
| Mention & Win T&Cs | **Missing** | Upload, Drive |
| TI7541 budget/ticket, TI7542 evidence | **Missing** (TI7542 is not mentioned in any file) | Upload, Drive |
| Malek gross-deposits export | **Missing** | Upload, Drive |
| KPI workbook behind the deck | **Missing** | Upload, Drive |
| Most Active Trader Raffle T&Cs | **Found** in Google Drive (19 Jul 2026). Excerpt saved to `input/campaign-evidence/` | Drive |

Evidence extracted from the Enhanced deck (not separate source files):
- `input/analytics/GA-home-card-21-27Sep-extracted-from-Enhanced-slide6.png`
- `input/campaign-evidence/Instagram-grid-extracted-from-Enhanced-slide10.png`

## 3. What the Enhanced deck looks like (for reference)

- Canvas 20 x 11.25 in (16:9), Calibri throughout, no speaker notes.
- Colors: navy `121A2F` / `22376F`, gold accent `C9A84C`, grey text `2D3748`, green `1BA97C`, red `C53030`.
- 10 slides: title, exec summary, KPI tiles, IG/LinkedIn charts, tracker, paid media, EN vs AR, risks, next steps, appendix.
- This is not the Board style. The Board deck is needed to copy the approved gradient, spacing and chart treatment.

## 4. Audit findings so far (full detail in the CSVs)

**Incorrect**
1. LinkedIn "+6,075%" (slide 2). Follower growth 284 to 531 is **+87%**. 6,075% only works as 247 new followers vs 4 in the prior period.
2. "Impressions and reactions both up double digits" (slide 2). They are up +118% and +155%, so triple digits.
3. Instagram "all-time" views of 2.27M (slide 4) cannot be all-time: one reel (Musheera) shows 5.3M views.
4. Mention & Win reel "26.2K views" and How-to video "319 views" are stale. The screenshot shows 26.3K and 332.
5. "Presenter-led reels outperform static graphics by orders of magnitude" (slide 10). The $500 static graphic has 186K views and the presenter how-to video has 332.

**Unverified, needs a source before publication**
6. **793% active users week over week.** The GA card shows 384 active users for 21-27 Sep, up 793.0% vs the previous 7 days. That implies about 43 users in 14-20 Sep. The card covers all traffic and all countries. "Saudi Arabia 7" on the screenshot is the realtime 30-minute panel only. So this must read "Reported increase; source calculation pending verification". It cannot be credited to the Saudi test launch yet. The line on the screenshot plots Event count, not Active users.
7. 69.7% paid share and the rest of the source mix. The arithmetic is consistent (409 of 587), but there is no GA export.
8. EN 54.5% / AR 45.2% session shares only match a 389 total, and the denominator isn't stated. The engagement rates (97.17%, 86.93%) recalculate correctly.
9. $30,000 budget, Ignite, TI7541. Seen only in the deck text; no invoice or ticket.
10. Targeting (Google Search to Starter/Emerging, LinkedIn to Affluent, Saudi residents only).
11. $12,206 gross deposits. Verbal only, and the period isn't stated.
12. LinkedIn followers: the deck says both "+247 in 30 days" and "dropped 586 to 531".
13. 558 new users since launch is higher than 384 active users on the 7-day GA card.
14. Mention & Win dates, mechanics, eligibility, "GCC-wide" (while ads are Saudi-only), and 1,129 likes.
15. Ali Sabeel "Completed 18 Sep" is inferred from the last invoice. The rollout row says August.

**Campaign separation**
16. Mention & Win ($500, 23 Sep to 20 Oct) and the Most Active Trader Raffle ($500 to $5,000 by volume tier, 1 Aug to 31 Dec 2026, draw 10 Jan 2027) are **different campaigns**. The raffle T&Cs don't show whether it is live.

**No data at all**
17. X, TikTok, Market Pulse, LinkedIn Ads leads and spend.

## 5. What is needed to proceed

Minimum to start the redesign with confidence:
1. The finalized Board deck (or 3 to 5 screenshots of it) and the eMYAA logo.
2. The primary CEO deck `eMYAA-Marketing-Update-01.10.2026.pptx`.
3. GA4 export: sessions by source/medium and active users, 14 to 30 Sep, with country.

Needed to clear the unverified items:
4. IG, LinkedIn, X and TikTok exports on one common period.
5. Google Ads and LinkedIn campaign reports (spend, clicks, leads, EN vs AR).
6. Mention & Win T&Cs, the TI7541 document, and Malek's gross-deposits export.
7. `Content-Calendar-OCT-2026-V2-eMYAA.pptx`.

If some of these can't be supplied, the fallback is to build now from the Enhanced deck.
Every figure above would carry an "unverified" or "reported" label, and the palette
would be an approximation of the Board style.
