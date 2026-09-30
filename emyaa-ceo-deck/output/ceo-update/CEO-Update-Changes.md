# CEO update deck (Performance Analysis): 29 Sep update

- **Source:** `input/source/supporting/eMYAA_Performance_Analysis_-_01.10.2026_-_Audited.pptx`. It is unchanged and read-only.
- **Output:** `eMYAA_Performance_Analysis_-_01.10.2026_-_Updated_29_Sep.pptx`. This is the full deck (now 24 slides). The 29 Sep figures stay. The IT update of 01 Oct is added, see the section at the end. Slide numbers in the sections before it refer to the original 19-slide order.
- **Data:** Daily Performance Dashboard, 29 Sep 2026. Deposits of $12,206 were confirmed by management on 29 Sep.
- **"Prev" values** are the previous update (414, 51, $10,195.88, $8,529.69, 47).

## Slide 2, Performance Update

| Card | Was | Now |
|---|---|---|
| Total downloads | 480, +15.9% | 490, +18.4% (vs 414) |
| Total trades | 54, +14.9% | 55, +17.0% (vs 47) |

Unchanged on 29 Sep:
- Users onboarded 61 (+19.6%)
- Incomplete KYC 70
- Deposits $12,206 (+19.7%)
- AUM $10,539.81 (+23.6%)

Revenue ($107.52) and withdrawals ($1,666.19) are also unchanged. The files contain no newer figure for either.

**Key Highlights** are rewritten as a summary of the whole deck:

1. **Platform:** downloads up, onboarding lagging. 490 downloads (+18.4%) and 61 clients (+19.6%). Only 4 of the 53 downloads since 22 Sep have onboarded, and 70 users have KYC incomplete.
2. **Marketing:** paid media testing, campaigns live. Google Ads brought 70% of September site traffic. Mention & Win and the $5,000 campaign are live. The Manal Talal video is due in October.
3. **Client Experience and IT:** clearing onboarding blockers. 8 of 31 contacted clients onboarded (25.8%). The Helpdesk is in compliance review, and the new app build is waiting on Exante tests.

These replace three earlier lines:
- "continued organic interest": paid search drove most of the traffic.
- "Onboarding kept pace with acquisition": the numbers since launch contradict it.
- "Paid Media Resumed": out of date.

**Formatting:**
- The first two highlight lines are back to regular weight, matching the third line and the earlier version of the slide.
- The "Key Highlights" heading box is wider so it can't wrap.

## Slide 3, KPI table

- **September trades:** 8 becomes 9 (55 total, less 46 to end of August).
- **Trades gap:** -31 becomes -30.
- The rest of the table is unchanged.

## Slides 9, 16 and 18

| Slide | Was | Now | Basis |
|---|---|---|---|
| 9, Mention & Win note | 26.2K views, 1,129 likes on launch reel | 26.3K views, 1,129 likes on launch post | Later Instagram screenshot (views only go up) |
| 9, Manal Talal | Scheduled for end of October | Scheduled for 22 October | October content calendar V2 |
| 9, Google Ads and LinkedIn box | "Launched 22 Sep, testing phase" | "Testing since 22 Sep" | Still in testing (Board runway) |
| 16, LinkedIn note | Followers up 247 in 30 days (+6,075%) | 247 new followers in 30 days (vs 4 before, +6,075%) | +6,075% is growth in new followers (LinkedIn export: 4 in August) |
| 18, Paid row note | new campaign launched 22 Sep | campaign live since 22 Sep | Wording |
| 18, Referral | 39 (6.6%); ajyadcapital.com, Instagram, ChatGPT, Google Tag Assistant | 32 (5.5%); ajyadcapital.com, Instagram, Facebook, Google Tag Assistant (7 test visits) | GA4 export: the referral medium does not include ChatGPT |
| 18, Footnote | 579 of 587 (98.6%); 8 outside the breakdown | 572 of 587 (97.4%); the other 15 are ChatGPT (9) and source not available (6) | GA4 export, 1-28 Sep |

**Not changed:** the "Q3-Q4 2026 Rollout Schedule" graphic on slide 9 is an image. It still shows Abdulelah Al Harbi as "Scheduled" and says "two live, four scheduled", while the table beside it says the video is live. It needs replacing at the source.

## IT update, 01 Oct 2026 (added after the 29 Sep update)

**Source:** `input/source/it/IT_Status_Update_CEO_01_10_2026.pptx` (read-only), including the notes written into its table and its five review comments. **Build:** `working/build_ceo_it_update.py` from `working/ceo-update-base-29Sep.pptx`. New slides use plain shapes and tables, so they import into Canva.

**Slide order:** see "Page numbers and appendix" at the end.

| Comment or note in the IT file | What was done |
|---|---|
| Google Play: say it is for security, so it is not confused with the new app | Row and slide named "Google Play security update (current app)"; slide 7 says "Not the Next App" |
| Shufti KYC: say it is an internal demo that will not go live | "Internal demo only; it will not go live" |
| Exante: give the sanity test date | **Left as "to be confirmed". The date was blank in the file.** |
| Shufti billing: say the amount was given as compensation | "USD 50 production test credit given as compensation". Please check whether "compensation" should cover the USD 9,605 refund too |
| Add CRM and the other platform in the same format | WealthTech CRM (29 Oct, Planned) and Reporting Dashboard (TBC, Exploring) added as rows |
| Add a target date column | "Target / done" column added |
| Helpdesk slide: remove names and IDs | Removed |
| Helpdesk slide: appendix | Moved to the appendix (page 18) |

**Also fixed in the IT content:**
- The Next App slide said testing starts when v1 arrives "on 09 Oct". Every other place says 08 Oct, so 08 Oct is used.
- The empty table row was removed.
- The tester names on the Android slide were removed, to match the Helpdesk comment.
- The donut chart was replaced with a shape-drawn bar.
- **Checked:** 158+19+4+1 = 182, and 158/159 = 99.4%. The module rows add up to the totals, the Helpdesk counts are 7 = 3+1+3, and 10,750 to 5,625 is -47.7%.

**Existing slides brought in line with the IT file:**

| Slide | Was | Now |
|---|---|---|
| 2, Key Highlights, IT line | 8 of 31 contacted clients onboarded (25.8%); Helpdesk in compliance review; new app build awaits Exante tests | 6 of 21 clients called onboarded (28.6%); Helpdesk live 17 Sep; Android security update live 29 Sep; Next App v1 due 08 Oct |
| 8, IT Progress Update | "The Next Mobile App - In Development" | Retitled "Product Roadmap"; Next App "v1 due 08 Oct 2026", built by Exante and tested by QATestLab |
| 16, Platform Initiatives Roadmap | Helpdesk IN PROGRESS, "Under review by the Compliance since 07.09.2026" | DONE, "Live since 17 Sep 2026" |
| 17, Systems Planning | Helpdesk planning and deployment IN PROGRESS; "managed manually via Excel"; footer "Board Executive Committee Update", page "6" | All DONE; "Helpdesk is live"; footer "CEO Update", stray page number removed |
| 19, Exante status | "Following up" on release and on sanity test results | v1 due 08 Oct, release plan next week; Android update passed two rounds and is live; Next App sanity test date to be confirmed |

Long dashes in the edited text were replaced with short ones.

## Client Experience update, 17 Sep to 1 Oct 2026

**Sources:**
- Team figures: 21 called, 72 emailed, 4 WhatsApp, 6 completed onboarding.
- The Canva comments on the Overview slide (`input/source/client-experience/overview-slide-canva-comments.webp`).
- The Client Monthly Tracker.

The tracker holds client names, phone numbers and emails. It is kept on the machine only and is **not committed** (see `.gitignore`). Only counts are committed: `working/cx_tracker_summary.py` writes `working/audit-data/cx-tracker-weekly.csv`.

**The old Overview slide (31 contacted, 8 onboarded, 25.8%) is replaced by two slides:**

| Slide | What it holds |
|---|---|
| 10, Client Experience Overview | Six boxes: 21 called, 72 emailed, 4 WhatsApp, 6 completed onboarding, 15 still pending, 28.6% conversion. The 21 called are split by where they are now: onboarded 6, not onboarded 11, funding gap 4 (bar and cards). Summary with call results from the tracker: 12 reached, 2 no answer, 7 not logged |
| 11, Client Feedback and Next Steps | Five feedback points, including the new funding-gap point. Proposed next step for each group. Clients followed up per week from the tracker. A note on the tracker fields to fill in |

| Canva comment | What was done |
|---|---|
| From the last CEO meeting until now, update the numbers | Period 17 Sep (last CEO meeting) to 1 Oct; your figures used. In the tracker this is weeks 8 and 9 (21 clients): week 7 has no first calls on 17 Sep, and 18-19 Sep is the weekend |
| Email contacts and phone called | Boxes read "Clients called (phone)" and "Email contacts" |
| Segregate client types: funding gap, not onboarded, onboarded | Split of the 21: 6 / 11 / 4 |
| Add a point on people needing support to fund | "Funding gap" feedback point and next step |

**Checks:**
- **Tracker split:** weeks 8-9 have 17 Potential Customer and 4 Funded Gap Customer rows, 21 in total. The 6 who onboarded come from the 17, which leaves 11.
- **Onboarding in the tracker:** none of the 6 completions is marked in the tracker yet.
- **Error in the old Canva version:** it said 6 of 20 = 20% and 14 = 80%. The right figures were 30% and 70%. With 21 called: 28.6% converted and 15 (71.4%) pending.
- **Missing tracker fields:** no row in weeks 8-9 has an owner, due date, next follow-up or funding value. 7 rows in week 9 have no stage.

## Page numbers and appendix

**Page numbers** match the slide position in Canva. They appear on every content slide in a small navy tag next to the eMYAA logo, with the footer "eMYAA Trading Platform | CEO Update | 01 Oct 2026". Appendix pages add "| Appendix". The title slide, the section dividers, Thank You and the Appendix divider have no number. The old Board footer and page number on Systems Planning were removed.

| Page | Slide |
|---|---|
| 1 | Title |
| 2-3 | Performance Update, KPIs |
| 4-8 | IT: divider, IT Status at a Glance, Next App, Android security update, Product Roadmap |
| 9-11 | Client Experience: divider, Overview, Feedback and Next Steps |
| 12-13 | Marketing: divider, Sponsored Ads |
| 14 | Thank You |
| 15 | Appendix divider, now with a contents line |
| 16-19 | Appendix, IT: Platform Initiatives Roadmap, Systems Planning, Helpdesk backlog, Exante status |
| 20 | Appendix, Client Experience: Client Interaction Journey |
| 21-24 | Appendix, Marketing: Website Analytics, Instagram, X/LinkedIn/TikTok, September Artwork |

The appendix follows the same section order as the main deck. Before this, the IT, marketing and client pages were mixed.
