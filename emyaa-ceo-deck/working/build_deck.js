// Builds eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pptx in the Board deck's visual language
// (eMYAA Board of Directors Update - 17.09.2026): navy gradient header #101B3B -> #22376F, gold rule #D4AF37,
// Exo 2 type, navy #132257 / green #1BA97C / gold accents, Ajyad Capital and eMYAA logos.
// Run from emyaa-ceo-deck/: NODE_PATH=<dir with pptxgenjs> node working/build_deck.js
// then: python3 working/embed_fonts.py (embeds the Board's Exo 2 font files)
const path = require("path");
const pptxgen = require("pptxgenjs");

const IMG = path.join(__dirname, "charts");
const OUT = path.join(__dirname, "..", "output", "final-deck", "eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pptx");

const C = {
  navy: "132257", green: "1BA97C", gold: "D4AF37", goldText: "9A7B1F", panel: "E3E8F2",
  line: "D9DFEA", soft: "C6CFE4", text: "1F2A44", muted: "7A849C", foot: "8A93A8", red: "C53030", white: "FFFFFF",
};
const F = "Exo 2";
const STATUS = {
  Live: C.green, Testing: C.goldText, Posted: C.navy, Completed: C.navy,
  "In production": C.muted, Scheduled: C.muted, Open: C.red, Conflicting: C.red,
  "Pending evidence": C.goldText, Reported: C.goldText, Verified: C.green,
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5, same 16:9 ratio as the Board deck
pres.author = "eMYAA Marketing";
pres.title = "Marketing Performance Update - CEO Update 1 October 2026";

function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: F, margin: 0, valign: "top", color: C.text }, o));
}
function status(label) {
  return { text: [
    { text: "● ", options: { color: STATUS[label] || C.muted } },
    { text: label, options: { bold: true, color: STATUS[label] || C.muted } },
  ] };
}
const none = { type: "none" };
const rule = { type: "solid", pt: 0.75, color: C.line };
function hdr(cells, fs = 11) {
  return cells.map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: fs, valign: "middle", border: [none, none, none, none] } }));
}
function row(cells, fs = 11) {
  return cells.map((c) => {
    const base = { fontSize: fs, color: C.text, valign: "middle", border: [none, none, rule, none] };
    if (typeof c === "string") return { text: c, options: base };
    return { text: c.text, options: Object.assign(base, c.options || {}) };
  });
}
function chartBase(title) {
  return {
    showTitle: true, title, titleFontSize: 11, titleColor: C.navy, titleFontFace: F,
    catAxisLabelColor: C.text, catAxisLabelFontSize: 10, catAxisLabelFontFace: F,
    dataLabelFontFace: F, dataLabelFontSize: 10, dataLabelColor: C.navy,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    legendFontFace: F, legendFontSize: 10, legendColor: C.text,
  };
}

// Board content-slide chrome
function contentSlide({ kicker, title, source, takeaway, n, notes }) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addImage({ path: path.join(IMG, "board_header.png"), x: 0, y: 0, w: 13.333, h: 1.42 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.42, w: 13.333, h: 0.027, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
  txt(s, kicker.toUpperCase(), { x: 0.62, y: 0.3, w: 10, h: 0.2, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, title, { x: 0.62, y: 0.52, w: 10.2, h: 0.82, fontSize: 20, bold: true, color: C.white, valign: "middle" });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 11.27, y: 0.5, w: 1.45, h: 0.42 });
  if (takeaway) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.62, y: 6.2, w: 12.1, h: 0.5, fill: { color: C.panel }, line: { color: C.panel, width: 0 } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.62, y: 6.2, w: 0.05, h: 0.5, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
    txt(s, [
      { text: "Takeaway   ", options: { bold: true, color: C.navy } },
      { text: takeaway, options: { color: C.text } },
    ], { x: 0.85, y: 6.2, w: 11.75, h: 0.5, fontSize: 11.5, valign: "middle" });
  }
  txt(s, "Source: " + source, { x: 0.62, y: 6.78, w: 10.6, h: 0.3, fontSize: 8, color: C.muted });
  txt(s, "eMYAA Trading Platform  |  CEO Marketing Update", { x: 0.62, y: 7.12, w: 6, h: 0.22, fontSize: 8, color: C.foot });
  txt(s, String(n), { x: 10.72, y: 7.12, w: 0.5, h: 0.22, fontSize: 8, color: C.foot, align: "right" });
  s.addImage({ path: path.join(IMG, "emyaa_navy.png"), x: 11.6, y: 6.93, w: 1.12, h: 0.4 });
  if (notes) s.addNotes(notes);
  return s;
}
function card(s, x, y, w, h, top) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, line: { color: C.line, width: 0.75 } });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.035, fill: { color: top }, line: { color: top, width: 0 } });
}

// ---------- 1. Title ----------
{
  const s = pres.addSlide();
  s.background = { path: path.join(IMG, "board_title_bg.png") };
  s.addImage({ path: path.join(IMG, "emyaa_white.png"), x: 0.75, y: 0.75, w: 1.75, h: 0.62 });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 10.67, y: 0.75, w: 1.92, h: 0.56 });
  txt(s, "Marketing Performance Update", { x: 0.75, y: 2.8, w: 11.9, h: 0.95, fontSize: 40, bold: true, color: C.white });
  txt(s, "CEO UPDATE", { x: 0.75, y: 3.8, w: 11, h: 0.45, fontSize: 20, bold: true, color: C.gold, charSpacing: 1 });
  txt(s, "1 October 2026", { x: 0.75, y: 4.3, w: 8, h: 0.4, fontSize: 16, color: C.white });
  txt(s, "Web: GA4, 1-28 Sep 2026  |  Platform: daily dashboard, 27 Sep 2026  |  Plan and budget: Board update, 17 Sep 2026", { x: 0.75, y: 6.7, w: 11.9, h: 0.3, fontSize: 10, color: C.soft });
  s.addNotes("Sources: GA4 exports 1-28 Sep 2026; GA4 home card 21-27 Sep; daily performance dashboards to 27 Sep; eMYAA Board of Directors Update 17 Sep 2026 (plan, budget, segments, CAC targets); content calendars Sep and Oct 2026 (V2); Instagram profile grid screenshot; $5,000 campaign T&Cs, approval checklist and 24F briefing. Anything without a source file is labelled 'reported' or 'per team'.");
}

// ---------- 2. Executive summary ----------
{
  const s = contentSlide({
    kicker: "Executive summary", n: 2,
    title: "Paid search now brings most site traffic and more downloads; onboarding has not kept pace",
    takeaway: "Early paid-media signals are positive at the top of the funnel. Treat 20 October as the go or hold point, judged on accounts opened.",
    source: "GA4 exports 1-28 Sep 2026; daily performance dashboards 14-27 Sep 2026; Instagram profile grid screenshot; Board update 17 Sep 2026 (CAC targets).",
    notes: "70% = 409 google / cpc sessions of 587 total, 1-28 Sep. Downloads: 437 on the 21 Sep dashboard, 473 on 27 Sep (+36); the week before added +14. Onboarded: 57 to 60 (+3) over the same six days. 1.3M and 186K are Instagram view counts from the profile grid screenshot.",
  });
  const cards = [
    ["DELIVERED", "1.3M", "views on the Ali Sabeel reel", "How-to and Abdulelah Al Harbi videos also posted in September. Manal Talal video set for 22 Oct.", C.navy],
    ["LIVE", "186K", "views on the Mention & Win post", "Monthly $500 contest across the GCC; current round 23 Sep to 20 Oct. Entries and accounts due at close.", C.green],
    ["TESTING", "70%", "of Sep site sessions from Google Ads", "409 of 587 sessions, 1-28 Sep. Downloads +36 in the six days after launch vs +14 the week before.", C.gold],
    ["DECISION", "20 Oct", "go or hold on paid media", "Judge against Board CAC targets: Starter $167, Emerging $417, Affluent $882. Onboarding added only +3.", C.red],
  ];
  const w = 2.87, gap = 0.2;
  cards.forEach((c, i) => {
    const x = 0.62 + i * (w + gap), y = 1.75;
    card(s, x, y, w, 4.2, c[4]);
    txt(s, c[0], { x: x + 0.25, y: y + 0.3, w: w - 0.5, h: 0.25, fontSize: 9, bold: true, color: C.goldText, charSpacing: 2 });
    txt(s, c[1], { x: x + 0.25, y: y + 0.65, w: w - 0.5, h: 0.85, fontSize: 40, bold: true, color: i === 1 ? C.green : C.navy });
    txt(s, c[2], { x: x + 0.25, y: y + 1.6, w: w - 0.5, h: 0.7, fontSize: 14.5, bold: true, color: C.navy });
    txt(s, c[3], { x: x + 0.25, y: y + 2.45, w: w - 0.5, h: 1.6, fontSize: 12.5, color: C.text });
  });
}

// ---------- 3. Status table ----------
{
  const s = contentSlide({
    kicker: "Marketing activity status", n: 3,
    title: "One campaign is live, two paid channels are in testing, and four videos are posted or in production",
    takeaway: "Nothing is labelled complete until close-out is confirmed. LinkedIn Ads has no reported results yet.",
    source: "GA4 exports 1-28 Sep 2026; Instagram grid screenshot; Board update 17 Sep 2026; Sep and Oct 2026 content calendars; 24F briefing; team updates ('per team' = no file).",
    notes: "Ali Sabeel: posted 2 Sep per the September calendar; the draft's 'Completed 18 Sep' is inferred from no further Ignite invoice; the Board budget shows a $2,200 Instagram sponsor ad for this video spent before 6 Sep. Shown as Posted. Mention & Win is a recurring monthly contest (24F briefing): calendar posts on 15 Sep and 18 Oct; the team gives 23 Sep to 20 Oct for the current round; the Board budget lists $3,200 of prizes and ads for Sep to Dec. Manal Talal is on the October calendar (V2) for 22 Oct.",
  });
  const rows = [
    hdr(["Initiative", "Objective", "Status", "Date", "Evidence", "Next step"]),
    row(["Mention & Win", "Account openings ($500 monthly prize)", status("Live"), "Round 23 Sep to 20 Oct", "Instagram: 186K and 26.3K views; $3,200 budgeted Sep-Dec", "Report entries and accounts after 20 Oct"]),
    row(["Google Ads", "Starter and Emerging investors, Saudi", status("Testing"), "From 22 Sep", "GA4: 409 sessions, 400 key events (1-28 Sep)", "Add spend and cost per key event"]),
    row(["LinkedIn Ads", "Affluent lead generation, Saudi", status("Testing"), "From 22 Sep", "No campaign report yet", "Pull leads and spend from Campaign Manager"]),
    row(["Ali Sabeel video", "Presenter-led awareness", status("Posted"), "Posted 2 Sep", "Instagram: 1.3M views; $2,200 sponsor ad", "Confirm close-out with Ignite"]),
    row(["How-to video, human presenter", "Onboarding education", status("Posted"), "September", "Instagram: 332 views", "Track clicks to account opening"]),
    row(["Abdulelah Al Harbi", "Creator video, Sep slot", status("Posted"), "September", "Per team; no view count supplied", "Add views to next update"]),
    row(["Manal Talal", "Creator video, Oct slot", status("In production"), "Calendar: 22 Oct", "Per team: brief sent, filming started", "Approve final cut before posting"]),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 12.1, colW: [1.9, 2.2, 1.45, 1.6, 2.75, 2.2], rowH: 0.52, fontFace: F, margin: [2, 5, 2, 5] });
}

// ---------- 4. Timeline ----------
{
  const s = contentSlide({
    kicker: "Campaign timeline", n: 4,
    title: "Five dated milestones run from 18 September to 22 October; the paid-media plan runs to December",
    takeaway: "20 October is the natural review point: the Mention & Win round closes and paid testing will have four weeks of data.",
    source: "Board update 17 Sep 2026 (runway, $30,000 plan); October 2026 content calendar V2 (22 Oct); team dates. Milestones not drawn to scale.",
    notes: "18 Sep is the date the draft gives for Ali Sabeel completion; it is inferred, not confirmed by Ignite (the video itself was posted 2 Sep per the September calendar). The Board runway shows 'LinkedIn and Google Ads, testing phase' started, followed by 'Scale LinkedIn and Google Ads' to December. Hamed Al Bloushi: November per the rollout plan; the September calendar had listed 29 Sep.",
  });
  const ms = [
    ["18 Sep", "Ali Sabeel video", "Marked complete (per team). 1.3M views"],
    ["22 Sep", "Paid testing starts", "Google Ads and LinkedIn Ads, Saudi Arabia"],
    ["23 Sep", "Mention & Win round", "$500 prize, ahead of salary days"],
    ["20 Oct", "Round closes", "Proposed paid-media review"],
    ["22 Oct", "Manal Talal video", "October calendar date, in production"],
  ];
  const x0 = 1.75, x1 = 11.6, y = 2.75, step = (x1 - x0) / (ms.length - 1);
  s.addShape(pres.shapes.LINE, { x: x0, y, w: x1 - x0, h: 0, line: { color: C.soft, width: 2 } });
  ms.forEach((m, i) => {
    const cx = x0 + i * step, done = i < 3;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.12, y: y - 0.12, w: 0.24, h: 0.24, fill: { color: done ? C.green : C.white }, line: { color: done ? C.green : C.navy, width: 2 } });
    txt(s, m[0], { x: cx - 1.1, y: y - 0.68, w: 2.2, h: 0.38, fontSize: 17, bold: true, color: C.navy, align: "center" });
    txt(s, m[1], { x: cx - 1.15, y: y + 0.3, w: 2.3, h: 0.32, fontSize: 12.5, bold: true, color: C.navy, align: "center" });
    txt(s, m[2], { x: cx - 1.1, y: y + 0.66, w: 2.2, h: 0.6, fontSize: 10.5, color: C.muted, align: "center" });
  });
  const by = 4.75, bx = 1.2, bw = 10.9, mw = bw / 4;
  ["SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"].forEach((m, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx + i * mw + 0.03, y: by - 0.42, w: mw - 0.06, h: 0.28, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
    txt(s, m, { x: bx + i * mw, y: by - 0.42, w: mw, h: 0.28, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle", charSpacing: 1 });
    if (i > 0) s.addShape(pres.shapes.LINE, { x: bx + i * mw, y: by - 0.1, w: 0, h: 1.1, line: { color: C.line, width: 0.75 } });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + mw * (21 / 30), y: by, w: bw - mw * (21 / 30), h: 0.36, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
  txt(s, "Paid media: test from 22 Sep, then scale to December ($30,000, Board plan)", { x: bx + mw * (21 / 30) + 0.15, y: by, w: 8, h: 0.36, fontSize: 10.5, bold: true, color: C.navy, valign: "middle" });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + mw * (22 / 30), y: by + 0.52, w: mw * (8 / 30) + mw * (20 / 31), h: 0.36, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
  txt(s, "Mention & Win round", { x: bx + mw * (22 / 30) + 0.15, y: by + 0.52, w: 3, h: 0.36, fontSize: 10.5, bold: true, color: C.white, valign: "middle" });
  s.addShape(pres.shapes.OVAL, { x: bx + 2 * mw + 0.12, y: by + 0.61, w: 0.18, h: 0.18, fill: { color: C.white }, line: { color: C.muted, width: 1.5 } });
  txt(s, "Hamed Al Bloushi video (scheduled)", { x: bx + 2 * mw + 0.4, y: by + 0.52, w: 2.4, h: 0.36, fontSize: 10, color: C.muted, valign: "middle" });
}

// ---------- 5. Paid media ----------
{
  const s = contentSlide({
    kicker: "Paid media launch", n: 5,
    title: "Google Ads and LinkedIn Ads are in testing against Board CAC targets; only Google has results so far",
    takeaway: "The plan and targets are set. What is missing is spend and conversion data, which Ignite must supply before 20 October.",
    source: "Board update 17 Sep 2026 (segments, CAC targets, $30,000 split); GA4 Google Ads campaigns export 1-28 Sep 2026. Residents-only rule per management.",
    notes: "Board segments: Starter under $1,000 and Emerging $1,000 to $5,000 via Google Ads; Affluent $5,000 to $50,000 via LinkedIn lead generation then in-person meeting. CAC targets: Starter $167, Emerging $417, Affluent $882. Budget: Google Search Saudi $9,000, LinkedIn Affluent Saudi $9,000, Google Search Qatar/UAE/Kuwait/Bahrain/Oman $6,000, LinkedIn Affluent same markets $6,000; funded from the $62,477 reserved budget. GA4 shows only the two Saudi Google campaigns so far. LinkedIn lead forms keep users on LinkedIn, so no LinkedIn sessions in GA4 is expected. The Ignite invoice TI7541 was cited by the previous audit but is not in this review set; TI7542 is not in the review set either.",
  });
  const lab = (t) => ({ text: t, options: { bold: true, color: C.muted } });
  const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: 12, valign: "middle", border: [none, none, none, none] } });
  const rows = [
    [{ text: "", options: { border: [none, none, none, none] } }, H("Google Ads"), H("LinkedIn Ads")],
    row([lab("Audience"), "Starter (under $1K) and Emerging ($1K-5K)", "Affluent ($5K-50K)"]),
    row([lab("Objective"), "Website conversions", "Leads, then in-person meeting"]),
    row([lab("Markets"), { text: "Saudi Arabia first; Qatar, UAE, Kuwait, Bahrain, Oman in plan", options: { colspan: 2 } }]),
    row([lab("Phase"), { text: "Testing since 22 September; scale phase follows (Board runway)", options: { colspan: 2 } }]),
    row([lab("Target CAC"), { text: "$167 Starter  |  $417 Emerging", options: { bold: true, color: C.navy } }, { text: "$882 Affluent", options: { bold: true, color: C.navy } }]),
    row([lab("Evidence"), "409 sessions, 400 key events (GA4, 1-28 Sep)", "None yet; lead forms stay on LinkedIn"]),
    row([lab("Next read-out"), "20 Oct: spend, cost per key event, accounts", "20 Oct: leads, cost per lead, meetings"]),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 8.1, colW: [1.35, 3.5, 3.25], rowH: [0.42, 0.52, 0.45, 0.52, 0.52, 0.45, 0.52, 0.52], fontFace: F, margin: [2, 6, 2, 6] });
  s.addChart(pres.charts.BAR, [
    { name: "Saudi Arabia", labels: ["Google Search", "LinkedIn Affluent"], values: [9000, 9000] },
    { name: "Other GCC", labels: ["Google Search", "LinkedIn Affluent"], values: [6000, 6000] },
  ], Object.assign(chartBase("Board-approved budget, Sep-Dec 2026 (USD)"), {
    x: 9.0, y: 1.65, w: 3.72, h: 4.4, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55, chartColors: [C.navy, C.gold],
    showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: '"$"#,##0', dataLabelColor: C.white,
    valAxisMaxVal: 16000, showLegend: true, legendPos: "b",
  }));
}

// ---------- 6. Paid-media evidence ----------
{
  const s = contentSlide({
    kicker: "Paid-media evidence", n: 6,
    title: "Active users reached 384 in the week of 21-27 September, a reported 793% rise as paid testing began",
    takeaway: "The rise lines up with the paid launch, and paid search is the first source for 74% of September users. The 793% stays 'reported' for now.",
    source: "GA4 home card, last 7 days (21-27 Sep) vs previous 7 days; GA4 Reports snapshot export, 1-28 Sep 2026. Prior-week value implied from GA's own % change.",
    notes: "Reported increase; source calculation pending verification. GA shows 384 active users, +793.0% vs the previous 7 days, which implies about 43 users in 14-20 Sep (384 / 8.93). The card covers all sources and all countries. The realtime 'Saudi Arabia 7' panel on the original screenshot covers the last 30 minutes only and was cropped out. The line on the card plots event count, not active users. From the 1-28 Sep export: 403 of 542 active users (74.4%) had google / cpc as first source.",
  });
  s.addChart(pres.charts.BAR, [{ name: "Active users", labels: ["14-20 Sep (implied)", "21-27 Sep"], values: [43, 384] }], Object.assign(chartBase("Active users per week, all sources, all countries"), {
    x: 0.62, y: 1.65, w: 5.2, h: 3.3, barDir: "col", barGapWidthPct: 70, chartColors: [C.soft, C.navy],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 15, dataLabelFontBold: true,
    catAxisLabelFontSize: 11, valAxisMaxVal: 450, showLegend: false, catAxisLineShow: true, catAxisLineColor: C.line,
  }));
  txt(s, [
    { text: "+793%   ", options: { bold: true, color: C.green, fontSize: 22 } },
    { text: "Reported increase; calculation pending verification", options: { color: C.navy, fontSize: 11.5, bold: true } },
  ], { x: 0.62, y: 5.05, w: 5.6, h: 0.45, valign: "middle" });
  txt(s, "Prior week of about 43 users is implied by GA's % change; it is not in the export.", { x: 0.62, y: 5.55, w: 5.4, h: 0.45, fontSize: 10.5, color: C.muted });
  s.addImage({ path: path.join(IMG, "ga_panel.png"), x: 6.55, y: 1.7, w: 6.17, h: 3.96 });
  txt(s, "GA4 home card as captured. The line plots event count, not users.", { x: 6.55, y: 5.75, w: 6.17, h: 0.3, fontSize: 9.5, color: C.muted, italic: true });
}

// ---------- 7. Social and content ----------
{
  const s = contentSlide({
    kicker: "Social and content performance", n: 7,
    title: "Instagram reach sits in three posts; other channels have screenshots but no September exports",
    takeaway: "Two presenter-led reels and one campaign graphic carry the views. Other posts are under 600, so reach depends on a few assets.",
    source: "Instagram grid screenshot (capture date and boost spend unknown); GA4 export 1-28 Sep; LinkedIn export 31 May-28 Aug; X and TikTok from the Performance Analysis deck (screenshots).",
    notes: "Chart uses only view counts visible on the Instagram grid screenshot. Excluded as unverified or inconsistent: 2,271,132 'all-time' views (one reel alone shows 5.3M); format shares, which differ between decks (73.3/25.9/0.7 in the Marketing Update vs 59.2/38.1 in the Performance Analysis deck). LinkedIn: 247 new followers in the last 30 days vs 4 in the prior 30 days is the +6,075%; it is growth in new followers, not in total followers. The LinkedIn export (31 May-28 Aug) shows 498 new followers, 493 of them in June and 4 in August, which contradicts the draft's derived baseline of 284 followers '30 days ago'.",
  });
  const posts = ["Musheera reel", "Ali Sabeel reel", "Mention & Win $500 graphic", "Open account, win $500 graphic", "Trade & Win graphic", "No limits graphic", "How-to video, presenter", "Start trading graphic"];
  const views = [5300000, 1300000, 186000, 26300, 549, 347, 332, 262];
  s.addChart(pres.charts.BAR, [{ name: "Views", labels: posts, values: views }], Object.assign(chartBase("Instagram views per post (as shown on profile grid)"), {
    x: 0.62, y: 1.62, w: 6.2, h: 4.45, barDir: "bar", barGapWidthPct: 45, catAxisOrientation: "maxMin", chartColors: [C.navy],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0',
    valAxisMaxVal: 6500000, showLegend: false, catAxisLineShow: false,
  }));
  const rows = [
    hdr(["Channel", "Latest figure", "Period", "Status"], 10.5),
    row(["Website", "587 sessions; 11 from Instagram", "1-28 Sep", status("Verified")], 10),
    row(["Instagram", "2,116 followers", "Late Sep", status("Reported")], 10),
    row(["LinkedIn", "531 followers; 247 new vs 4 prior (+6,075% new)", "30 days", status("Conflicting")], 10),
    row(["LinkedIn", "559 impressions, +118%", "13-27 Sep vs prior", status("Reported")], 10),
    row(["X", "738 followers, 99 posts; 96.4% Bahrain", "Late Sep", status("Reported")], 10),
    row(["TikTok", "15 followers, 176 likes", "Late Sep", status("Reported")], 10),
  ];
  s.addTable(rows, { x: 7.1, y: 1.68, w: 5.62, colW: [0.95, 2.35, 1.12, 1.2], rowH: 0.56, fontFace: F, margin: [2, 4, 2, 4] });
  txt(s, "LinkedIn total also appears as 586 in an earlier count; export ends 28 Aug.", { x: 7.1, y: 5.72, w: 5.62, h: 0.35, fontSize: 9, color: C.muted, italic: true });
}

// ---------- 8. Campaigns ----------
{
  const s = contentSlide({
    kicker: "Campaign performance", n: 8,
    title: "Mention & Win is a monthly $500 contest; the $5,000 Prize Campaign is a separate, larger campaign",
    takeaway: "Report the two separately. Mention & Win needs sign-up results per round; the $5,000 campaign still needs its T&C sign-off on file.",
    source: "Instagram grid screenshot; 24F campaign briefing; $5,000 campaign approval checklist (20 Jul 2026); raffle T&Cs; Board update 17 Sep 2026 (budget).",
    notes: "Mention & Win (24F briefing): recurring monthly contest across all six GCC countries, approved caption 'Follow, mention, and open an eMYAA account for a chance to win $500'; paid boost via Ignite. Calendar posts 15 Sep and 18 Oct; the team gives 23 Sep to 20 Oct for the current round. Board budget: $1,100 prizes Jun-Aug; $3,200 prizes and ads Sep-Dec. The 26.3K post is the 'open your account, chance to win $500' creative. 1,129 likes on the launch reel is reported, not visible on the screenshot. $5,000 Prize Campaign: approval checklist of 20 Jul shows landing pages built and live, bilingual T&Cs 'sent for approval' with no sign-off recorded, six influencer scripts finalized. Campaign period 1 Aug to 31 Dec 2026, draw 10 Jan 2027; $5,000 prize in the Board plan to year end; launch animation on the October calendar for 5 Oct.",
  });
  txt(s, [
    { text: "Mention & Win   ", options: { bold: true, color: C.navy, fontSize: 16 } },
    { text: "● Live, monthly", options: { bold: true, color: C.green, fontSize: 11 } },
  ], { x: 0.62, y: 1.65, w: 5.9, h: 0.4, valign: "middle" });
  s.addImage({ path: path.join(IMG, "tile_mw_mention.png"), x: 0.62, y: 2.15, w: 1.3, h: 2.03 });
  s.addImage({ path: path.join(IMG, "tile_mw_account.png"), x: 2.07, y: 2.15, w: 1.3, h: 2.03 });
  txt(s, "186K views", { x: 0.62, y: 4.22, w: 1.3, h: 0.28, fontSize: 10.5, bold: true, color: C.navy, align: "center" });
  txt(s, "26.3K views", { x: 2.07, y: 4.22, w: 1.3, h: 0.28, fontSize: 10.5, bold: true, color: C.navy, align: "center" });
  const mw = [["Round", "23 Sep to 20 Oct (per team)"], ["Mechanic", "Follow, mention, open an account"], ["Markets", "All six GCC countries"], ["Budget", "$3,200 prizes and ads, Sep-Dec"], ["Results", "Entries and accounts due after 20 Oct"]];
  mw.forEach((r, i) => {
    txt(s, r[0], { x: 3.62, y: 2.15 + i * 0.5, w: 1.0, h: 0.46, fontSize: 10.5, bold: true, color: C.muted });
    txt(s, r[1], { x: 4.62, y: 2.15 + i * 0.5, w: 2.0, h: 0.46, fontSize: 10.5, color: C.text });
  });
  txt(s, "Launch reel: 1,129 likes (reported)", { x: 0.62, y: 4.6, w: 3, h: 0.3, fontSize: 10, color: C.muted, italic: true });
  s.addShape(pres.shapes.LINE, { x: 6.85, y: 1.7, w: 0, h: 4.35, line: { color: C.line, width: 1 } });
  txt(s, [
    { text: "$5,000 Prize Campaign   ", options: { bold: true, color: C.navy, fontSize: 16 } },
    { text: "● Pages live; T&C sign-off not on file", options: { bold: true, color: C.goldText, fontSize: 11 } },
  ], { x: 7.1, y: 1.65, w: 5.62, h: 0.4, valign: "middle" });
  const rf = [["Dates", "1 Aug to 31 Dec 2026; draw 10 Jan 2027"], ["Mechanic", "1 entry per $500 traded; one winner"], ["Eligibility", "Fully verified, active eMYAA account"]];
  rf.forEach((r, i) => {
    txt(s, r[0], { x: 7.1, y: 2.15 + i * 0.36, w: 1.1, h: 0.34, fontSize: 10.5, bold: true, color: C.muted });
    txt(s, r[1], { x: 8.2, y: 2.15 + i * 0.36, w: 4.5, h: 0.34, fontSize: 10.5, color: C.text });
  });
  s.addChart(pres.charts.BAR, [{ name: "Prize", labels: ["$0.5K-5K", "$5K-15K", "$15K-50K", "$50K-100K", "$100K-500K", "$500K+"], values: [500, 1000, 2000, 3000, 4000, 5000] }], Object.assign(chartBase("Prize if you win, by trading volume Aug-Dec 2026 (USD)"), {
    x: 7.05, y: 3.25, w: 5.67, h: 2.85, barDir: "col", barGapWidthPct: 40, chartColors: [C.gold],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"#,##0', catAxisLabelFontSize: 9, catAxisLabelColor: C.muted,
    valAxisMaxVal: 6000, showLegend: false,
  }));
}

// ---------- 9. Content production ----------
{
  const s = contentSlide({
    kicker: "Content production", n: 9,
    title: "The one-creator-a-month plan is on track: three presenter videos posted, Manal Talal set for 22 October",
    takeaway: "The pipeline is on plan through November. The gap is measurement: Abdulelah Al Harbi has no view count yet.",
    source: "Instagram grid screenshot (videos matched to names by view count); Sep and Oct 2026 content calendars; team rollout plan. No item is awaiting approval per the team.",
    notes: "The how-to video showed 319 views in the previous draft and 332 on the screenshot, so the screenshot is the later capture. Ali Sabeel posted 2 Sep per the September calendar. Manal Talal: brief sent and production started; October calendar V2 date 22 Oct. Hamed Al Bloushi: November per the rollout plan (the September calendar had listed 29 Sep). The same six creators, plus Ammar Alshaikh, have finalized scripts for the $5,000 Prize Campaign (approval checklist).",
  });
  const items = [
    ["JUL", "Musheera", "Posted", "5.3M views", "tile9_musheera.png"],
    ["AUG", "Ali Sabeel", "Posted", "1.3M views", "tile9_ali.png"],
    ["SEP", "Abdulelah Al Harbi", "Posted", "Views not supplied", null],
    ["SEP", "How-to, human presenter", "Posted", "332 views", "tile9_howto.png"],
    ["OCT", "Manal Talal", "In production", "Calendar: 22 Oct", null],
    ["NOV", "Hamed Al Bloushi", "Scheduled", "Per rollout plan", null],
  ];
  const x0 = 0.62, cw = 1.85, gap = 0.2, ly = 2.12;
  s.addShape(pres.shapes.LINE, { x: x0 + cw / 2, y: ly, w: 5 * (cw + gap), h: 0, line: { color: C.soft, width: 2 } });
  items.forEach((it, i) => {
    const x = x0 + i * (cw + gap), cx = x + cw / 2, posted = it[2] === "Posted";
    txt(s, it[0], { x, y: 1.7, w: cw, h: 0.25, fontSize: 10, bold: true, color: C.goldText, align: "center", charSpacing: 2 });
    s.addShape(pres.shapes.OVAL, { x: cx - 0.11, y: ly - 0.11, w: 0.22, h: 0.22, fill: { color: posted ? C.green : C.white }, line: { color: posted ? C.green : C.muted, width: 2 } });
    const iw = 1.55, ih = 2.0, iy = 2.42;
    if (it[4]) s.addImage({ path: path.join(IMG, it[4]), x: cx - iw / 2, y: iy, w: iw, h: ih });
    else {
      s.addShape(pres.shapes.RECTANGLE, { x: cx - iw / 2, y: iy, w: iw, h: ih, fill: { color: C.panel }, line: { color: C.soft, width: 0.75, dashType: "dash" } });
      txt(s, posted ? "No asset in review set" : it[2], { x: cx - iw / 2 + 0.1, y: iy, w: iw - 0.2, h: ih, fontSize: 10, color: C.muted, align: "center", valign: "middle" });
    }
    txt(s, it[1], { x, y: 4.55, w: cw, h: 0.55, fontSize: 12, bold: true, color: C.navy, align: "center" });
    txt(s, [{ text: "● ", options: { color: STATUS[it[2]] } }, { text: it[2], options: { bold: true, color: STATUS[it[2]] } }], { x, y: 5.15, w: cw, h: 0.3, fontSize: 10.5, align: "center" });
    txt(s, it[3], { x, y: 5.47, w: cw, h: 0.3, fontSize: 10.5, color: C.muted, align: "center" });
  });
}

// ---------- 10. Website and analytics ----------
{
  const s = contentSlide({
    kicker: "Website and analytics", n: 10,
    title: "Google Ads brought 70% of September site sessions; English ads engage more, Arabic ads log more key events",
    takeaway: "Do not cut the Arabic campaign on engagement alone. First confirm what the key event measures and what each campaign costs.",
    source: "GA4 Reports snapshot and Google Ads campaigns exports, 1-28 Sep 2026. Referral includes 7 Tag Assistant test sessions; 21 paid sessions sit outside both campaigns.",
    notes: "Corrections vs the previous drafts: referral is 5.5% (32 sessions by GA medium) not 6.6% (the draft grouped ChatGPT's 9 sessions into referral and left out facebook.com's 2); top four sources are 97.4% (572 of 587), not 98.6%; EN/AR session shares are 54.6%/45.4%, not 54.5%/45.2%; EN engagement time per session is 19.9s, not 19s; the period is 1-28 Sep, not 'since 22/23 Sep'; 558 new users is the September total, not 'since campaign launch'. Key events: EN 200 on 212 sessions (0.94 per session), AR 185 on 176 sessions (1.05). The key event is not defined in the export; 400 key events on 409 paid sessions suggests a light action, so do not call it a conversion yet.",
  });
  const k = [["587", "sessions"], ["542", "active users"], ["558", "new users"], ["23s", "avg engagement per active user"]];
  k.forEach((v, i) => {
    const x = 0.62 + i * 3.05;
    txt(s, v[0], { x, y: 1.62, w: 1.15, h: 0.52, fontSize: 26, bold: true, color: C.green, valign: "middle" });
    txt(s, v[1] + ", 1-28 Sep", { x: x + 1.15, y: 1.62, w: 1.8, h: 0.52, fontSize: 10, color: C.muted, valign: "middle" });
  });
  s.addShape(pres.shapes.LINE, { x: 0.62, y: 2.25, w: 12.1, h: 0, line: { color: C.line, width: 0.75 } });
  s.addChart(pres.charts.BAR, [{ name: "Share of sessions", labels: ["Google Ads, cpc (409)", "Direct (92)", "Organic search (39)", "Referral (32)", "ChatGPT (9)", "Not available (6)"], values: [69.7, 15.7, 6.6, 5.5, 1.5, 1.0] }], Object.assign(chartBase("Share of website sessions by source, 1-28 Sep (n = 587)"), {
    x: 0.62, y: 2.38, w: 6.6, h: 3.72, barDir: "bar", barGapWidthPct: 45, catAxisOrientation: "maxMin", chartColors: [C.navy],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"', valAxisMaxVal: 82, showLegend: false,
  }));
  s.addChart(pres.charts.BAR, [
    { name: "English campaign", labels: ["Engagement rate, %", "Key events per 100 sessions"], values: [97.2, 94.3] },
    { name: "Arabic campaign", labels: ["Engagement rate, %", "Key events per 100 sessions"], values: [86.9, 105.1] },
  ], Object.assign(chartBase("Google Ads Saudi campaigns: EN 212 vs AR 176 sessions"), {
    x: 7.5, y: 2.38, w: 5.22, h: 3.72, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60, chartColors: [C.navy, C.gold],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 125, showLegend: true, legendPos: "b",
  }));
}

// ---------- 11. Traffic to accounts ----------
{
  const s = contentSlide({
    kicker: "From traffic to accounts", n: 11,
    title: "Downloads picked up after the paid launch, but onboarding did not: +36 downloads, +3 accounts in six days",
    takeaway: "Paid media is filling the top of the funnel. KYC and onboarding completion is the constraint to fix before spend is scaled.",
    source: "eMYAA Daily Performance Dashboards, all-time cumulative figures on report dates 1, 8, 14, 21 and 27 Sep 2026. Periods are 6 to 7 days.",
    notes: "Net change between dashboard report dates. Downloads = Android + iOS: 373 (1 Sep), 403 (8 Sep), 423 (14 Sep), 437 (21 Sep), 473 (27 Sep). Users onboarded: 48, 49, 52, 57, 60. Funded accounts: 12 to 13 (21 Sep) to 15 (27 Sep). Trades: 54 on 27 Sep. The 28 Sep dashboard (#32) cited by the previous audit shows 480 downloads; it is not in this review set. The Board shows 61 clients 'as at 17 Sep' while the 17 Sep dashboard shows 54; logged as an open item.",
  });
  s.addChart(pres.charts.BAR, [
    { name: "App downloads added", labels: ["1-8 Sep", "8-14 Sep", "14-21 Sep", "21-27 Sep (paid live)"], values: [30, 20, 14, 36] },
    { name: "Users onboarded", labels: ["1-8 Sep", "8-14 Sep", "14-21 Sep", "21-27 Sep (paid live)"], values: [1, 3, 5, 3] },
  ], Object.assign(chartBase("Net new app downloads and onboarded users between dashboard dates"), {
    x: 0.62, y: 1.65, w: 7.9, h: 4.45, barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, chartColors: [C.navy, C.green],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontBold: true, catAxisLabelFontSize: 10.5,
    valAxisMaxVal: 42, showLegend: true, legendPos: "b",
  }));
  const k = [["473", "app downloads", "all time, 27 Sep", C.navy], ["60", "users onboarded", "all time, 27 Sep", C.navy], ["15", "funded accounts", "up from 13 on 21 Sep", C.green], ["54", "trades", "all time, 27 Sep", C.navy]];
  k.forEach((v, i) => {
    const x = 8.95 + (i % 2) * 1.9, y = 1.75 + Math.floor(i / 2) * 2.15;
    card(s, x, y, 1.77, 1.95, i === 2 ? C.green : C.gold);
    txt(s, v[0], { x: x + 0.18, y: y + 0.28, w: 1.45, h: 0.62, fontSize: 28, bold: true, color: v[3] });
    txt(s, v[1], { x: x + 0.18, y: y + 0.95, w: 1.45, h: 0.45, fontSize: 11, bold: true, color: C.navy });
    txt(s, v[2], { x: x + 0.18, y: y + 1.4, w: 1.45, h: 0.45, fontSize: 9.5, color: C.muted });
  });
}

// ---------- 12. Risks and data gaps ----------
{
  const s = contentSlide({
    kicker: "Risks and data gaps", n: 12,
    title: "Seven open items limit what we can claim; three of them block the 20 October paid-media decision",
    takeaway: "Owners and dates are proposed. Items 1, 4 and 5 must close before the paid-media decision.",
    source: "Audit of the previous drafts against GA4 exports, daily dashboards, LinkedIn export, Board update and campaign files. See Percentage-Audit.csv and Source-Traceability.csv.",
    notes: "Gross deposits: the Board shows $12,206 (Malek's figure, verbal); the dashboard figure of $10,029.69 is relabelled from 'Total Deposits' to 'Total AUM' from 20 Sep; the KPI workbook shows $11,195.88 for September. Client count: Board 61 'as at 17 Sep'; dashboards 54 (17 Sep) and 60 (27 Sep); KPI workbook 54. $5,000 campaign: bilingual T&Cs 'sent for approval' on 20 Jul, no sign-off on file. Market Pulse: calendar posts 30 Sep and 29 Oct; the in-app Market Pulse feature is 'in discussion', with Exante verifying data-provider access.",
  });
  const rows = [
    hdr(["#", "Issue", "Impact", "Owner", "Required by", "Status"]),
    row(["1", "Deposits: $12,206 (Board, verbal) vs $10,029.69 on dashboard, relabelled 'AUM' from 20 Sep", "Deposit story rests on an unsourced figure", "Malek", "8 Oct", status("Open")]),
    row(["2", "No September exports for Instagram, LinkedIn, X, TikTok (LinkedIn export ends 28 Aug)", "Social results are screenshot-only", "Marketing", "8 Oct", status("Open")]),
    row(["3", "793% and LinkedIn +118% / +155% not traceable to an export", "Risk of overstating results", "Marketing", "8 Oct", status("Pending evidence")]),
    row(["4", "No spend, cost per key event or LinkedIn leads; key event undefined", "Cannot test Board CAC targets", "Marketing with Ignite", "15 Oct", status("Open")]),
    row(["5", "Clients: Board 61 'as at 17 Sep' vs dashboard 54 (17 Sep) and 60 (27 Sep)", "Board figure may run ahead of source", "Malek", "8 Oct", status("Open")]),
    row(["6", "LinkedIn followers: 531 now, 586 in an earlier count", "Growth claim unreliable", "Marketing with Ignite", "8 Oct", status("Open")]),
    row(["7", "$5,000 campaign T&Cs sent for approval 20 Jul, no sign-off on file; Market Pulse feature in discussion", "Regulatory and content dependency", "Compliance, IT", "15 Oct", status("Pending evidence")]),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 12.1, colW: [0.4, 4.55, 2.75, 1.7, 1.05, 1.65], rowH: 0.52, fontFace: F, margin: [2, 5, 2, 5] });
}

// ---------- 13. Decisions ----------
{
  const s = contentSlide({
    kicker: "CEO decisions and next steps", n: 13,
    title: "Five decisions for today, each with an owner, a date and a measurable outcome",
    takeaway: "With these in place, the next update can report cost per account opened against the Board targets, not just traffic.",
    source: "Proposed by Marketing based on this update. Dates are proposals for CEO confirmation. CAC targets from the Board update, 17 Sep 2026.",
    notes: "Decision 2 follows from the GA4 campaign export: Arabic has a lower engagement rate (86.9% vs 97.2%) but more key events per session (1.05 vs 0.94). Decision 3 follows from the dashboards: +36 downloads but +3 onboarded in the first six days of paid testing. Decision 5: the $5,000 campaign T&Cs were sent for approval on 20 Jul and no sign-off is on file.",
  });
  const num = (t) => ({ text: t, options: { bold: true, color: C.gold, fontSize: 16 } });
  const rows = [
    hdr(["", "Decision required", "Owner", "By", "Expected outcome"], 11.5),
    row([num("1"), "Set 20 Oct as the go or hold point for paid media, judged on cost per account against Board CAC targets", "Marketing with Ignite", "20 Oct", "Remaining budget goes to what converts"], 12),
    row([num("2"), "Keep both Arabic and English ads live until the key event is defined and costed", "Marketing", "10 Oct", "No cut to the version with more key events"], 12),
    row([num("3"), "Fix onboarding drop-off: follow up every new download that has not completed KYC", "Client Experience", "10 Oct", "More of the new downloads become accounts"], 12),
    row([num("4"), "Require a monthly data pack: GA4, social, Google Ads, LinkedIn Ads, deposits", "Marketing and Malek", "First pack 8 Oct", "Every figure in the next update is sourced"], 12),
    row([num("5"), "Confirm sign-off of the $5,000 campaign T&Cs and require a Mention & Win close-out", "Compliance and Marketing", "27 Oct", "Both campaigns on record and measured"], 12),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 12.1, colW: [0.5, 5.2, 2.1, 1.3, 3.0], rowH: [0.42, 0.74, 0.74, 0.74, 0.74, 0.74], fontFace: F, margin: [2, 6, 2, 6] });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote", f));
