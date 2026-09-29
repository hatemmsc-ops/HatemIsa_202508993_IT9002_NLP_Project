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

// ---------- 2. What we did ----------
{
  const s = contentSlide({
    kicker: "Marketing efforts", n: 2,
    title: "September: paid media went into testing, two campaigns are live and four creator videos are out",
    takeaway: "All efforts are on plan. Only paid media carries a decision: $30,000 (Board plan), judged against CAC targets of $167 / $417 / $882.",
    source: "GA4 exports 1-28 Sep 2026; Board update 17 Sep 2026; 24F briefing; $5,000 campaign checklist and T&Cs; Sep and Oct 2026 calendars; Instagram grid screenshot; Performance Analysis deck (X, TikTok).",
    notes: "Paid media (Board plan): $30,000 Sep to Dec from the $62,477 reserve; Google Search Saudi $9K and other GCC $6K; LinkedIn Affluent Saudi $9K and other GCC $6K. Segments: Starter under $1K and Emerging $1K-5K via Google Ads; Affluent $5K-50K via LinkedIn lead gen. CAC targets Starter $167, Emerging $417, Affluent $882. GA4 shows only the two Saudi Google campaigns so far; LinkedIn lead forms keep users on LinkedIn, so no LinkedIn sessions is expected. Mention & Win: monthly $500 contest in all six GCC countries (24F briefing), approved caption 'Follow, mention, and open an eMYAA account for a chance to win $500'; $3,200 prizes and ads budgeted Sep-Dec; 186K and 26.3K views on campaign posts; current round 23 Sep to 20 Oct per team. $5,000 Prize Campaign: 1 Aug to 31 Dec 2026, draw 10 Jan 2027, prize $500 to $5,000 by trading-volume tier, one entry per $500 traded; landing pages live; bilingual T&Cs sent for approval 20 Jul, no sign-off on file. Creators: Musheera (Jul) 5.3M, Ali Sabeel (Aug, posted 2 Sep) 1.3M, Abdulelah Al Harbi (Sep) no view count yet, how-to video with presenter 332, Manal Talal on the October calendar for 22 Oct, Hamed Al Bloushi November per rollout. Social: Instagram 2,116 followers; LinkedIn 531 (586 in an earlier count; 247 new vs 4 in the prior 30 days); X 738 followers; TikTok 15 followers; no September exports for any channel.",
  });
  const st = (label, extra) => ({ text: [
    { text: "● ", options: { color: STATUS[label] || C.muted } },
    { text: label, options: { bold: true, color: STATUS[label] || C.muted } },
    { text: extra ? "  " + extra : "", options: { color: C.text } },
  ] });
  const eff = (a, b) => ({ text: [{ text: a, options: { bold: true, color: C.navy, breakLine: true } }, { text: b, options: { color: C.muted, fontSize: 9 } }] });
  const rows = [
    hdr(["Effort", "Status", "Result so far", "Next checkpoint"], 10.5),
    row([eff("Google Ads", "Starter and Emerging, Saudi"), st("Testing", "since 22 Sep"), "409 sessions, 70% of September site traffic", "Spend and cost per account, 20 Oct"], 10.5),
    row([eff("LinkedIn Ads", "Affluent lead generation, Saudi"), st("Testing", "since 22 Sep"), "No results reported yet", "Leads and cost per lead, 20 Oct"], 10.5),
    row([eff("Mention & Win", "$500 monthly contest, GCC"), st("Live", "round to 20 Oct"), "186K views on the campaign post", "Entries and accounts opened"], 10.5),
    row([eff("$5,000 Prize Campaign", "1 Aug to 31 Dec, draw 10 Jan"), st("Live", "pages up"), "T&C sign-off not on file", "Compliance sign-off"], 10.5),
    row([eff("Creator videos", "one a month"), st("Posted", "4 so far"), "5.3M and 1.3M views; Manal Talal 22 Oct, Hamed Nov", "Abdulelah view count"], 10.5),
    row([eff("Social channels", "Instagram, LinkedIn, X, TikTok"), st("Pending evidence", ""), "Followers: 2,116 / 531 / 738 / 15", "September exports, 8 Oct"], 10.5),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 8.75, colW: [2.2, 1.85, 2.65, 2.05], rowH: [0.34, 0.66, 0.66, 0.66, 0.66, 0.66, 0.66], fontFace: F, margin: [2, 5, 2, 5] });
  const tiles = [["tile9_musheera.png", "Musheera", "5.3M"], ["tile9_ali.png", "Ali Sabeel", "1.3M"], ["tile9_abdulelah.png", "Abdulelah Al Harbi", "views pending"], ["tile9_howto.png", "How-to, presenter", "332"]];
  tiles.forEach((t, i) => {
    const x = 9.7 + (i % 2) * 1.55, y = 1.72 + Math.floor(i / 2) * 2.2, iw = 1.35, ih = 1.62;
    s.addImage({ path: path.join(IMG, t[0]), x, y, w: iw, h: ih });
    txt(s, [{ text: t[1], options: { bold: true, color: C.navy, breakLine: true } }, { text: t[2] + (t[2].match(/\d$/) ? " views" : ""), options: { color: C.muted } }], { x: x - 0.1, y: y + ih + 0.04, w: iw + 0.2, h: 0.46, fontSize: 9, align: "center" });
  });
}

// ---------- 3. What it delivered ----------
{
  const s = contentSlide({
    kicker: "Results", n: 3,
    title: "Paid search lifted traffic and downloads; onboarding did not follow",
    takeaway: "The constraint is onboarding, not traffic. Fix KYC completion before spend is scaled.",
    source: "GA4 exports 1-28 Sep 2026; eMYAA Daily Performance Dashboards, cumulative figures on report dates 1, 8, 14, 21 and 27 Sep 2026 (periods of 6 to 7 days).",
    notes: "Traffic: 587 sessions 1-28 Sep; google / cpc 409 (69.7%), direct 92, organic 39, referral 32 (incl. 7 Tag Assistant test sessions), ChatGPT 9, not available 6. 403 of 542 active users (74.4%) had Google Ads as first source. The GA home card reports 384 active users for 21-27 Sep, +793% vs the previous week; the prior week (about 43) is implied, not exported, so the figure stays 'reported'. Downloads (Android + iOS) on dashboard dates: 373, 403, 423, 437, 473; onboarded 48, 49, 52, 57, 60; funded accounts 13 on 21 Sep, 15 on 27 Sep; 54 trades. English vs Arabic Google Ads (1-28 Sep): engagement 97.2% vs 86.9%, key events per session 0.94 vs 1.05; the key event is not defined, so keep both live until it is defined and costed.",
  });
  const stats = [
    ["70%", "of September site sessions came from Google Ads", "409 of 587 sessions, 1-28 Sep", C.navy],
    ["+36", "app downloads in the six days after launch", "vs +14 the week before", C.navy],
    ["+3", "users onboarded in the same six days", "60 onboarded, 15 funded, as of 27 Sep", C.red],
  ];
  stats.forEach((v, i) => {
    const y = 1.72 + i * 1.47;
    card(s, 0.62, y, 3.9, 1.3, i === 2 ? C.red : C.gold);
    txt(s, v[0], { x: 0.85, y: y + 0.18, w: 1.3, h: 0.8, fontSize: 34, bold: true, color: v[3], valign: "middle" });
    txt(s, v[1], { x: 2.15, y: y + 0.2, w: 2.25, h: 0.6, fontSize: 11.5, bold: true, color: C.navy });
    txt(s, v[2], { x: 2.15, y: y + 0.82, w: 2.25, h: 0.35, fontSize: 9.5, color: C.muted });
  });
  s.addChart(pres.charts.BAR, [
    { name: "App downloads added", labels: ["1-8 Sep", "8-14 Sep", "14-21 Sep", "21-27 Sep (paid live)"], values: [30, 20, 14, 36] },
    { name: "Users onboarded", labels: ["1-8 Sep", "8-14 Sep", "14-21 Sep", "21-27 Sep (paid live)"], values: [1, 3, 5, 3] },
  ], Object.assign(chartBase("Net new app downloads and onboarded users between dashboard dates"), {
    x: 4.85, y: 1.65, w: 7.87, h: 4.45, barDir: "col", barGrouping: "clustered", barGapWidthPct: 55, chartColors: [C.navy, C.green],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontBold: true, catAxisLabelFontSize: 10.5,
    valAxisMaxVal: 42, showLegend: true, legendPos: "b",
  }));
}

// ---------- 4. Decisions and open items ----------
{
  const s = contentSlide({
    kicker: "Decisions and open items", n: 4,
    title: "Five decisions for today; three open data items must close before 20 October",
    takeaway: "Decision 4 closes most of the open items. The three marked 'Blocks 20 Oct' must close before the paid-media call.",
    source: "Proposed by Marketing based on this update; owners and dates are proposals for CEO confirmation. CAC targets from the Board update, 17 Sep 2026. Detail in Open-Items.md and Percentage-Audit.csv.",
    notes: "Decision 2 follows from the GA4 campaign export: Arabic has a lower engagement rate (86.9% vs 97.2%) but more key events per session (1.05 vs 0.94). Decision 3 follows from the dashboards: +36 downloads but +3 onboarded in the first six days of paid testing. Open items: deposits $12,206 (Board, verbal from Malek) vs $10,029.69 on the dashboard (relabelled 'Total AUM' from 20 Sep) vs $11,195.88 in the KPI workbook; clients 61 on the Board 'as at 17 Sep' vs 54 on the 17 Sep dashboard and 60 on 27 Sep; no Ignite spend, cost per key event or LinkedIn lead data, and the key event is undefined; no September social exports; the 793% active-user rise (GA card, prior week implied) and LinkedIn +118% / +155% are not traceable to an export; LinkedIn followers 531 vs 586; $5,000 campaign T&Cs sent for approval on 20 Jul with no sign-off on file. Market Pulse: in-app feature in discussion (Exante data access).",
  });
  const num = (t) => ({ text: t, options: { bold: true, color: C.gold, fontSize: 14 } });
  const rows = [
    hdr(["", "Decision", "Owner", "By"], 11),
    row([num("1"), "Go or hold on paid media, judged on cost per account vs Board CAC targets", "Marketing with Ignite", "20 Oct"], 11),
    row([num("2"), "Keep Arabic and English ads live until the key event is defined and costed", "Marketing", "10 Oct"], 11),
    row([num("3"), "Follow up every new download that has not completed KYC", "Client Experience", "10 Oct"], 11),
    row([num("4"), "Monthly data pack: GA4, social, Google and LinkedIn Ads, deposits", "Marketing and Malek", "8 Oct"], 11),
    row([num("5"), "Sign off the $5,000 campaign T&Cs; close out Mention & Win", "Compliance and Marketing", "27 Oct"], 11),
  ];
  s.addTable(rows, { x: 0.62, y: 1.68, w: 7.3, colW: [0.4, 4.1, 1.85, 0.95], rowH: [0.36, 0.78, 0.78, 0.78, 0.78, 0.78], fontFace: F, margin: [2, 5, 2, 5] });
  const blk = { text: "Blocks 20 Oct", options: { bold: true, color: C.red } };
  const opn = { text: "Open", options: { bold: true, color: C.goldText } };
  const orows = [
    hdr(["Open item", "Status"], 10.5),
    row(["Deposits: three different figures; Board $12,206 is verbal", blk], 9.5),
    row(["Clients: Board 61 vs dashboard 54 on 17 Sep", blk], 9.5),
    row(["No ad spend, LinkedIn leads or key-event definition", blk], 9.5),
    row(["No September exports for social channels", opn], 9.5),
    row(["793% and LinkedIn +118% / +155% not exported", opn], 9.5),
    row(["LinkedIn followers: 531 vs 586", opn], 9.5),
    row(["$5,000 campaign T&C sign-off not on file", opn], 9.5),
  ];
  s.addTable(orows, { x: 8.2, y: 1.68, w: 4.52, colW: [3.3, 1.22], rowH: [0.36, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.54], fontFace: F, margin: [2, 5, 2, 5] });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote", f));
