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


// ---- Shape-drawn charts (import cleanly into Canva and Google Slides, like the Board deck's charts) ----
function legend(s, items, x, y, w) {
  let cx = x + w / 2 - items.reduce((a, it) => a + 0.3 + it[0].length * 0.075, 0) / 2;
  items.forEach(([name, color]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: y + 0.08, w: 0.14, h: 0.14, fill: { color }, line: { color, width: 0 } });
    txt(s, name, { x: cx + 0.2, y, w: name.length * 0.075 + 0.1, h: 0.3, fontSize: 10, color: C.text, valign: "middle" });
    cx += 0.3 + name.length * 0.075;
  });
}
function colBars(s, o) {
  const { x, y, w, h, title, cats, series, max } = o;
  txt(s, title, { x, y, w, h: 0.3, fontSize: 11, color: C.navy, align: "center" });
  const top = y + 0.55, bottom = y + h - 0.75, ph = bottom - top, gw = w / cats.length;
  const bw = (gw * 0.62) / series.length;
  s.addShape(pres.shapes.LINE, { x, y: bottom, w, h: 0, line: { color: C.soft, width: 1 } });
  cats.forEach((c, i) => {
    const gx = x + i * gw + (gw - bw * series.length) / 2;
    series.forEach((se, k) => {
      const v = se.values[i], bh = Math.max((v / max) * ph, 0.02), bx = gx + k * bw;
      s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.02, y: bottom - bh, w: bw - 0.04, h: bh, fill: { color: se.color }, line: { color: se.color, width: 0 } });
      txt(s, String(v), { x: bx - 0.1, y: bottom - bh - 0.3, w: bw + 0.2, h: 0.28, fontSize: 11, bold: true, color: C.navy, align: "center", valign: "bottom" });
    });
    txt(s, c, { x: x + i * gw, y: bottom + 0.08, w: gw, h: 0.3, fontSize: 10.5, color: C.text, align: "center" });
  });
  legend(s, series.map((se) => [se.name, se.color]), x, y + h - 0.32, w);
}
function lineShape(s, o) {
  const { x, y, w, h, title, cats, values, min, max, step, marker } = o;
  txt(s, title, { x, y, w, h: 0.3, fontSize: 11, color: C.navy, align: "center" });
  const lx = x + 0.5, top = y + 0.5, bottom = y + h - 0.6, pw = w - 0.6, ph = bottom - top;
  const Y = (v) => bottom - ((v - min) / (max - min)) * ph;
  for (let v = min; v <= max; v += step) {
    s.addShape(pres.shapes.LINE, { x: lx, y: Y(v), w: pw, h: 0, line: { color: v === min ? C.soft : "E6EAF2", width: v === min ? 1 : 0.75 } });
    txt(s, String(v), { x: x, y: Y(v) - 0.12, w: 0.42, h: 0.24, fontSize: 9, color: C.muted, align: "right", valign: "middle" });
  }
  const X = (i) => lx + 0.15 + (i * (pw - 0.3)) / (values.length - 1);
  if (marker !== undefined) {
    s.addShape(pres.shapes.LINE, { x: X(marker), y: top - 0.05, w: 0, h: bottom - top + 0.05, line: { color: C.gold, width: 1.25, dashType: "dash" } });
  }
  for (let i = 0; i < values.length - 1; i++) {
    const x1 = X(i), x2 = X(i + 1), y1 = Y(values[i]), y2 = Y(values[i + 1]);
    s.addShape(pres.shapes.LINE, { x: x1, y: Math.min(y1, y2), w: x2 - x1, h: Math.abs(y2 - y1), flipV: y2 < y1, line: { color: C.navy, width: 2.5 } });
  }
  values.forEach((v, i) => {
    s.addShape(pres.shapes.OVAL, { x: X(i) - 0.05, y: Y(v) - 0.05, w: 0.1, h: 0.1, fill: { color: C.navy }, line: { color: C.navy, width: 0 } });
    txt(s, cats[i], { x: X(i) - 0.3, y: bottom + 0.06, w: 0.6, h: 0.24, fontSize: 8.5, color: i === marker ? C.goldText : C.muted, bold: i === marker, align: "center" });
  });
  if (marker !== undefined) txt(s, "Ads live", { x: X(marker) - 0.5, y: bottom + 0.28, w: 1.0, h: 0.22, fontSize: 8.5, bold: true, color: C.goldText, align: "center" });
}
function divBars(s, o) {
  const { x, y, w, h, title, cats, values, min, max, posColor, negColor, posName, negName } = o;
  txt(s, title, { x, y, w, h: 0.3, fontSize: 11, color: C.navy, align: "center" });
  const labW = 1.95, px = x + labW + 0.1, pw = w - labW - 0.15, top = y + 0.45, rowH = (h - 0.45 - 0.45) / cats.length;
  const X = (v) => px + ((v - min) / (max - min)) * pw, zx = X(0);
  s.addShape(pres.shapes.LINE, { x: zx, y: top - 0.05, w: 0, h: rowH * cats.length + 0.05, line: { color: C.soft, width: 1 } });
  cats.forEach((c, i) => {
    const ry = top + i * rowH, bh = rowH * 0.62, by = ry + (rowH - bh) / 2, v = values[i];
    txt(s, c, { x, y: ry, w: labW, h: rowH, fontSize: 10.5, color: C.text, align: "right", valign: "middle" });
    const bx = v >= 0 ? zx : X(v), bw = Math.abs(X(v) - zx), col = v >= 0 ? posColor : negColor;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: col }, line: { color: col, width: 0 } });
    const lab = (v >= 0 ? "+" : "") + v.toFixed(1) + "%";
    if (v >= 0) txt(s, lab, { x: bx + bw + 0.05, y: by, w: 0.8, h: bh, fontSize: 10, bold: true, color: C.navy, valign: "middle" });
    else txt(s, lab, { x: bx - 0.85, y: by, w: 0.8, h: bh, fontSize: 10, bold: true, color: C.goldText, align: "right", valign: "middle" });
  });
  legend(s, [[posName, posColor], [negName, negColor]], x, y + h - 0.32, w);
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
  colBars(s, {
    x: 4.85, y: 1.65, w: 7.87, h: 4.45, max: 40,
    title: "Net new app downloads and onboarded users between dashboard dates",
    cats: ["1-8 Sep", "8-14 Sep", "14-21 Sep", "21-27 Sep (paid live)"],
    series: [{ name: "App downloads added", color: C.navy, values: [30, 20, 14, 36] }, { name: "Users onboarded", color: C.green, values: [1, 3, 5, 3] }],
  });
}

// ---------- 4. Google Ads and installs ----------
{
  const s = contentSlide({
    kicker: "Google Ads and installs", n: 4,
    title: "Google Ads went live on 22 September; installs then grew at their fastest weekly pace of the month",
    takeaway: "Paid search is moving installs, but installs are not yet tracked back to the ads. Connect ad clicks to the app link so cost per install can be reported.",
    source: "App installs: eMYAA Daily Performance Dashboards (Android + iOS, cumulative, on report dates 1-27 Sep 2026). Ads: GA4 Reports snapshot and Google Ads campaigns exports, 1-28 Sep 2026.",
    notes: "Installs (Android + iOS, cumulative) on dashboard report dates: 1 Sep 373, 2 Sep 389, 3 Sep 395, 8 Sep 403, 9 Sep 410, 10 Sep 414, 13 Sep 419, 14 Sep 423, 16 Sep 430, 17 Sep 432, 20 Sep 437, 21 Sep 437, 22 Sep 452, 23 Sep 454, 24 Sep 455, 27 Sep 473. Report dates are irregular, so the line is spaced by report, not by calendar day. Pace: +36 in the six days 21-27 Sep (6 a day) vs +30 in 1-8 Sep (4.3 a day, when the Ali Sabeel video and baseline sponsor ads were running) and +14 in 14-21 Sep. Split 21-27 Sep: Android +17, iOS +19. Funnel: 409 google / cpc sessions and 400 key events in GA4 (1-28 Sep); the key event is not defined in the export. Installs and onboarding come from the platform dashboard and cover all sources; there is no link yet from an ad click to an install. The app link (onelink.to/gpbwrg) already exists and can carry campaign tags. Onboarded users 57 to 60 over 21-27 Sep.",
  });
  const labels = ["1 Sep", "2", "3", "8", "9", "10", "13", "14", "16", "17", "20", "21", "22 Sep", "23", "24", "27 Sep"];
  const installs = [373, 389, 395, 403, 410, 414, 419, 423, 430, 432, 437, 437, 452, 454, 455, 473];
  lineShape(s, {
    x: 0.62, y: 1.65, w: 7.35, h: 4.45, min: 360, max: 480, step: 30, marker: 12,
    title: "App installs, cumulative (Android + iOS), by dashboard report date",
    cats: labels, values: installs,
  });
  // callouts on the chart
  txt(s, [{ text: "+36 in 6 days", options: { bold: true, color: C.green, breakLine: true } }, { text: "21-27 Sep, after launch", options: { color: C.muted } }], { x: 6.6, y: 3.45, w: 1.35, h: 0.7, fontSize: 10.5 });
  txt(s, [{ text: "+14 in 7 days", options: { bold: true, color: C.navy, breakLine: true } }, { text: "14-21 Sep, before launch", options: { color: C.muted } }], { x: 4.45, y: 4.05, w: 2.3, h: 0.5, fontSize: 10.5 });

  // funnel diagram
  const fx = 8.3, fw = 4.42;
  txt(s, "FROM AD TO ACCOUNT", { x: fx, y: 1.62, w: fw, h: 0.25, fontSize: 9, bold: true, color: C.goldText, charSpacing: 2 });
  const steps = [
    ["409", "Google Ads sessions", "GA4, 1-28 Sep", C.navy, 4.42],
    ["400", "key events on those sessions", "event not yet defined", C.navy, 3.9],
    ["+36", "app installs, 21-27 Sep", "all sources", C.gold, 3.38],
    ["+3", "onboarded, 21-27 Sep", "all sources", C.red, 2.86],
  ];
  steps.forEach((st, i) => {
    const y = 1.98 + i * 0.98 + (i >= 2 ? 0.22 : 0), w = st[4], x = fx + (fw - w) / 2;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.78, fill: { color: i < 2 ? C.navy : C.white }, line: { color: st[3], width: 1.5 } });
    txt(s, st[0], { x: x + 0.15, y, w: 0.95, h: 0.78, fontSize: 22, bold: true, color: i < 2 ? C.white : st[3], valign: "middle" });
    txt(s, [{ text: st[1], options: { bold: true, breakLine: true } }, { text: st[2], options: { fontSize: 8.5 } }], { x: x + 1.1, y, w: w - 1.2, h: 0.78, fontSize: 10, color: i < 2 ? C.white : C.navy, valign: "middle" });
  });
  s.addShape(pres.shapes.LINE, { x: fx, y: 3.97, w: fw, h: 0, line: { color: C.red, width: 1, dashType: "dash" } });
  txt(s, "Tracking gap: installs not yet linked to ad clicks", { x: fx, y: 3.99, w: fw, h: 0.2, fontSize: 8.5, bold: true, color: C.red, align: "center" });
}

// ---------- 5. English vs Arabic ----------
{
  const s = contentSlide({
    kicker: "English vs Arabic", n: 5,
    title: "English ads lead Arabic on traffic and engagement; Arabic still logs more key events per session",
    takeaway: "English is ahead on 8 of 9 measures. Refresh the Arabic creative to lift engagement, but keep it live: it converts slightly better per session.",
    source: "GA4 Google Ads campaigns export, 1-28 Sep 2026: 'Website conversion - Saudi - English' vs '... - Arabic'. Campaigns live from 22 Sep. Key event not defined in the export.",
    notes: "English vs Arabic, 1-28 Sep: sessions 212 vs 176 (+20.5%); active users 208 vs 175 (+18.9%); engaged sessions 206 vs 153 (+34.6%); engagement rate 97.2% vs 86.9% (+10.2 points, +11.8% relative); average engagement time per session 19.9s vs 16.0s (+24.2%); events per session 6.49 vs 5.44 (+19.2%); event count 1,376 vs 958 (+43.6%); key events 200 vs 185 (+8.1%); key events per session 0.94 vs 1.05 (Arabic ahead by 10.2%). Early data from about one week of spend, and spend per campaign is not in the export, so this compares engagement, not cost efficiency. The previous draft's claim that English leads 'on every measured metric' is not correct.",
  });
  const cardRows = [
    ["English", C.navy, [["212", "sessions"], ["97.2%", "engagement rate"], ["19.9s", "per session"], ["0.94", "key events per session"]]],
    ["Arabic", C.gold, [["176", "sessions"], ["86.9%", "engagement rate"], ["16.0s", "per session"], ["1.05", "key events per session"]]],
  ];
  cardRows.forEach((c, i) => {
    const x = 0.62, y = 1.72 + i * 2.2, w = 4.6, h = 2.0;
    card(s, x, y, w, h, c[1]);
    txt(s, c[0].toUpperCase() + " CAMPAIGN", { x: x + 0.25, y: y + 0.2, w: 3, h: 0.25, fontSize: 9, bold: true, color: C.goldText, charSpacing: 2 });
    c[2].forEach((m, j) => {
      const mx = x + 0.25 + (j % 2) * 2.2, my = y + 0.55 + Math.floor(j / 2) * 0.7;
      const win = (j < 3 && i === 0) || (j === 3 && i === 1);
      txt(s, m[0], { x: mx, y: my, w: 2.1, h: 0.38, fontSize: 20, bold: true, color: win ? C.green : C.navy });
      txt(s, m[1], { x: mx, y: my + 0.36, w: 2.1, h: 0.25, fontSize: 9, color: C.muted });
    });
  });
  const metrics = ["Sessions", "Active users", "Engaged sessions", "Engagement rate", "Time per session", "Events per session", "Total events", "Key events", "Key events per session"];
  const diff = [20.5, 18.9, 34.6, 11.8, 24.2, 19.2, 43.6, 8.1, -10.2];
  divBars(s, {
    x: 5.55, y: 1.65, w: 7.17, h: 4.45, min: -20, max: 55, cats: metrics, values: diff,
    title: "English relative to Arabic, % difference, 1-28 Sep",
    posColor: C.navy, negColor: C.gold, posName: "English ahead", negName: "Arabic ahead",
  });
}

// ---------- 6. Decisions and open items ----------
{
  const s = contentSlide({
    kicker: "Decisions and open items", n: 6,
    title: "Five decisions for today; one open data item blocks the 20 October call",
    takeaway: "Ignite must supply spend and lead data before 20 October. Decision 4 closes most of the other open items.",
    source: "Proposed by Marketing based on this update; owners and dates are proposals for CEO confirmation. CAC targets from the Board update, 17 Sep 2026. Detail in Open-Items.md and Percentage-Audit.csv.",
    notes: "Decision 2 follows from the GA4 campaign export: Arabic has a lower engagement rate (86.9% vs 97.2%) but more key events per session (1.05 vs 0.94). Decision 3 follows from the dashboards: +36 downloads but +3 onboarded in the first six days of paid testing. Gross deposits: $12,206 is the total recorded on the platform to date, confirmed by management on 29 Sep (the dashboard's $10,029.69 is assets under management, a different measure; the Board shows both). Open items: the Board figures (61 clients, $12,206) are as at 28 Sep, and Board slide 3 wrongly says 'as at 17 September', corrected in the separate Board update file; no Ignite spend, cost per key event or LinkedIn lead data, and the key event is undefined; no September social exports; the 793% active-user rise (GA card, prior week implied) and LinkedIn +118% / +155% are not traceable to an export; LinkedIn followers 531 vs 586; $5,000 campaign T&Cs sent for approval on 20 Jul with no sign-off on file. Market Pulse: in-app feature in discussion (Exante data access).",
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
    row(["No ad spend, LinkedIn leads or key-event definition", blk], 9.5),
    row(["Deposits: $12,206 to date, confirmed 29 Sep", { text: "Confirmed", options: { bold: true, color: C.green } }], 9.5),
    row(["No September exports for social channels", opn], 9.5),
    row(["793% and LinkedIn +118% / +155% not exported", opn], 9.5),
    row(["LinkedIn followers: 531 vs 586", opn], 9.5),
    row(["$5,000 campaign T&C sign-off not on file", opn], 9.5),
    row(["Board slide 3 date: 17 Sep should read 28 Sep", { text: "Fixed", options: { bold: true, color: C.green } }], 9.5),
  ];
  s.addTable(orows, { x: 8.2, y: 1.68, w: 4.52, colW: [3.3, 1.22], rowH: [0.36, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.54], fontFace: F, margin: [2, 5, 2, 5] });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote", f));
