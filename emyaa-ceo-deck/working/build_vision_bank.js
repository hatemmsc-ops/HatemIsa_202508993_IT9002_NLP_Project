// eMYAA x Vision Bank: introductory deck on eMYAA's Saudi market plans (Board visual language).
// Includes costs (Google/LinkedIn budget from the Board update; Jodel and outdoor from Nadher Media proposals).
// No client or deposit figures.
// Run from emyaa-ceo-deck/: NODE_PATH=<dir with pptxgenjs> node working/build_vision_bank.js
// then: python3 working/embed_fonts.py output/vision-bank/eMYAA_Saudi_Market_Plan_-_Vision_Bank.pptx
const path = require("path");
const pptxgen = require("pptxgenjs");

const IMG = path.join(__dirname, "charts");
const OUT = path.join(__dirname, "..", "output", "vision-bank", "eMYAA_Saudi_Market_Plan_-_Vision_Bank.pptx");
const C = {
  navy: "132257", green: "1BA97C", gold: "D4AF37", goldText: "9A7B1F", panel: "EEF1F8",
  line: "D9DFEA", soft: "C6CFE4", text: "1F2A44", muted: "6B7590", foot: "8A93A8", white: "FFFFFF", blue: "3C61B6",
};
const F = "Exo 2";
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "eMYAA Marketing, Ajyad Capital";
pres.title = "eMYAA in Saudi Arabia - Market Entry Plan";

function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: F, margin: 0, valign: "top", color: C.text }, o));
}
function box(s, x, y, w, h, fill, line) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: line ? 0.75 : 0 } });
}
function card(s, x, y, w, h, top) {
  box(s, x, y, w, h, C.white, C.line);
  box(s, x, y, w, 0.045, top || C.navy);
}
let page = 1;
function content(kicker, title, notes) {
  const s = pres.addSlide();
  page += 1;
  s.background = { color: C.white };
  s.addImage({ path: path.join(IMG, "board_header.png"), x: 0, y: 0, w: 13.333, h: 1.42 });
  box(s, 0, 1.42, 13.333, 0.03, C.gold);
  txt(s, kicker.toUpperCase(), { x: 0.62, y: 0.3, w: 10, h: 0.22, fontSize: 9.5, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, title, { x: 0.62, y: 0.52, w: 10.3, h: 0.82, fontSize: 22, bold: true, color: C.white, valign: "middle" });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 11.27, y: 0.5, w: 1.45, h: 0.42 });
  txt(s, "eMYAA Trading Platform  |  Saudi Market Plan", { x: 0.62, y: 7.1, w: 6, h: 0.22, fontSize: 8, color: C.foot });
  txt(s, String(page), { x: 10.72, y: 7.1, w: 0.5, h: 0.22, fontSize: 8, color: C.foot, align: "right" });
  s.addImage({ path: path.join(IMG, "emyaa_navy.png"), x: 11.6, y: 6.92, w: 1.12, h: 0.4 });
  if (notes) s.addNotes(notes);
  return s;
}
function stat(s, x, y, w, value, label, sub, color) {
  card(s, x, y, w, 1.35, color);
  txt(s, value, { x: x + 0.2, y: y + 0.18, w: w - 0.4, h: 0.55, fontSize: 26, bold: true, color: color || C.navy });
  txt(s, label, { x: x + 0.2, y: y + 0.75, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, color: C.navy });
  if (sub) txt(s, sub, { x: x + 0.2, y: y + 1.02, w: w - 0.4, h: 0.28, fontSize: 9, color: C.muted });
}
function bullets(s, items, o) {
  txt(s, items.map((t, i) => {
    const runs = Array.isArray(t) ? t : [{ text: t }];
    return runs.map((r, k) => ({ text: r.text, options: Object.assign({ bullet: k === 0 ? { indent: 14 } : undefined, breakLine: k === runs.length - 1 && i < items.length - 1, paraSpaceAfter: 6 }, r.o || {}) }));
  }).flat(), Object.assign({ fontSize: 12, color: C.text }, o));
}
const B = (text) => ({ text, o: { bold: true, color: C.navy } });

// ---------- 1. Title ----------
{
  const s = pres.addSlide();
  s.background = { path: path.join(IMG, "board_title_bg.png") };
  s.addImage({ path: path.join(IMG, "emyaa_white.png"), x: 0.75, y: 0.75, w: 1.75, h: 0.62 });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 10.67, y: 0.75, w: 1.92, h: 0.56 });
  txt(s, "eMYAA in Saudi Arabia", { x: 0.75, y: 2.75, w: 11.9, h: 0.95, fontSize: 42, bold: true, color: C.white });
  txt(s, "MARKET ENTRY PLAN", { x: 0.75, y: 3.75, w: 11, h: 0.45, fontSize: 20, bold: true, color: C.gold, charSpacing: 1 });
  txt(s, "Channels, timing and cost  |  Vision Bank  |  October 2026", { x: 0.75, y: 4.3, w: 11, h: 0.4, fontSize: 16, color: C.white });
  txt(s, "Confidential. Prepared by eMYAA Marketing, Ajyad Capital.", { x: 0.75, y: 6.75, w: 11, h: 0.3, fontSize: 10, color: C.soft });
}

// table helper: rows of cells; first row is the header
function table(s, x, y, colW, rows, o) {
  o = o || {};
  const rh = o.rowH || 0.55, fs = o.fontSize || 11.5;
  let cx;
  rows.forEach((r, i) => {
    cx = x;
    const yy = y + (i === 0 ? 0 : 0.45 + (i - 1) * rh), h = i === 0 ? 0.45 : rh;
    r.forEach((c, j) => {
      const cell = typeof c === "string" ? { t: c } : c;
      box(s, cx, yy, colW[j], h, i === 0 ? C.navy : (cell.fill || (i % 2 ? C.white : C.panel)), i === 0 ? null : C.line);
      txt(s, cell.t, { x: cx + 0.12, y: yy, w: colW[j] - 0.22, h, fontSize: i === 0 ? 10.5 : fs, bold: i === 0 || cell.b || j === 0,
        color: i === 0 ? C.white : (cell.c || (j === 0 ? C.navy : C.text)), valign: "middle", align: cell.a || "left" });
      cx += colW[j];
    });
  });
}
const tag = (t, c) => ({ t, b: true, c });

// ---------- 2. eMYAA at a glance ----------
{
  const s = content("Who we are", "eMYAA: a Shari'ah-compliant trading app for GCC investors",
    "Ajyad Capital is licensed by the Central Bank of Bahrain as an Islamic investment firm, Category 1. eMYAA went live in May 2026.");
  const facts = [
    ["Shari'ah compliant", "Automated Shari'ah screening and purification calculator"],
    ["Arabic and English", "Full bilingual app with AI support"],
    ["USD 50 to start", "No account fees"],
    ["Four markets", "US, Saudi (Tadawul), Dubai (DFM), Hong Kong"],
  ];
  facts.forEach((f, i) => {
    const x = 0.62 + (i % 2) * 6.1, y = 1.85 + Math.floor(i / 2) * 1.6;
    card(s, x, y, 5.95, 1.4, i % 2 ? C.gold : C.navy);
    txt(s, f[0], { x: x + 0.3, y: y + 0.28, w: 5.4, h: 0.45, fontSize: 19, bold: true, color: C.navy });
    txt(s, f[1], { x: x + 0.3, y: y + 0.8, w: 5.4, h: 0.4, fontSize: 13, color: C.muted });
  });
  box(s, 0.62, 5.25, 12.1, 0.8, C.panel);
  box(s, 0.62, 5.25, 0.06, 0.8, C.gold);
  txt(s, [{ text: "Licensed  ", options: { bold: true, color: C.navy } }, { text: "Powered by Ajyad Capital, an Islamic investment firm (Category 1) regulated by the Central Bank of Bahrain. Live since May 2026.", options: { color: C.text } }],
    { x: 0.9, y: 5.25, w: 11.7, h: 0.8, fontSize: 13, valign: "middle" });
}

// ---------- 3. The Saudi plan and its cost ----------
{
  const s = content("Saudi market plan", "The Saudi plan at a glance: channels, timing and cost",
    "Google Ads and LinkedIn: $30,000 approved for Sep to Dec 2026 from the reserved budget (Board update, 17 Sep 2026): Google Search Saudi $9,000, LinkedIn Affluent Saudi $9,000, Google and LinkedIn for Qatar, UAE, Kuwait, Bahrain and Oman $6,000 each. Jodel: Nadher Media proposal, 16 Sep 2026, USD 3,862.50 (SAR 14,496.25) + VAT. Outdoor: Nadher Media proposal, 23 Sep 2026, 8 weeks, SAR 168,000 + VAT (about USD 44,800 at 3.75). Jodel and outdoor are proposals and not yet approved.");
  table(s, 0.62, 1.75, [3.2, 2.6, 1.9, 2.4, 2.0], [
    ["Channel", "Target", "Timing", "Cost", "Status"],
    [{ t: "Google Search ads, Saudi" }, "Starter and Emerging", "Sep to Dec", { t: "USD 9,000", b: true }, tag("Live since 22 Sep", C.green)],
    [{ t: "LinkedIn lead generation, Saudi" }, "Affluent", "Oct to Dec", { t: "USD 9,000", b: true }, tag("Launching Oct", C.navy)],
    [{ t: "Google and LinkedIn, rest of GCC" }, "All segments", "Nov to Dec", { t: "USD 12,000", b: true }, tag("Planned", C.navy)],
    [{ t: "Jodel launch burst, all KSA" }, "Students, ages 21 to 35", "7 days", { t: "USD 3,862.50 + VAT", b: true }, tag("Proposed", C.goldText)],
    [{ t: "Outdoor and elevator screens, KSA" }, "Professionals, high-end", "8 weeks", { t: "SAR 168,000 + VAT", b: true }, tag("Proposed", C.goldText)],
    [{ t: "Saudi creator videos" }, "Starter", "Monthly", "Per creator", tag("Live", C.green)],
  ], { rowH: 0.6, fontSize: 12 });
  card(s, 0.62, 5.95, 3.9, 0.85, C.green);
  txt(s, [{ text: "USD 30,000", options: { bold: true, color: C.navy, fontSize: 18, breakLine: true } }, { text: "Approved: Google and LinkedIn, Sep to Dec", options: { color: C.muted, fontSize: 10 } }], { x: 0.85, y: 6.03, w: 3.6, h: 0.75 });
  card(s, 4.72, 5.95, 3.9, 0.85, C.gold);
  txt(s, [{ text: "About USD 48,700", options: { bold: true, color: C.navy, fontSize: 18, breakLine: true } }, { text: "Proposed: Jodel and outdoor, before VAT", options: { color: C.muted, fontSize: 10 } }], { x: 4.95, y: 6.03, w: 3.6, h: 0.75 });
  txt(s, "SAR converted at 3.75 to the USD. Outdoor and Jodel are vendor proposals, not yet approved.", { x: 8.82, y: 6.0, w: 3.9, h: 0.8, fontSize: 10, italic: true, color: C.muted, valign: "middle" });
}

// ---------- 4. Google Ads and LinkedIn ----------
{
  const s = content("Digital performance channels", "Google Ads and LinkedIn: USD 30,000, Saudi first",
    "Google Ads targets Saudi Arabia and Riyadh, Saudi residents only (Ignite). GA4, 1-28 Sep: 409 of 587 site sessions (70%) from google / cpc; English ads 212 sessions (97.2% engagement), Arabic 176 (86.9%).");
  table(s, 0.62, 1.75, [3.0, 2.5, 3.3, 3.3], [
    ["Channel and market", "Budget, Sep to Dec", "Target", "Status"],
    ["Google Search, Saudi", { t: "USD 9,000", b: true }, "Starter and Emerging, Riyadh first", tag("Live since 22 Sep", C.green)],
    ["LinkedIn, Saudi", { t: "USD 9,000", b: true }, "Affluent, lead generation forms", tag("Launching October", C.navy)],
    ["Google Search, rest of GCC", { t: "USD 6,000", b: true }, "Qatar, UAE, Kuwait, Bahrain, Oman", tag("After the Saudi test", C.navy)],
    ["LinkedIn, rest of GCC", { t: "USD 6,000", b: true }, "Qatar, UAE, Kuwait, Bahrain, Oman", tag("After the Saudi test", C.navy)],
    [{ t: "Total", b: true }, { t: "USD 30,000", b: true, c: C.navy }, "", ""],
  ], { rowH: 0.58, fontSize: 12.5 });
  txt(s, "FIRST RESULTS, SEPTEMBER", { x: 0.62, y: 5.25, w: 6, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  stat(s, 0.62, 5.6, 3.9, "70%", "of website visits", "came from Saudi Google Ads", C.navy);
  stat(s, 4.72, 5.6, 3.9, "97% / 87%", "engagement rate", "English ads / Arabic ads", C.navy);
  box(s, 8.82, 5.6, 3.9, 1.35, C.panel);
  txt(s, "A one-month test result for each channel will decide where the budget goes next.", { x: 9.0, y: 5.6, w: 3.6, h: 1.35, fontSize: 12, bold: true, color: C.navy, valign: "middle" });
}

// ---------- 5. Outdoor in Riyadh ----------
{
  const s = content("Outdoor and elevator media, Riyadh", "Outdoor options from Nadher Media: 91 screens in Saudi Arabia",
    "Nadher Media proposals 16 and 23 Sep 2026, Saudi screens only (the Al Liwan Bahrain weekend screens in the 16 Sep offer are left out). Elevator screens in 83 commercial towers (Al Nakhlah, Tamkeen, Faseelah Square, Laysen Valley and others in Riyadh and the Eastern Province) and 8 LED screens at The Zone, Riyadh. Plays per Nadher: 2,490 an hour (ad plays, not unique viewers). 2-week rate has no original price in the offer. USD at 3.75.");
  s.addImage({ path: path.join(IMG, "vb_elevator_gallery.png"), x: 0.62, y: 1.75, w: 5.1, h: 2.87 });
  s.addImage({ path: path.join(IMG, "vb_zone_mockup.png"), x: 5.9, y: 1.75, w: 2.45, h: 2.87 });
  txt(s, "Elevator screens, 83 towers  |  The Zone LED screens, 8 faces (eMYAA mock-up)", { x: 0.62, y: 4.66, w: 7.8, h: 0.28, fontSize: 9.5, italic: true, color: C.muted });
  table(s, 0.62, 5.0, [1.5, 2.2, 2.2, 1.9], [
    ["Duration", "Original rate", "eMYAA rate", "Ad plays"],
    ["2 weeks", "-", { t: "SAR 60,000", b: true }, "0.84M"],
    ["4 weeks", "SAR 120,000", { t: "SAR 102,000", b: true }, "1.67M"],
    ["8 weeks", "SAR 240,000", { t: "SAR 168,000", b: true, c: C.goldText }, "3.35M"],
    ["12 weeks", "SAR 360,000", { t: "SAR 252,000", b: true }, "5.02M"],
  ], { rowH: 0.38, fontSize: 11 });
  card(s, 8.6, 1.75, 4.12, 2.2, C.gold);
  txt(s, "LATEST OFFER, 23 SEP", { x: 8.85, y: 1.95, w: 3.7, h: 0.3, fontSize: 10, bold: true, color: C.goldText, charSpacing: 1 });
  txt(s, "SAR 168,000 + VAT", { x: 8.85, y: 2.3, w: 3.7, h: 0.55, fontSize: 24, bold: true, color: C.navy });
  txt(s, "8 weeks, 91 screens. About USD 44,800, 30% below the original rate.", { x: 8.85, y: 2.9, w: 3.7, h: 0.9, fontSize: 12, color: C.text });
  stat(s, 8.6, 4.15, 2.0, "83", "towers", "elevator screens", C.navy);
  stat(s, 10.72, 4.15, 2.0, "8", "LED faces", "The Zone, Riyadh", C.navy);
  box(s, 8.6, 5.7, 4.12, 0.95, C.panel);
  txt(s, "All prices + VAT. Proposal only, not yet approved.", { x: 8.75, y: 5.7, w: 3.85, h: 0.95, fontSize: 11, bold: true, color: C.navy, valign: "middle" });
}

// ---------- 6. Jodel ----------
{
  const s = content("Jodel, Saudi Arabia", "Jodel: 2.5 million impressions in 7 days for USD 3,862.50",
    "Nadher Media proposal for eMYAA, 16 Sep 2026: USD 3,862.50 (SAR 14,496.25) + VAT, 7 days, all KSA: 525,000 display and video impressions plus a 24-hour takeover poll of 2,000,000 impressions, one round of ad consulting, unlimited creatives. Cost per 1,000 impressions: about USD 1.53. Jodel KSA: 2.5M active users, 800,000 monthly unique users, average CTR 0.44%. BISB case study (Nadher Media, Oct 2025).");
  card(s, 0.62, 1.8, 5.95, 2.6, C.navy);
  txt(s, "THE OFFER", { x: 0.9, y: 2.0, w: 5, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  txt(s, "USD 3,862.50 + VAT", { x: 0.9, y: 2.35, w: 5.4, h: 0.6, fontSize: 28, bold: true, color: C.navy });
  txt(s, "SAR 14,496.25  |  7 days  |  all of Saudi Arabia", { x: 0.9, y: 3.0, w: 5.4, h: 0.3, fontSize: 12, color: C.muted });
  txt(s, "525,000 display and video impressions + 24-hour takeover poll (2,000,000 impressions)", { x: 0.9, y: 3.4, w: 5.4, h: 0.8, fontSize: 12, color: C.text });
  stat(s, 0.62, 4.6, 1.9, "2.5M", "impressions", "in 7 days", C.navy);
  stat(s, 2.65, 4.6, 1.9, "USD 1.53", "per 1,000", "impressions", C.goldText);
  stat(s, 4.68, 4.6, 1.9, "21 to 35", "audience age", "2.5M users in KSA", C.navy);
  card(s, 6.95, 1.8, 5.77, 4.15, C.gold);
  txt(s, "PEER RESULT: BAHRAIN ISLAMIC BANK IN KSA", { x: 7.25, y: 2.0, w: 5.3, h: 0.3, fontSize: 10, bold: true, color: C.goldText, charSpacing: 1 });
  txt(s, "7 days over Saudi National Day", { x: 7.25, y: 2.32, w: 5.2, h: 0.4, fontSize: 15, bold: true, color: C.navy });
  [["4.15M", "impressions"], ["0.43%", "click rate"], ["700+", "app downloads"], ["240", "accounts opened"]].forEach((k, i) => {
    const x = 7.25 + (i % 2) * 2.7, y = 2.95 + Math.floor(i / 2) * 1.35;
    txt(s, k[0], { x, y, w: 2.5, h: 0.6, fontSize: 28, bold: true, color: C.navy });
    txt(s, k[1], { x, y: y + 0.6, w: 2.5, h: 0.3, fontSize: 11, color: C.muted });
  });
  box(s, 0.62, 6.2, 12.1, 0.55, C.panel);
  txt(s, "Proposal only, not yet approved. Best timed with a salary week, as in the BISB campaign.", { x: 0.85, y: 6.2, w: 11.8, h: 0.55, fontSize: 11.5, bold: true, color: C.navy, valign: "middle" });
}

// ---------- 7. For discussion ----------
{
  const s = content("For discussion", "Where eMYAA and Vision Bank could work together",
    "Discussion topics only; nothing here is agreed.");
  const ideas = [
    ["Market insight", "What Saudi digital customers expect, and what has worked in your marketing."],
    ["Easy funding", "Simple transfers between Saudi bank accounts and eMYAA."],
    ["Joint campaigns", "Co-branded content and offers for first-time investors."],
  ];
  ideas.forEach((d, i) => {
    const x = 0.62 + i * 4.1;
    card(s, x, 1.85, 3.9, 3.0, [C.navy, C.gold, C.green][i]);
    txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: 2.1, w: 1, h: 0.6, fontSize: 28, bold: true, color: C.blue });
    txt(s, d[0], { x: x + 0.3, y: 2.8, w: 3.3, h: 0.45, fontSize: 19, bold: true, color: C.navy });
    txt(s, d[1], { x: x + 0.3, y: 3.35, w: 3.3, h: 1.3, fontSize: 13, color: C.text });
  });
  box(s, 0.62, 5.2, 12.1, 0.7, C.panel);
  box(s, 0.62, 5.2, 0.06, 0.7, C.gold);
  txt(s, "We would value your advice as we enter the Saudi market.", { x: 0.9, y: 5.2, w: 11.7, h: 0.7, fontSize: 13, italic: true, color: C.navy, valign: "middle" });
}

// ---------- 8. Thank you ----------
{
  const s = pres.addSlide();
  s.background = { path: path.join(IMG, "board_title_bg.png") };
  s.addImage({ path: path.join(IMG, "emyaa_white.png"), x: 0.75, y: 0.75, w: 1.75, h: 0.62 });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 10.67, y: 0.75, w: 1.92, h: 0.56 });
  txt(s, "Thank You", { x: 0.75, y: 2.75, w: 11.9, h: 0.95, fontSize: 42, bold: true, color: C.white, align: "center" });
  txt(s, "Invest Wisely. Trade Precisely.", { x: 0.75, y: 3.7, w: 11.9, h: 0.45, fontSize: 18, color: C.gold, align: "center" });
  txt(s, [
    { text: "Hatem Isa Hatem  |  Marketing Officer, WealthTech", options: { breakLine: true } },
    { text: "hhatem@ajyadcapital.com  |  +973 17 565 007", options: { breakLine: true } },
    { text: "Ajyad Capital, 37th Floor, Almoayyed Tower, Manama, Kingdom of Bahrain" },
  ], { x: 0.75, y: 4.75, w: 11.9, h: 1.0, fontSize: 12, color: C.white, align: "center" });
  txt(s, "Ajyad Capital is an Islamic investment firm (Category 1) licensed by the Central Bank of Bahrain.", { x: 0.75, y: 6.75, w: 11.9, h: 0.3, fontSize: 9.5, color: C.soft, align: "center" });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("saved", OUT));
