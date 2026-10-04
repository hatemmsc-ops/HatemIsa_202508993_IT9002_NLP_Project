// eMYAA x Vision Bank: introductory deck on eMYAA's Saudi market plans (Board visual language).
// External audience: no vendor prices (Nadher Media and Jodel quotes are confidential to eMYAA),
// no budgets and no client or deposit figures. Reach and audience figures are the vendors' own.
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
  txt(s, "Introductory meeting with Vision Bank  |  October 2026", { x: 0.75, y: 4.3, w: 11, h: 0.4, fontSize: 16, color: C.white });
  txt(s, "Confidential. Prepared by eMYAA Marketing, Ajyad Capital.", { x: 0.75, y: 6.75, w: 11, h: 0.3, fontSize: 10, color: C.soft });
}

// ---------- 2. About eMYAA ----------
{
  const s = content("Who we are", "A Shari'ah-compliant trading app built for GCC investors",
    "Product facts from the eMYAA Board pack (17 and 29 Sep 2026). Ajyad Capital is licensed by the Central Bank of Bahrain as an Islamic investment firm, Category 1.");
  txt(s, "eMYAA is the digital trading platform of Ajyad Capital, an Islamic investment firm (Category 1) licensed by the Central Bank of Bahrain. It went live in May 2026.",
    { x: 0.62, y: 1.75, w: 12.1, h: 0.6, fontSize: 13.5, color: C.text });
  const facts = [
    ["Shari'ah compliant", "Automated Shari'ah screening and a purification calculator"],
    ["Bilingual", "Full Arabic and English app, with AI support built in"],
    ["Easy to start", "No account fees and a minimum deposit of just USD 50"],
    ["Global and regional markets", "US, Saudi (Tadawul), Dubai (DFM) and Hong Kong in one app"],
  ];
  facts.forEach((f, i) => {
    const x = 0.62 + (i % 2) * 6.1, y = 2.6 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 5.95, 1.35, i % 2 ? C.gold : C.navy);
    txt(s, f[0], { x: x + 0.3, y: y + 0.25, w: 5.4, h: 0.4, fontSize: 16, bold: true, color: C.navy });
    txt(s, f[1], { x: x + 0.3, y: y + 0.72, w: 5.4, h: 0.5, fontSize: 12, color: C.muted });
  });
  box(s, 0.62, 5.85, 12.1, 0.75, C.panel);
  box(s, 0.62, 5.85, 0.06, 0.75, C.gold);
  txt(s, [{ text: "Products  ", options: { bold: true, color: C.navy } }, { text: "Shari'ah-compliant stocks, Sukuk and ETFs, with a full app experience from onboarding to trading.", options: { color: C.text } }],
    { x: 0.9, y: 5.85, w: 11.7, h: 0.75, fontSize: 12.5, valign: "middle" });
}

// ---------- 3. Why Saudi ----------
{
  const s = content("Why Saudi Arabia", "Saudi Arabia is our first priority market in the GCC",
    "Early signal: Google Ads targeted at Saudi residents went live on 22 Sep 2026 and brought 70% of eMYAA website sessions in September (409 of 587, GA4, 1-28 Sep). Budgets are not shown in this external deck.");
  const cols = [
    ["The audience", "A young, mobile-first population that is open to investing through apps.", C.navy],
    ["The demand", "Strong interest in Shari'ah-compliant products, which is where eMYAA starts.", C.gold],
    ["The fit", "Arabic-first app, Tadawul access and a low USD 50 entry point.", C.green],
  ];
  cols.forEach((c, i) => {
    const x = 0.62 + i * 4.1;
    card(s, x, 1.8, 3.9, 2.0, c[2]);
    txt(s, c[0], { x: x + 0.3, y: 2.05, w: 3.3, h: 0.4, fontSize: 16, bold: true, color: C.navy });
    txt(s, c[1], { x: x + 0.3, y: 2.55, w: 3.3, h: 1.1, fontSize: 12, color: C.text });
  });
  txt(s, "WHAT WE HAVE SEEN SO FAR", { x: 0.62, y: 4.15, w: 8, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  stat(s, 0.62, 4.55, 3.9, "70%", "of September website visits", "came from Saudi-targeted Google Ads", C.navy);
  stat(s, 4.72, 4.55, 3.9, "22 Sep", "Saudi paid media went live", "Google Search, Arabic and English", C.goldText);
  stat(s, 8.82, 4.55, 3.9, "Riyadh", "core city in our targeting", "ads shown to Saudi residents only", C.green);
}

// ---------- 4. Who we target ----------
{
  const s = content("Segments", "Four investor segments, each with its own channel and message",
    "Segments and hooks from the eMYAA Board pack (29 Sep 2026). Customer acquisition cost targets are internal and not shown.");
  const segs = [
    ["Starter", "Under USD 1,000", "21 to 35, salaried, often a first investment product", "Social media, sponsored ads and creators", "Simple bilingual app, no account fees, USD 50 to start"],
    ["Emerging", "USD 1,000 to 5,000", "Young professionals building a regular saving habit", "Google Search and app install ads", "Low-cost access to US and Saudi markets in one place"],
    ["Affluent", "USD 5,000 to 50,000", "Professionals and business owners", "LinkedIn lead generation", "Licensed Islamic firm, wide product range, personal service"],
    ["High net worth", "Over USD 50,000", "Senior decision makers", "Ajyad placement team, in person", "Relationship-led, with the app as the daily tool"],
  ];
  const colX = [0.62, 2.55, 4.75, 7.45, 10.0], colW = [1.9, 2.15, 2.65, 2.5, 2.72];
  const heads = ["Segment", "Deposit size", "Who they are", "How we reach them", "Why eMYAA"];
  heads.forEach((h, j) => { box(s, colX[j], 1.75, colW[j], 0.45, C.navy); txt(s, h, { x: colX[j] + 0.12, y: 1.75, w: colW[j] - 0.2, h: 0.45, fontSize: 11, bold: true, color: C.white, valign: "middle" }); });
  segs.forEach((r, i) => {
    const y = 2.2 + i * 1.08;
    r.forEach((c, j) => {
      box(s, colX[j], y, colW[j], 1.08, i % 2 ? C.white : C.panel, C.line);
      txt(s, c, { x: colX[j] + 0.12, y, w: colW[j] - 0.22, h: 1.08, fontSize: j === 0 ? 13 : 11, bold: j === 0, color: j === 0 ? C.navy : C.text, valign: "middle" });
    });
  });
  txt(s, "WealthTech reaches the first three segments at scale through digital channels. The Ajyad placement team closes high net worth relationships one at a time.",
    { x: 0.62, y: 6.6, w: 12.1, h: 0.35, fontSize: 10.5, italic: true, color: C.muted });
}

// ---------- 5. Rollout ----------
{
  const s = content("Our approach", "Launch in Saudi Arabia, learn fast, then expand across the GCC",
    "Rollout phases from the paid media plan with Ignite (Sep 2026) and the Board runway (Sep to Dec 2026).");
  const phases = [
    ["Phase 1", "Launch in Saudi Arabia", "Sep to Oct 2026", ["Google Search and app ads live", "LinkedIn lead generation for Affluent investors", "First Saudi creator video", "Arabic and English ads tested side by side"], C.navy],
    ["Phase 2", "Measure and refine", "Oct to Nov 2026", ["One-month test results by channel", "Keep what converts, cut what does not", "Add brand awareness: outdoor and Jodel bursts", "Tighten onboarding for Saudi users"], C.goldText],
    ["Phase 3", "Scale and expand", "Nov to Dec 2026", ["Scale the best Saudi channels", "Extend to Qatar, UAE, Kuwait, Bahrain and Oman", "Monthly creator content across the GCC", "Partnerships with regional institutions"], C.green],
  ];
  phases.forEach((p, i) => {
    const x = 0.62 + i * 4.1;
    card(s, x, 1.8, 3.9, 4.75, p[4]);
    txt(s, p[0].toUpperCase(), { x: x + 0.3, y: 2.0, w: 3.3, h: 0.3, fontSize: 10, bold: true, color: p[4], charSpacing: 1 });
    txt(s, p[1], { x: x + 0.3, y: 2.32, w: 3.3, h: 0.45, fontSize: 17, bold: true, color: C.navy });
    txt(s, p[2], { x: x + 0.3, y: 2.8, w: 3.3, h: 0.3, fontSize: 11, color: C.muted });
    box(s, x + 0.3, 3.2, 3.3, 0.02, C.line);
    bullets(s, p[3], { x: x + 0.3, y: 3.35, w: 3.35, h: 3.0, fontSize: 13.5 });
  });
}

// ---------- 6. Google Ads and LinkedIn ----------
{
  const s = content("Digital performance channels", "Google Ads and LinkedIn: reaching investors when they are looking",
    "Google Ads live since 22 Sep 2026, targeting Saudi Arabia and Riyadh, restricted to Saudi residents (Ignite). LinkedIn lead generation targets the Affluent segment; launch being confirmed with Ignite in October. GA4: 409 of 587 September sessions from google / cpc; English ads 212 sessions (97.2% engagement), Arabic 176 (86.9%), 1-28 Sep.");
  const ch = [
    ["Google Ads", "Live since 22 Sep 2026", C.green, [
      [B("Who: "), { text: "Starter and Emerging investors searching for ways to invest" }],
      [B("Where: "), { text: "Saudi Arabia, with Riyadh as the core, Saudi residents only" }],
      [B("How: "), { text: "Search ads in Arabic and English, plus app install and video ads" }],
      [B("So far: "), { text: "70% of September website visits, and both languages engaging well" }],
    ]],
    ["LinkedIn", "Launching in October 2026", C.goldText, [
      [B("Who: "), { text: "Affluent professionals, business owners and senior managers" }],
      [B("Where: "), { text: "Saudi Arabia first, then the wider GCC" }],
      [B("How: "), { text: "Lead generation forms, so interested investors can ask to be contacted" }],
      [B("Next: "), { text: "Our sales team follows up every lead personally" }],
    ]],
  ];
  ch.forEach((c, i) => {
    const x = 0.62 + i * 6.1;
    card(s, x, 1.8, 5.95, 4.0, c[2]);
    txt(s, c[0], { x: x + 0.3, y: 2.02, w: 3.5, h: 0.45, fontSize: 20, bold: true, color: C.navy });
    box(s, x + 3.55, 2.08, 2.1, 0.36, C.panel);
    txt(s, c[1], { x: x + 3.55, y: 2.08, w: 2.1, h: 0.36, fontSize: 9.5, bold: true, color: c[2], align: "center", valign: "middle" });
    bullets(s, c[3], { x: x + 0.3, y: 2.75, w: 5.4, h: 2.9, fontSize: 14.5 });
  });
  box(s, 0.62, 6.0, 12.1, 0.7, C.panel);
  box(s, 0.62, 6.0, 0.06, 0.7, C.gold);
  txt(s, [{ text: "One-month test  ", options: { bold: true, color: C.navy } }, { text: "Results for both channels (traffic, installs, onboarding and cost per client) will guide the November scale-up.", options: { color: C.text } }],
    { x: 0.9, y: 6.0, w: 11.7, h: 0.7, fontSize: 12, valign: "middle" });
}

// ---------- 7. Influencers ----------
{
  const s = content("Creators and community", "Trusted Saudi and GCC voices to explain investing simply",
    "Creator programme: one creator video a month (24F and eMYAA content calendars). Musheera 5.3M and Ali Sabeel 1.3M views on Instagram (profile screenshot). Abdulelah Al Harbi masterclass posted on Instagram in October and sponsored through Ignite. Manal Talal scheduled for 22 Oct (October calendar V2). Mention & Win: monthly $500 contest across the GCC.");
  s.addImage({ path: path.join(IMG, "tile9_abdulelah.png"), x: 0.62, y: 1.8, w: 3.3, h: 4.26 });
  txt(s, [{ text: "Abdulelah Al Harbi", options: { bold: true, color: C.navy, breakLine: true } }, { text: "Saudi creator, investing masterclass, October 2026", options: { color: C.muted } }],
    { x: 0.62, y: 6.12, w: 3.3, h: 0.6, fontSize: 10.5, align: "center" });
  txt(s, "HOW WE WORK WITH CREATORS", { x: 4.3, y: 1.8, w: 8, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  bullets(s, [
    [B("One creator a month: "), { text: "short, practical videos on how to start investing the Shari'ah-compliant way" }],
    [B("Saudi voices first: "), { text: "Saudi creators lead the launch, with GCC creators to follow" }],
    [B("Paid boost: "), { text: "each video is sponsored to reach the right age groups and cities" }],
    [B("Community: "), { text: "a monthly Mention & Win contest with a USD 500 prize across the GCC" }],
  ], { x: 4.3, y: 2.2, w: 8.4, h: 2.3, fontSize: 13 });
  txt(s, "CREATOR VIDEOS SO FAR (INSTAGRAM VIEWS)", { x: 4.3, y: 4.55, w: 8, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  stat(s, 4.3, 4.95, 2.7, "5.3M", "Musheera", "July 2026", C.navy);
  stat(s, 7.15, 4.95, 2.7, "1.3M", "Ali Sabeel", "August 2026", C.navy);
  stat(s, 10.0, 4.95, 2.72, "22 Oct", "Manal Talal", "next creator video", C.goldText);
}

// ---------- 8. Outdoor in Riyadh ----------
{
  const s = content("Outdoor and indoor media", "Being seen where Saudi professionals work and spend time",
    "Options proposed by Nadher Media (Sep 2026): elevator screens in 83 commercial towers in Riyadh and the Eastern Province (e.g. Al Nakhlah Tower, Tamkeen Tower, Faseelah Square, Laysen Valley) and 8 prime LED screens at The Zone, Riyadh. Frequency per Nadher: 2,490 plays an hour, about 1.67 million plays in four weeks (ad plays, not unique viewers). Rates are confidential to eMYAA and not shown. Status: under evaluation.");
  s.addImage({ path: path.join(IMG, "vb_elevator_gallery.png"), x: 0.62, y: 1.75, w: 7.6, h: 4.28 });
  txt(s, "Elevator screens in business towers across Riyadh and the Eastern Province", { x: 0.62, y: 6.08, w: 7.6, h: 0.3, fontSize: 10, italic: true, color: C.muted });
  s.addImage({ path: path.join(IMG, "vb_zone_mockup.png"), x: 8.5, y: 1.75, w: 2.05, h: 2.4 });
  txt(s, [{ text: "The Zone, Riyadh", options: { bold: true, color: C.navy, breakLine: true } }, { text: "8 prime LED screens at a high-end lifestyle destination (eMYAA mock-up)", options: { color: C.muted } }],
    { x: 10.7, y: 1.8, w: 2.05, h: 2.3, fontSize: 10.5 });
  stat(s, 8.5, 4.3, 2.05, "83", "towers", "elevator screens", C.navy);
  stat(s, 10.67, 4.3, 2.05, "1.67M", "ad plays", "in four weeks", C.goldText);
  box(s, 8.5, 5.8, 4.22, 0.6, C.panel);
  txt(s, "Status: under evaluation for Phase 2", { x: 8.6, y: 5.8, w: 4.0, h: 0.6, fontSize: 11, bold: true, color: C.navy, align: "center", valign: "middle" });
}

// ---------- 9. Jodel ----------
{
  const s = content("Reaching young Saudis", "Jodel: a short, high-impact burst to build awareness fast",
    "Jodel facts from Nadher Media's proposal (Sep 2026): 2.5M active users in KSA, 800,000 monthly unique users, 192M monthly ad impressions, average CTR 0.44%, audience students and young professionals aged 21 to 35. Proposed 7-day eMYAA campaign: 525,000 display and video impressions plus a 24-hour takeover poll of about 2M impressions. Peer case: Bahrain Islamic Bank (BISB), 7 days over Saudi National Day: 4.15M impressions, 8,973 clicks (0.43% CTR), poll 2.68M impressions and 6,739 votes, 700+ downloads and 240 accounts. Price confidential, not shown.");
  txt(s, "JODEL IN SAUDI ARABIA", { x: 0.62, y: 1.75, w: 6, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  stat(s, 0.62, 2.12, 2.9, "2.5M", "active users in KSA", "students and young professionals", C.navy);
  stat(s, 3.65, 2.12, 2.9, "21 to 35", "core age group", "matches our Starter segment", C.navy);
  txt(s, "A PROPOSED 7-DAY LAUNCH BURST", { x: 0.62, y: 3.75, w: 6, h: 0.3, fontSize: 10, bold: true, color: C.blue, charSpacing: 1 });
  bullets(s, [
    [B("Display and video ads: "), { text: "full-screen ads across the Kingdom" }],
    [B("24-hour takeover poll: "), { text: "an interactive question every user sees for a day" }],
    [B("About 2.5 million impressions "), { text: "in one week" }],
  ], { x: 0.62, y: 4.1, w: 5.95, h: 1.7, fontSize: 12.5 });
  card(s, 6.95, 1.8, 5.77, 4.6, C.gold);
  txt(s, "PEER EXAMPLE", { x: 7.25, y: 2.0, w: 5, h: 0.3, fontSize: 10, bold: true, color: C.goldText, charSpacing: 1 });
  txt(s, "Bahrain Islamic Bank in Saudi Arabia, 7 days over National Day", { x: 7.25, y: 2.3, w: 5.2, h: 0.7, fontSize: 15, bold: true, color: C.navy });
  [["4.15M", "impressions"], ["0.43%", "click rate"], ["700+", "app downloads"], ["240", "accounts opened"]].forEach((k, i) => {
    const x = 7.25 + (i % 2) * 2.7, y = 3.15 + Math.floor(i / 2) * 1.35;
    txt(s, k[0], { x, y, w: 2.5, h: 0.6, fontSize: 28, bold: true, color: C.navy });
    txt(s, k[1], { x, y: y + 0.6, w: 2.5, h: 0.3, fontSize: 11, color: C.muted });
  });
  txt(s, "Source: Nadher Media case study, October 2025", { x: 7.25, y: 5.95, w: 5.2, h: 0.3, fontSize: 9, italic: true, color: C.muted });
}

// ---------- 10. Timeline ----------
{
  const s = content("The plan ahead", "What runs when, October to December 2026",
    "Live: Google Ads (since 22 Sep). Planned: LinkedIn (October), creator videos (monthly). Proposed and under evaluation: Jodel burst and outdoor media (Phase 2). GCC expansion after the one-month test.");
  const months = ["October", "November", "December"];
  const x0 = 3.6, mw = 3.04;
  months.forEach((m, i) => { box(s, x0 + i * mw, 1.75, mw - 0.06, 0.42, C.navy); txt(s, m, { x: x0 + i * mw, y: 1.75, w: mw - 0.06, h: 0.42, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" }); });
  const rows = [
    ["Google Ads, Saudi", "Live", 0, 3, C.green],
    ["LinkedIn, Affluent", "Planned", 0.3, 3, C.navy],
    ["Saudi creator videos", "Planned", 0, 3, C.navy],
    ["Jodel launch burst", "Proposed", 1.0, 1.25, C.gold],
    ["Outdoor media, Riyadh", "Proposed", 1.0, 2.0, C.gold],
    ["One-month test review", "Planned", 1.0, 1.35, C.blue],
    ["Expand to the wider GCC", "Planned", 1.6, 3, C.green],
  ];
  rows.forEach((r, i) => {
    const y = 2.35 + i * 0.6;
    box(s, 0.62, y, 12.1, 0.52, i % 2 ? C.white : C.panel);
    txt(s, r[0], { x: 0.75, y, w: 2.2, h: 0.52, fontSize: 11, bold: true, color: C.navy, valign: "middle" });
    txt(s, r[1], { x: 2.75, y, w: 0.8, h: 0.52, fontSize: 9, color: C.muted, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x0 + r[2] * mw, y: y + 0.13, w: (r[3] - r[2]) * mw - 0.08, h: 0.26, fill: { color: r[4] }, line: { color: r[4], width: 0 }, rectRadius: 0.13 });
  });
  [["Live", C.green], ["Planned", C.navy], ["Proposed", C.gold]].forEach((l, i) => {
    box(s, 0.62 + i * 1.6, 6.65, 0.18, 0.18, l[1]);
    txt(s, l[0], { x: 0.88 + i * 1.6, y: 6.6, w: 1.2, h: 0.28, fontSize: 10, color: C.text, valign: "middle" });
  });
}

// ---------- 11. Working together ----------
{
  const s = content("For discussion", "Where eMYAA and Vision Bank could work together",
    "Discussion topics only; nothing here is agreed. Purpose of the meeting: learn from Vision Bank's experience of the Saudi market and explore areas of collaboration.");
  const ideas = [
    ["Market insight", "Your view on what Saudi digital customers expect from an investing app, and what has worked in your own marketing."],
    ["Easy funding", "Simple ways for Saudi customers to move money between their bank account and their eMYAA account."],
    ["Joint campaigns", "Co-branded content and offers that help customers start saving and investing the Shari'ah-compliant way."],
    ["Events and education", "Joint sessions on investing basics for young professionals, in Arabic and English."],
  ];
  ideas.forEach((d, i) => {
    const x = 0.62 + (i % 2) * 6.1, y = 1.8 + Math.floor(i / 2) * 2.15;
    card(s, x, y, 5.95, 1.95, i % 2 ? C.gold : C.navy);
    txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: y + 0.25, w: 0.8, h: 0.6, fontSize: 26, bold: true, color: i % 2 ? C.goldText : C.blue });
    txt(s, d[0], { x: x + 1.15, y: y + 0.3, w: 4.5, h: 0.45, fontSize: 17, bold: true, color: C.navy });
    txt(s, d[1], { x: x + 1.15, y: y + 0.82, w: 4.55, h: 1.0, fontSize: 12, color: C.text });
  });
  box(s, 0.62, 6.15, 12.1, 0.55, C.panel);
  box(s, 0.62, 6.15, 0.06, 0.55, C.gold);
  txt(s, "We would value your advice as we enter the Saudi market, and we are open to ideas on your side.", { x: 0.9, y: 6.15, w: 11.7, h: 0.55, fontSize: 12, italic: true, color: C.navy, valign: "middle" });
}

// ---------- 12. Thank you ----------
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
