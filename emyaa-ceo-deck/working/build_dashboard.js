// Daily Performance Dashboard rebuilt as an editable deck in the Board visual language.
// Data: working/audit-data/dashboard-pages.json (from working/dashboard_extract.py), newest report first.
// Run from emyaa-ceo-deck/: NODE_PATH=<dir with pptxgenjs> node working/build_dashboard.js
//                           then python3 working/embed_fonts.py output/dashboard-pptx/<file>.pptx
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const IMG = path.join(__dirname, "charts");
const OUT = path.join(__dirname, "..", "output", "dashboard-pptx", "eMYAA_-_Daily_Performance_Dashboard.pptx");
const pages = JSON.parse(fs.readFileSync(path.join(__dirname, "audit-data", "dashboard-pages.json"), "utf8"));

const C = {
  navy: "132257", navyMid: "22376F", green: "1BA97C", gold: "D4AF37", goldText: "9A7B1F", panel: "E3E8F2",
  line: "D9DFEA", soft: "C6CFE4", text: "1F2A44", muted: "7A849C", foot: "8A93A8", red: "C53030", white: "FFFFFF", bg: "F7F9FC",
};
const F = "Exo 2";
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "eMYAA";
pres.title = "eMYAA Daily Performance Dashboard";

const num = (s) => parseFloat(String(s).replace(/[$,\s]/g, ""));
const money = (v) => "$" + v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const int = (v) => Math.round(v).toLocaleString("en-US");
const shortDate = (d) => d.replace(/^0/, "").replace(/ 2026$/, "");
function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, fontFace: F, margin: 0, valign: "top", color: C.text }, o)); }

function chrome(s, { kicker, title, sub, n, source }) {
  s.background = { color: C.bg };
  s.addImage({ path: path.join(IMG, "board_header.png"), x: 0, y: 0, w: 13.333, h: 1.3 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.3, w: 13.333, h: 0.027, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
  txt(s, kicker, { x: 0.62, y: 0.28, w: 9, h: 0.2, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
  txt(s, title, { x: 0.62, y: 0.5, w: 10, h: 0.45, fontSize: 22, bold: true, color: C.white, valign: "middle" });
  if (sub) txt(s, sub, { x: 0.62, y: 0.95, w: 10.3, h: 0.25, fontSize: 10, color: C.soft });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 11.27, y: 0.44, w: 1.45, h: 0.42 });
  txt(s, "Source: " + source, { x: 0.62, y: 6.86, w: 10.3, h: 0.25, fontSize: 8, color: C.muted });
  txt(s, "eMYAA Trading Platform  |  Daily Performance Dashboard", { x: 0.62, y: 7.12, w: 6, h: 0.22, fontSize: 8, color: C.foot });
  txt(s, String(n), { x: 10.72, y: 7.12, w: 0.5, h: 0.22, fontSize: 8, color: C.foot, align: "right" });
  s.addImage({ path: path.join(IMG, "emyaa_navy.png"), x: 11.6, y: 6.93, w: 1.12, h: 0.4 });
}

// ---------- data helpers ----------
const moneyKey = (c) => (c["Total AUM"] ? "Total AUM" : "Total Deposits");
function metrics(p) {
  const c = p.cards, v = (k) => num(c[k].value);
  return {
    android: v("Android Downloads"), ios: v("iOS Downloads"), aRate: c["Android App Rating"].value, iRate: c["iOS App Rating"].value,
    aRateSub: c["Android App Rating"].sub, iRateSub: c["iOS App Rating"].sub,
    deposited: v("Users Deposited"), money: v(moneyKey(c)), moneyLabel: moneyKey(c), traded: v("Users Traded"), trades: v("Total No. of Trades"),
    onboarded: v("Users Onboarded"), rejected: v("Rejected KYC Requests"), incomplete: v("Incomplete KYC Users"),
    created: v("Support Tickets Created"), resolved: v("Support Tickets Resolved"), email: v("Email Queries"), ai: v("AI Client Support Queries"),
  };
}

// ---------- 1. Cover ----------
{
  const s = pres.addSlide();
  s.background = { path: path.join(IMG, "board_title_bg.png") };
  s.addImage({ path: path.join(IMG, "emyaa_white.png"), x: 0.75, y: 0.75, w: 1.75, h: 0.62 });
  s.addImage({ path: path.join(IMG, "ajyad_white.png"), x: 10.67, y: 0.75, w: 1.92, h: 0.56 });
  txt(s, "Daily Performance Dashboard", { x: 0.75, y: 2.8, w: 11.9, h: 0.95, fontSize: 40, bold: true, color: C.white });
  txt(s, "PLATFORM ANALYTICS", { x: 0.75, y: 3.8, w: 11, h: 0.45, fontSize: 20, bold: true, color: C.gold, charSpacing: 1 });
  txt(s, `${pages.length} daily reports, ${pages[pages.length - 1].date} to ${pages[0].date}`, { x: 0.75, y: 4.3, w: 10, h: 0.4, fontSize: 16, color: C.white });
  txt(s, "Source: eMYAA Daily Performance Dashboard PDFs. Figures are cumulative, all time, unless a report states a period.", { x: 0.75, y: 6.7, w: 11.9, h: 0.3, fontSize: 10, color: C.soft });
}

// ---------- 2. Trend overview ----------
{
  const chron = pages.slice().reverse();
  const m = chron.map(metrics);
  const latest = metrics(pages[0]);
  const dl = latest.android + latest.ios;
  const s = pres.addSlide();
  chrome(s, {
    kicker: "OVERVIEW", n: 2,
    title: `Since launch: ${int(dl)} downloads, ${int(latest.onboarded)} onboarded, ${int(latest.deposited)} funded accounts`,
    sub: `Cumulative totals on each report date, ${chron[0].date} to ${pages[0].date}`,
    source: `eMYAA Daily Performance Dashboard, ${pages.length} reports. Downloads = Android + iOS.`,
  });
  // shape-drawn line chart, x spaced by calendar day
  const day = (d) => Date.parse(d.replace(/^(\d) /, "0$1 ")) / 86400000;
  const d0 = day(chron[0].date), d1 = day(chron[chron.length - 1].date);
  const x = 1.25, y = 1.75, w = 7.2, h = 4.25, top = y + 0.35, bottom = y + h - 0.45, max = 500;
  const X = (d) => x + ((day(d) - d0) / (d1 - d0)) * w, Y = (v) => bottom - (v / max) * (bottom - top);
  txt(s, "Cumulative downloads, onboarded users and funded accounts", { x, y, w, h: 0.28, fontSize: 11, color: C.navy });
  for (let v = 0; v <= max; v += 100) {
    s.addShape(pres.shapes.LINE, { x, y: Y(v), w, h: 0, line: { color: v === 0 ? C.soft : "E6EAF2", width: v === 0 ? 1 : 0.75 } });
    txt(s, String(v), { x: x - 0.55, y: Y(v) - 0.12, w: 0.45, h: 0.24, fontSize: 9, color: C.muted, align: "right", valign: "middle" });
  }
  ["Jun", "Jul", "Aug", "Sep"].forEach((mo, i) => {
    const d = `01 ${mo} 2026`;
    s.addShape(pres.shapes.LINE, { x: X(d), y: bottom, w: 0, h: 0.06, line: { color: C.soft, width: 1 } });
    txt(s, `1 ${mo}`, { x: X(d) - 0.4, y: bottom + 0.08, w: 0.8, h: 0.22, fontSize: 9, color: C.muted, align: "center" });
  });
  const series = [
    ["Downloads", C.navy, m.map((r) => r.android + r.ios)],
    ["Onboarded", C.green, m.map((r) => r.onboarded)],
    ["Funded accounts", C.gold, m.map((r) => r.deposited)],
  ];
  series.forEach(([name, col, vals]) => {
    for (let i = 0; i < vals.length - 1; i++) {
      const x1 = X(chron[i].date), x2 = X(chron[i + 1].date), y1 = Y(vals[i]), y2 = Y(vals[i + 1]);
      if (x2 - x1 < 0.001 && Math.abs(y2 - y1) < 0.001) continue;
      s.addShape(pres.shapes.LINE, { x: x1, y: Math.min(y1, y2), w: Math.max(x2 - x1, 0.001), h: Math.abs(y2 - y1), flipV: y2 < y1, line: { color: col, width: 2.25 } });
    }
    const last = vals[vals.length - 1];
    s.addShape(pres.shapes.OVAL, { x: X(pages[0].date) - 0.05, y: Y(last) - 0.05, w: 0.1, h: 0.1, fill: { color: col }, line: { color: col, width: 0 } });
    txt(s, [{ text: int(last) + "  ", options: { bold: true, color: col } }, { text: name, options: { color: C.muted } }], { x: X(pages[0].date) + 0.1, y: Y(last) - 0.13, w: 2.2, h: 0.26, fontSize: 10, valign: "middle" });
  });
  // KPI cards
  const kp = [
    [int(dl), "app downloads", `Android ${int(latest.android)}  |  iOS ${int(latest.ios)}`, C.navy],
    [int(latest.onboarded), "users onboarded", `${(latest.onboarded / dl * 100).toFixed(1)}% of downloads`, C.green],
    [int(latest.deposited), "funded accounts", `${(latest.deposited / latest.onboarded * 100).toFixed(1)}% of onboarded`, C.gold],
    [int(latest.incomplete), "started KYC, not completed", `plus ${int(latest.rejected)} rejected, resubmission required`, C.red],
  ];
  kp.forEach((k, i) => {
    const cx = 10.0, cy = 1.7 + i * 1.12;
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 2.72, h: 1.0, fill: { color: C.white }, line: { color: C.line, width: 0.75 } });
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 2.72, h: 0.035, fill: { color: k[3] }, line: { color: k[3], width: 0 } });
    txt(s, k[0], { x: cx + 0.18, y: cy + 0.12, w: 1.1, h: 0.45, fontSize: 22, bold: true, color: k[3] === C.gold ? C.goldText : k[3] });
    txt(s, k[1], { x: cx + 1.2, y: cy + 0.16, w: 1.45, h: 0.42, fontSize: 9.5, bold: true, color: C.navy, valign: "middle" });
    txt(s, k[2], { x: cx + 0.18, y: cy + 0.64, w: 2.45, h: 0.25, fontSize: 8.5, color: C.muted });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.62, y: 6.2, w: 12.1, h: 0.5, fill: { color: C.panel }, line: { color: C.panel, width: 0 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.62, y: 6.2, w: 0.05, h: 0.5, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
  txt(s, [{ text: "Takeaway   ", options: { bold: true, color: C.navy } }, { text: `Downloads keep growing, but about 1 in ${Math.round(dl / latest.onboarded)} becomes an onboarded user and 1 in ${Math.round(latest.onboarded / latest.deposited)} of those funds an account.`, options: { color: C.text } }], { x: 0.85, y: 6.2, w: 11.75, h: 0.5, fontSize: 11.5, valign: "middle" });
}

// ---------- 3+. One slide per report ----------
const COLS = [
  { name: "DOWNLOADS & RATINGS", color: C.navy },
  { name: "USER ACTIVITY", color: C.green },
  { name: "KYC STATUS", color: C.gold },
  { name: "SUPPORT TICKETS", color: C.navyMid },
];
pages.forEach((p, idx) => {
  const n = idx + 3, cur = metrics(p), prevPage = pages[idx + 1], prev = prevPage ? metrics(prevPage) : null;
  const s = pres.addSlide();
  const since = prevPage ? shortDate(prevPage.date) : null;
  const isAum = cur.moneyLabel === "Total AUM";
  chrome(s, {
    kicker: "PLATFORM ANALYTICS  |  PERFORMANCE DASHBOARD", n,
    title: `Report date: ${p.date}`,
    sub: `Reporting period: ${p.period}` + (prevPage ? `   |   Change on each card is against the previous report, ${prevPage.date}` : "   |   First report in the series"),
    source: `eMYAA Daily Performance Dashboard, ${p.source}, page ${p.page}.` + (isAum ? "" : " This report labels the money card 'Total Deposits'; later reports label the same measure 'Total AUM'."),
  });
  // delta: good = true when an increase is good, false when an increase is bad, null = neutral
  const delta = (k, good, fmt) => {
    if (!prev) return { text: "First report", color: C.muted };
    const d = cur[k] - prev[k];
    if (Math.abs(d) < 0.005) return { text: "No change", color: C.muted };
    const up = d > 0, better = good === null ? null : (up === good);
    const f = fmt ? fmt(Math.abs(d)) : int(Math.abs(d));
    return { text: `${up ? "+" : "-"}${f}`, color: better === null ? C.muted : better ? C.green : C.red };
  };
  const pct = (a, b) => (b ? `${Math.round((a / b) * 100)}%` : "n/a");
  const cols = [
    [
      [int(cur.android), "Android downloads", "Google Play Store", delta("android", true)],
      [int(cur.ios), "iOS downloads", "Apple App Store", delta("ios", true)],
      [cur.aRate + " / 5", "Android app rating", cur.aRateSub.replace(/\s*·\s*/, ", "), null],
      [cur.iRate + " / 5", "iOS app rating", cur.iRateSub.replace(/\s*·\s*/, ", "), null],
    ],
    [
      [int(cur.deposited), "Users deposited", "Funded accounts", delta("deposited", true)],
      [money(cur.money), isAum ? "Total AUM" : "Total deposits", "Cumulative platform deposits", delta("money", true, money)],
      [int(cur.traded), "Users traded", "Placed at least 1 trade", delta("traded", true)],
      [int(cur.trades), "Total trades", "Executed across all users", delta("trades", true)],
    ],
    [
      [int(cur.onboarded), "Users onboarded", "Completed signup and verification", delta("onboarded", true)],
      [int(cur.incomplete), "Incomplete KYC", "Started KYC but did not proceed", delta("incomplete", false)],
      [int(cur.rejected), "Rejected KYC", "Resubmission required", delta("rejected", false)],
      [int(cur.onboarded + cur.incomplete + cur.rejected), "Total KYC applications", "Onboarded + incomplete + rejected", null],
    ],
    [
      [int(cur.created), "Tickets created", "Submitted by users", delta("created", null)],
      [int(cur.resolved), "Tickets resolved", `${pct(cur.resolved, cur.created)} of tickets created`, delta("resolved", null)],
      [int(cur.email), "Email queries", "Submitted by users", delta("email", null)],
      [int(cur.ai), "AI support queries", "Submitted by users", delta("ai", null)],
    ],
  ];
  const cw = 2.875, gap = 0.2, top = 1.5, ch = 1.1, vg = 0.1;
  cols.forEach((cards, i) => {
    const cx = 0.62 + i * (cw + gap);
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: top + 0.04, w: 0.05, h: 0.17, fill: { color: COLS[i].color }, line: { color: COLS[i].color, width: 0 } });
    txt(s, COLS[i].name, { x: cx + 0.12, y: top, w: cw - 0.12, h: 0.25, fontSize: 9, bold: true, color: C.muted, charSpacing: 2, valign: "middle" });
    cards.forEach((cd, j) => {
      const y = top + 0.36 + j * (ch + vg);
      s.addShape(pres.shapes.RECTANGLE, { x: cx, y, w: cw, h: ch, fill: { color: C.white }, line: { color: C.line, width: 0.75 } });
      s.addShape(pres.shapes.RECTANGLE, { x: cx, y, w: cw, h: 0.035, fill: { color: COLS[i].color }, line: { color: COLS[i].color, width: 0 } });
      txt(s, cd[0], { x: cx + 0.18, y: y + 0.12, w: cw - 0.36, h: 0.46, fontSize: cd[0].length > 8 ? 19 : 24, bold: true, color: C.navy, valign: "middle" });
      txt(s, cd[1], { x: cx + 0.18, y: y + 0.56, w: cw - 0.36, h: 0.24, fontSize: 10.5, bold: true, color: C.navy });
      txt(s, cd[2], { x: cx + 0.18, y: y + 0.8, w: cw - 0.36, h: 0.22, fontSize: 8.5, color: C.muted });
      if (cd[3]) txt(s, cd[3].text, { x: cx + 1.75, y: y + 0.2, w: cw - 1.9, h: 0.3, fontSize: 8.5, bold: true, color: cd[3].color, align: "right", valign: "middle" });
    });
  });
  s.addNotes(`Rebuilt from ${p.source}, page ${p.page}, report date ${p.date}. Figures are exactly as reported on the dashboard; the only derived figures are Total KYC applications (onboarded + incomplete + rejected), the tickets-resolved percentage and the change since the previous report.`);
});

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote", f, pages.length + 2, "slides"));
