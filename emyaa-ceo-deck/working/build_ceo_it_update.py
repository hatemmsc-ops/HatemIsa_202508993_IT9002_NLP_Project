"""Add the IT status update (01 Oct 2026) to the CEO update deck (Performance Analysis).

Base: working/ceo-update-base-29Sep.pptx (the 29 Sep update, before IT changes).
Source: input/source/it/IT_Status_Update_CEO_01_10_2026.pptx and the comments in it.
New slides are drawn from plain shapes and tables (Canva-safe) in the deck's own style.
Run from emyaa-ceo-deck/:  python3 working/build_ceo_it_update.py
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

BASE = 'working/ceo-update-base-29Sep.pptx'
OUT = 'output/ceo-update/eMYAA_Performance_Analysis_-_01.10.2026_-_Updated_29_Sep.pptx'

NAVY, BLUE, INK, GREY, MUTED = '121C36', '3C61B6', '1B1B3A', '5A6A85', '8A93A8'
PANEL, PANEL_LINE, RULE = 'F7F8FC', 'E4E7F2', 'C5CED8'
GREEN, AMBER, GOLD, RED = '1E9E6B', 'D98C1F', 'D4AF37', 'C0392B'
STATUS = {  # text colour, cell fill
    'DONE': (GREEN, 'E6F5EE'), 'RESOLVED': (GREEN, 'E6F5EE'),
    'IN PROGRESS': (AMBER, 'FDF3E4'), 'PLANNED': (BLUE, 'E8EEFA'),
    'EXPLORING': (GREY, 'EEF1F6'), 'NEW': (BLUE, 'E8EEFA'),
}

prs = Presentation(BASE)
slides = list(prs.slides)


# ---------- helpers ----------
def rgb(h):
    return RGBColor.from_string(h)


def find(group, sid):
    for sh in group.shapes:
        if sh.shape_id == sid:
            return sh
        if sh.shape_type == 6:
            r = find(sh, sid)
            if r is not None:
                return r
    return None


def set_text(shape, text, para=0):
    """Replace a paragraph's text, keeping the first run's formatting."""
    p = shape.text_frame.paragraphs[para]
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    return runs[0]


def font(run, size, bold=False, color=INK, italic=False):
    f = run.font
    f.name = 'Exo 2 Bold' if bold else 'Exo 2'
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', f.name)


def text(slide, x, y, w, h, paras, size=16, bold=False, color=INK, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=None):
    """paras: str, or list of paragraphs; a paragraph is str or list of (text, dict) runs."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if isinstance(paras, str):
        paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.space_after = Pt(spacing)
        runs = [(para, {})] if isinstance(para, str) else para
        for t, opt in runs:
            r = p.add_run()
            r.text = t
            font(r, opt.get('size', size), opt.get('bold', bold), opt.get('color', color), opt.get('italic', False))
    return tb


def box(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    s.text_frame.text = ''
    return s


def cell_border(cell, color=RULE, w=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=str(w), cap='flat', cmpd='sng', algn='ctr')
        sf = etree.SubElement(ln, qn('a:solidFill'))
        etree.SubElement(sf, qn('a:srgbClr'), val=color)
    # schema order: borders must come before the fill
    fill = tcPr.find(qn('a:solidFill'))
    if fill is not None:
        tcPr.remove(fill)
        tcPr.append(fill)


def table(slide, x, y, widths, rows, header_size=15, body_size=14, row_h=0.62, status_col=None,
          bold_col=0, aligns=None):
    n, m = len(rows), len(widths)
    gf = slide.shapes.add_table(n, m, Inches(x), Inches(y), Inches(sum(widths)), Inches(row_h * n))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for a in ('firstRow', 'bandRow'):
        tblPr.set(a, '0')
    sid = tblPr.find(qn('a:tableStyleId'))
    if sid is not None:
        tblPr.remove(sid)
    for j, w in enumerate(widths):
        tbl.columns[j].width = Inches(w)
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(row_h)
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.14)
            c.margin_top = c.margin_bottom = Inches(0.06)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (aligns[j] if aligns else PP_ALIGN.LEFT)
            r = p.add_run()
            r.text = val
            if i == 0:
                c.fill.solid(); c.fill.fore_color.rgb = rgb(NAVY)
                font(r, header_size, True, 'FFFFFF')
            elif status_col is not None and j == status_col:
                fg, bg = STATUS[val]
                c.fill.solid(); c.fill.fore_color.rgb = rgb(bg)
                p.alignment = PP_ALIGN.CENTER
                font(r, body_size - 1, True, fg)
            else:
                c.fill.solid(); c.fill.fore_color.rgb = rgb('FFFFFF' if i % 2 else PANEL)
                font(r, body_size, j == bold_col, NAVY if j == bold_col else GREY)
            cell_border(c)
    return gf


def clone_slide(src, title):
    """New slide from src keeping its header (title group, accent bar, logos); body removed."""
    new = prs.slides.add_slide(src.slide_layout)
    for sp in list(new.shapes._spTree):
        if sp.tag.endswith('}sp') or sp.tag.endswith('}grpSp') or sp.tag.endswith('}pic'):
            new.shapes._spTree.remove(sp)
    rid_map = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype.endswith('/image'):
            rid_map[rid] = new.part.relate_to(rel._target, rel.reltype)
    keep = {2, 3, 4, 7}  # footer logo, top logo, title group, accent bar (slide 14 ids)
    for sh in src.shapes:
        if sh.shape_id in keep:
            el = copy.deepcopy(sh._element)
            for node in el.iter():
                for att in (qn('r:embed'), qn('r:link'), qn('r:id')):
                    if node.get(att) in rid_map:
                        node.set(att, rid_map[node.get(att)])
            new.shapes._spTree.append(el)
    bg = src._element.find(qn('p:cSld')).find(qn('p:bg'))
    if bg is not None:
        new._element.find(qn('p:cSld')).insert(0, copy.deepcopy(bg))
    set_text(find(new, 6), title)
    return new


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


# ---------- new slides ----------
tmpl = slides[13]  # "IT Progress Update / Exante" slide: header, bar and logos

# A. IT status at a glance
a = clone_slide(tmpl, 'IT Status at a Glance')
text(a, 0.69, 1.72, 18.6, 0.4, 'VENDORS, PLATFORM AND INTERNAL SYSTEMS  |  AS AT 01 OCTOBER 2026', 13, True, BLUE)
rows = [
    ['Area', 'Update', 'Target / done', 'Status'],
    ['iGentech', 'Invoice overrun resolved: billing cut from USD 10,750 to USD 5,625 (-48%).', 'Closed', 'RESOLVED'],
    ['Shufti Pro, billing', 'Duplicate KYC charges refunded (USD 9,605), plus a USD 50 production test credit given as compensation.', 'Closed', 'RESOLVED'],
    ['Google Play security update (current app)', 'Security API update Google requires for the current Android app. Not the Next App. 2 test rounds, 99.4% pass; live on Google Play.', 'Done 29 Sep (deadline 01 Nov)', 'DONE'],
    ['Helpdesk', 'Deployed ahead of the 20 Sep target. Improvement backlog in the appendix.', 'Done 17 Sep', 'DONE'],
    ['Shufti Pro, KYC demo', 'Enhanced KYC flow in the Shufti SDK demo app. Internal demo only; it will not go live.', '08 Oct 2026', 'IN PROGRESS'],
    ['QATestLab', 'Reviewing the Next App and preparing test cases, ready for v1.', 'Testing from 08 Oct', 'IN PROGRESS'],
    ['Exante, Next App', 'First version due 08 Oct. Full release plan to be shared next week. Sanity test date to be confirmed.', 'v1: 08 Oct 2026', 'IN PROGRESS'],
    ['WealthTech CRM (Frappe)', 'One customer base in a single system. Starts now that Helpdesk is live.', '29 Oct 2026', 'PLANNED'],
    ['Reporting Dashboard', 'Real-time reporting options under review to support the business case.', 'TBC', 'EXPLORING'],
]
table(a, 0.69, 2.22, [3.9, 9.45, 3.0, 2.27], rows, header_size=15, body_size=14, row_h=0.8, status_col=3,
      aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
notes(a, 'Source: IT Status Update to CEO, 01 Oct 2026, slide 2, with the review comments applied: '
         'Shufti credit shown as compensation; CRM and Reporting Dashboard added in the same format; '
         'target date column added. The Google Play item is the security update to the current app, '
         'not the Next App. The Shufti KYC flow is an internal demo and will not go live. '
         'The sanity test date for Next App v1 was left blank in the source, so it shows as to be confirmed. '
         'iGentech: 10,750 to 5,625 is -47.7%.')

# B. Next App timeline
b = clone_slide(tmpl, 'Next App: Exante and QATestLab')
text(b, 0.69, 1.72, 18.6, 0.4, 'PLATFORM  |  NEXT APP DELIVERY', 13, True, BLUE)
steps = [
    ('Now', 'QATestLab engaged', 'Reviewing the Next App and writing test cases while Exante builds it.'),
    ('Next week', 'Release plan', 'Full plan shared once QATestLab confirms its estimate.'),
    ('08 Oct 2026', 'Next App v1', 'First version delivered by Exante.'),
    ('After 08 Oct', 'Independent QA', 'QATestLab tests v1 against the prepared cases. Sanity test date to be confirmed.'),
]
x0, gap, cw = 0.69, 0.3, (18.62 - 0.3 * 3) / 4
line_y = 3.05
box(b, x0 + cw / 2, line_y - 0.02, cw * 3 + gap * 3, 0.04, PANEL_LINE)
for i, (when, head, body) in enumerate(steps):
    x = x0 + i * (cw + gap)
    text(b, x, 2.3, cw, 0.4, when, 16, True, BLUE, PP_ALIGN.CENTER)
    dot = box(b, x + cw / 2 - 0.2, line_y - 0.2, 0.4, 0.4, BLUE if i == 0 else 'FFFFFF', None if i == 0 else BLUE, MSO_SHAPE.OVAL)
    if i:
        dot.line.width = Pt(2.5)
    box(b, x, 3.55, cw, 3.1, PANEL, PANEL_LINE)
    box(b, x, 3.55, cw, 0.08, NAVY if i != 2 else GOLD)
    text(b, x + 0.3, 3.9, cw - 0.6, 0.5, head, 23, True, NAVY)
    text(b, x + 0.3, 4.65, cw - 0.6, 1.9, body, 18, False, GREY)
box(b, 0.69, 7.05, 18.62, 2.45, PANEL, PANEL_LINE)
box(b, 0.69, 7.05, 0.1, 2.45, BLUE)
text(b, 1.1, 7.35, 17.9, 0.4, 'WHY IT MATTERS', 15, True, BLUE)
text(b, 1.1, 7.9, 17.9, 1.4,
     "QATestLab gives an independent quality check before the Next App reaches users. Test cases are written "
     "while Exante builds v1, so testing can start as soon as v1 arrives on 08 Oct.", 21, False, INK)
notes(b, 'Source: IT Status Update to CEO, 01 Oct 2026, slide 3. The source said testing starts when v1 arrives '
         'on 09 Oct but gives 08 Oct everywhere else; 08 Oct is used. Sanity test date not given in the source.')

# C. Android security update
c = clone_slide(tmpl, 'Android App: Google Play Security Update')
text(c, 0.69, 1.72, 18.6, 0.4, 'CURRENT APP  |  SECURITY COMPLIANCE  |  NOT THE NEXT APP', 13, True, BLUE)
text(c, 0.69, 2.2, 18.62, 0.9,
     'Google required a security API update, republished by 01 Nov 2026, to keep the current app on Google Play. '
     'We ran two internal test rounds and found no major bugs.', 17, False, INK)
# left: pass rate and result bar
box(c, 0.69, 3.3, 5.6, 6.1, PANEL, PANEL_LINE)
text(c, 0.99, 3.55, 5.0, 1.2, '99.4%', 60, True, GREEN)
text(c, 0.99, 4.8, 5.0, 0.4, 'pass rate (158 of 159 executed)', 16, False, GREY)
parts = [('Passed', 158, GREEN), ('Skipped', 19, AMBER), ('N/A', 4, MUTED), ('Failed', 1, RED)]
bx, bw = 0.99, 5.0
for label, v, col in parts:
    w = bw * v / 182
    box(c, bx, 5.5, max(w, 0.04), 0.5, col)
    bx += w
for i, (label, v, col) in enumerate(parts):
    yy = 6.3 + (i // 2) * 0.55
    xx = 0.99 + (i % 2) * 2.5
    box(c, xx, yy + 0.08, 0.22, 0.22, col)
    text(c, xx + 0.35, yy, 2.1, 0.4, f'{label} {v}', 16, False, INK)
text(c, 0.99, 7.55, 5.0, 0.4, '182 cases  |  159 executed  |  158 passed', 14, False, GREY)
box(c, 0.99, 8.2, 5.0, 0.95, 'E6F5EE')
text(c, 1.19, 8.25, 4.7, 0.85, [[('Live on Google Play: 29 Sep 2026', {'bold': True, 'color': GREEN, 'size': 15})],
                              [('More than a month before the deadline.', {'size': 13, 'color': GREY})]], anchor=MSO_ANCHOR.MIDDLE)
# middle: results by module
mrows = [['Module', 'Passed', 'Failed', 'Skipped', 'N/A'],
         ['Base requirements', '48', '1', '13', '2'],
         ['AI Chatbot', '31', '0', '0', '2'],
         ['T&C + SCC', '12', '0', '0', '0'],
         ['Other Q1 requirements', '42', '0', '3', '0'],
         ['KYC onboarding', '25', '0', '3', '0'],
         ['Total', '158', '1', '19', '4']]
C = PP_ALIGN.CENTER
tg = table(c, 6.6, 3.3, [3.1, 1.15, 1.1, 1.2, 0.95], mrows, 15, 15, row_h=0.78, aligns=[PP_ALIGN.LEFT, C, C, C, C])
t = tg.table
for j in range(5):
    r = t.cell(6, j).text_frame.paragraphs[0].runs[0]
    r.font.bold = True; r.font.name = 'Exo 2 Bold'; r.font.color.rgb = rgb(NAVY)
r = t.cell(1, 2).text_frame.paragraphs[0].runs[0]
r.font.color.rgb = rgb(RED); r.font.bold = True; r.font.name = 'Exo 2 Bold'
text(c, 6.6, 8.95, 7.5, 0.4, 'Round 1 and Round 2 combined.', 13, False, MUTED)
# right: minor issues
box(c, 14.4, 3.3, 4.91, 6.1, PANEL, PANEL_LINE)
text(c, 14.65, 3.55, 4.5, 0.4, 'MINOR ISSUES, NEXT RELEASE', 14, True, BLUE)
issues = ['Watchlist items disappear when the Shariah filter is on',
          'Instrument tree keeps loading on first open of the trade screen',
          'Withdrawal: beneficiary shows last name only',
          'Change display mode has no effect',
          'Arabic summary screen: daily P/L hard to see',
          'Exchange section should be hidden (USD base currency only)']
text(c, 14.65, 4.1, 4.45, 4.5, [[(f'{i + 1}.  ', {'bold': True, 'color': NAVY}), (s, {})] for i, s in enumerate(issues)],
     15, False, GREY, spacing=9)
text(c, 14.65, 8.75, 4.45, 0.5, 'None of these block the release.', 15, True, GREEN)
notes(c, 'Source: IT Status Update to CEO, 01 Oct 2026, slide 4. This is the security API update Google requires '
         'for the current Android app, not the Next App. 182 cases: 158 passed, 19 skipped, 4 N/A, 1 failed; '
         '158 of 159 executed = 99.4%. Module rows add up to the totals. Chart drawn from shapes so it imports into Canva. '
         'Tester names removed, in line with the review comment on the Helpdesk slide.')

# D. Helpdesk backlog (appendix)
d = clone_slide(tmpl, 'Helpdesk: Improvement Backlog')
text(d, 0.69, 1.72, 18.6, 0.4, 'FEEDBACK FROM THE TEAMS SINCE GO-LIVE ON 17 SEP', 13, True, BLUE)
cards = [('7', 'Items logged', NAVY), ('3', 'Completed', GREEN), ('1', 'In progress', AMBER), ('3', 'New', BLUE)]
for i, (v, lbl, col) in enumerate(cards):
    y = 2.3 + i * 1.62
    box(d, 0.69, y, 3.4, 1.42, PANEL, PANEL_LINE)
    box(d, 0.69, y, 0.08, 1.42, col)
    text(d, 1.0, y + 0.12, 3.0, 0.75, v, 36, True, col)
    text(d, 1.0, y + 0.9, 3.0, 0.4, lbl, 15, False, GREY)
hrows = [['Improvement', 'Priority', 'Status'],
         ['Standalone desktop app for Client Experience', 'Low', 'DONE'],
         ['Auto-acknowledgement email with ticket ID', 'Medium', 'DONE'],
         ['Notify Client Experience of new tickets', 'Medium', 'DONE'],
         ['Pause SLA timers on weekends', 'Low', 'IN PROGRESS'],
         ['Knowledge base linked to the eMYAA support page', 'Low', 'NEW'],
         ['Create a ticket when a chatbot user picks "Contact a Human"', 'Low', 'NEW'],
         ['CRM link: create Helpdesk accounts for new users automatically', 'Low', 'NEW']]
table(d, 4.45, 2.3, [10.3, 1.9, 2.66], hrows, 15, 15, row_h=0.78, status_col=2, bold_col=None,
      aligns=[PP_ALIGN.LEFT, C, C])
box(d, 4.45, 8.85, 14.86, 0.75, PANEL)
text(d, 4.7, 8.85, 14.4, 0.75, 'The weekend SLA pause is being built after an SLA breach on a client query raised on a Thursday.',
     15, False, INK, anchor=MSO_ANCHOR.MIDDLE)
notes(d, 'Source: IT Status Update to CEO, 01 Oct 2026, slide 5. Names and IDs removed and slide moved to the appendix, '
         'as the review comments asked. Counts: 7 logged = 3 done + 1 in progress + 3 new.')


# ---------- edits to existing slides ----------
# Slide 2, Key Highlights: IT line

# Slide 5, IT Progress Update: product roadmap
s5 = slides[4]
set_text(find(s5, 4), 'IT Progress Update: Product Roadmap')
r = find(s5, 15).text_frame.paragraphs[0].runs
r[0].text = 'The Next Mobile App - '; r[1].text = 'v1 due 08 Oct 2026'
set_text(find(s5, 18), 'Built by Exante and tested by QATestLab before release. A simpler, easier trading experience.')
find(s5, 24).text_frame.paragraphs[0].runs[0].text = 'Guest Access - '
r = find(s5, 33).text_frame.paragraphs[0].runs
r[0].text = 'Market Pulse Feature - '; r[1].text = ''

# Slide 12, Platform Initiatives Roadmap (appendix): Helpdesk is live
s12 = slides[11]
pill = find(s12, 7)
spPr = pill._element.spPr
blip = spPr.find(qn('a:blipFill'))
sf = etree.Element(qn('a:solidFill')); etree.SubElement(sf, qn('a:srgbClr'), val=GREEN)
blip.addprevious(sf); spPr.remove(blip)
set_text(find(s12, 8), 'DONE')
set_text(find(s12, 11), 'Helpdesk Deployment - Frappe Helpdesk')
set_text(find(s12, 14), 'Deployed by WealthTech IT on 17 Sep 2026, ahead of the 20 Sep target. Cybersecurity approval was received on 20 August.')
set_text(find(s12, 24), 'WealthTech CRM - Frappe CRM')
set_text(find(s12, 27), "Cybersecurity has approved Frappe applications. Frappe's CRM module will hold one customer base. Starts now that Helpdesk is live.")
set_text(find(s12, 66), 'Under R&D - date TBC')
run = set_text(find(s12, 69), 'Live since 17 Sep 2026')
run.font.color.rgb = rgb(GREEN)

# Slide 13, Systems Planning & Deployment (appendix): Helpdesk planning and deployment done
s13 = slides[12]
tree = s13.shapes._spTree
for gid in (281, 283, 285):                    # orange "IN PROGRESS" pill laid over Helpdesk planning
    tree.remove(find(s13, gid)._element)
for clr in find(s13, 68)._element.iter(qn('a:srgbClr')):   # Helpdesk deployment pill: amber to green
    if clr.get('val') == 'FBF0DE':
        clr.set('val', 'E4F5EC')
run = set_text(find(s13, 76), 'DONE')
run.font.color.rgb = rgb(GREEN)
set_text(find(s13, 272), 'Deployed on 17 Sep 2026, ahead of the 20 Sep target.')
set_text(find(s13, 273), 'Will bring the customer base into a single system. Starts now that Helpdesk is live.')
set_text(find(s13, 267), 'Helpdesk is live: Client Experience no longer tracks queries in Excel')
set_text(find(s13, 276), 'eMYAA Trading Platform  |  CEO Update  |  01 Oct 2026')
tree.remove(find(s13, 277)._element)          # stray page number "6" from the Board file
r = find(s13, 238).text_frame.paragraphs[0].runs
r[1].text = 'Annual cost of USD 600-1,000 (hosting only), about 83%-90% less than comparable vendor tools at USD 6,000-10,000.'

# Slide 14, Exante status (appendix)
s14 = slides[13]
r = find(s14, 14).text_frame.paragraphs[0].runs
r[0].text = 'Exante - Status Update'; r[1].text = ''
set_text(find(s14, 22), 'First version of the Next App due 08 Oct 2026. Full release plan to be shared next week.', 1)
set_text(find(s14, 26), 'Sanity / Smoke Tests', 0)
set_text(find(s14, 26), 'The Android security update passed two internal test rounds (99.4%) and went live on Google Play on 29 Sep. Sanity test date for Next App v1: to be confirmed.', 1)


# ---------- Client Experience (17 Sep to 1 Oct; tracker weeks 8-9 hold every client in that period) ----------
# Team figures: 21 called, 72 emailed, 4 WhatsApp, 6 completed onboarding.
# Tracker weeks 8-9 (working/audit-data/cx-tracker-weekly.csv): 21 clients = 17 not onboarded + 4 funding gap;
# 12 reached, 2 no answer, 7 not yet logged. The 6 who onboarded come from the 17, leaving 11.
def card(slide, x, y, w, h, value, label, vcol='FFFFFF'):
    s = box(slide, x, y, w, h, NAVY)
    s.fill.gradient()
    s.fill.gradient_angle = 45
    st = s.fill.gradient_stops
    st[0].color.rgb = rgb('2B4AA0'); st[0].position = 0
    st[1].color.rgb = rgb(NAVY); st[1].position = 1.0
    text(slide, x, y + 0.2, w, 0.8, value, 36, True, vcol, PP_ALIGN.CENTER)
    text(slide, x + 0.15, y + 1.05, w - 0.3, 0.6, label, 14, False, 'FFFFFF', PP_ALIGN.CENTER)


e = clone_slide(tmpl, 'Client Experience Overview')
text(e, 0.69, 1.72, 18.6, 0.4, 'SINCE THE LAST CEO UPDATE  |  17 SEP TO 1 OCT 2026', 13, True, BLUE)
kpis = [('21', 'Clients called', 'FFFFFF'), ('72', 'Clients emailed', 'FFFFFF'), ('4', 'WhatsApp clients', 'FFFFFF'),
        ('6', 'Completed onboarding', '4ADE9A'), ('15', 'Still pending', 'FF7A6B'), ('28.6%', 'Conversion (of clients called)', '4ADE9A')]
kw = (18.62 - 5 * 0.22) / 6
for i, (v, l, col) in enumerate(kpis):
    card(e, 0.69 + i * (kw + 0.22), 2.25, kw, 1.75, v, l, col)
text(e, 0.69, 4.35, 18.6, 0.4, 'CLIENTS CALLED, BY WHERE THEY ARE NOW', 13, True, BLUE)
groups = [(6, 'Onboarded', 'Completed onboarding after the call', GREEN),
          (11, 'Not onboarded', 'KYC or documents not finished', AMBER),
          (4, 'Funding gap', 'Onboarded, not yet funded: need support to fund', BLUE)]
bx = 0.69
for n, lbl, _, col in groups:
    w = 18.62 * n / 21
    box(e, bx, 4.8, w, 0.55, col)
    text(e, bx, 4.8, w, 0.55, f'{lbl}  {n}', 14, True, 'FFFFFF', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    bx += w
gw = (18.62 - 2 * 0.3) / 3
for i, (n, lbl, desc, col) in enumerate(groups):
    x = 0.69 + i * (gw + 0.3)
    box(e, x, 5.6, gw, 1.55, PANEL, PANEL_LINE)
    box(e, x, 5.6, 0.08, 1.55, col)
    text(e, x + 0.35, 5.72, 1.6, 0.8, str(n), 36, True, col)
    text(e, x + 1.85, 5.8, gw - 2.1, 0.45, f'{lbl}  ({n / 21:.1%})', 18, True, NAVY)
    text(e, x + 1.85, 6.3, gw - 2.1, 0.8, desc, 14, False, GREY)
text(e, 0.69, 7.45, 18.6, 0.4, 'SUMMARY', 13, True, BLUE)
text(e, 0.69, 7.9, 18.62, 1.6, [
    [('Out of 21 clients called, 6 completed onboarding (28.6%). ', {'bold': True, 'color': NAVY}),
     ('15 are still pending: 11 have not finished onboarding and 4 are onboarded but have not funded yet.', {})],
    [('Call results in the tracker: ', {'bold': True, 'color': NAVY}),
     ('12 reached, 2 no answer, 7 not yet logged. Email (72) and WhatsApp (4) cover clients we could not reach by phone.', {})]],
     16, False, GREY, spacing=8)
notes(e, 'Team figures to 1 Oct: 21 called, 72 emailed, 4 WhatsApp, 6 completed onboarding. Pending 21 - 6 = 15; '
         'conversion 6 / 21 = 28.6%. Client types from the Client Monthly Tracker, weeks 8 and 9 (20 Sep to 1 Oct; no calls were logged on 17 Sep, and 18-19 Sep is the weekend): '
         '17 not onboarded (Potential Customer) and 4 funding gap (Funded Gap Customer). The 6 who completed onboarding '
         'are among the 17, which leaves 11. The previous Canva version said 6 of 20 = 20% and 14 = 80%; '
         'the correct figures for 20 would have been 30% and 70%.')

f = clone_slide(tmpl, 'Client Feedback and Next Steps')
text(f, 0.69, 1.72, 18.6, 0.4, 'WHAT CLIENTS TOLD US AND WHAT WE DO NEXT', 13, True, BLUE)
box(f, 0.69, 2.25, 9.15, 6.3, PANEL, PANEL_LINE)
text(f, 1.0, 2.45, 8.6, 0.4, 'CLIENT FEEDBACK', 14, True, BLUE)
fb = [('Missing documents: ', "many clients don't have their KYC documents ready at registration, which causes delays or drop-off."),
      ('KYC feels too long: ', 'verification takes longer than clients expect, and some stop halfway.'),
      ('Fees: ', 'clients asked about our commissions and fees.'),
      ('Deposits: ', 'several clients asked how to deposit by bank transfer.'),
      ('Funding gap: ', '4 onboarded clients need support to fund. One asked about fees; two said they will fund soon.')]
text(f, 1.0, 2.95, 8.55, 3.4, [[(a, {'bold': True, 'color': NAVY}), (b, {})] for a, b in fb], 15, False, GREY, spacing=7)
# weekly follow-ups from the tracker (clients with a client type, per week)
import csv as _csv
wk = list(_csv.DictReader(open('working/audit-data/cx-tracker-weekly.csv')))
text(f, 1.0, 6.2, 8.6, 0.4, 'CLIENTS FOLLOWED UP PER WEEK (TRACKER)', 12, True, BLUE)
base, top, bw_, gap_ = 8.0, 6.75, 0.72, 0.21
mx = max(int(w['clients']) for w in wk)
for i, w in enumerate(wk):
    n = int(w['clients'])
    x = 1.05 + i * (bw_ + gap_)
    hgt = (base - top - 0.3) * n / mx
    col = BLUE if i >= len(wk) - 2 else 'A9B6D6'
    if n:
        box(f, x, base - hgt, bw_, hgt, col)
    text(f, x - 0.1, base - hgt - 0.3, bw_ + 0.2, 0.28, str(n), 11, True, NAVY, PP_ALIGN.CENTER)
    d = w['week_ending'][8:10].lstrip('0') + (' Aug' if w['week_ending'][5:7] == '08' else ' Sep' if w['week_ending'][5:7] == '09' else ' Oct')
    text(f, x - 0.15, base + 0.05, bw_ + 0.3, 0.28, d, 10, False, GREY, PP_ALIGN.CENTER)
box(f, 1.05, base, 8.4, 0.02, RULE)
text(f, 1.0, 8.28, 8.6, 0.28, 'Week ending. Dark bars = this update (21). Early August weeks were mostly spam numbers.', 10, False, MUTED)
box(f, 10.16, 2.25, 9.15, 6.3, PANEL, PANEL_LINE)
text(f, 10.47, 2.45, 8.6, 0.4, 'PROPOSED NEXT STEPS, BY GROUP', 14, True, BLUE)
steps = [('Funding gap (4)', 'Send the fee schedule and bank transfer steps, then a follow-up call.', BLUE),
         ('Not onboarded (11)', 'Onboarding email with the document checklist, then a second call on day 3.', AMBER),
         ('Onboarded (6)', 'Guide the first deposit by bank transfer so the account gets funded.', GREEN)]
for i, (h_, b_, col) in enumerate(steps):
    y = 3.0 + i * 1.8
    box(f, 10.47, y, 8.53, 1.55, 'FFFFFF', PANEL_LINE)
    box(f, 10.47, y, 0.08, 1.55, col)
    text(f, 10.8, y + 0.2, 8.0, 0.45, h_, 18, True, col)
    text(f, 10.8, y + 0.72, 8.0, 0.8, b_, 15, False, GREY)
box(f, 0.69, 8.8, 18.62, 1.05, 'EEF1F8')
text(f, 1.0, 8.8, 18.1, 1.05, [[('To make the weekly update faster: ', {'bold': True, 'color': NAVY}),
     ('log the result of every call (7 of 21 have none yet), and give each open client an owner, a due date and a next '
      'follow-up (none set in the last two weeks).', {})]], 15, False, GREY, anchor=MSO_ANCHOR.MIDDLE)
notes(f, 'Feedback from the current Canva version of the Overview slide, plus the funding-gap point asked for in the '
         'review comments. Funding-gap notes from the tracker (week 8): one asked about fees, two will fund soon, one '
         'with no note. Next steps are proposals that follow the Client Interaction Journey in the appendix. Tracker '
         'check: weeks 8-9 have no owner, due date, next follow-up or funding value filled in, and 7 week-9 rows '
         'have no stage.')

set_text(find(slides[1], 140), '6 of 21 clients called onboarded (28.6%); Helpdesk live 17 Sep; Android security update live 29 Sep; Next App v1 due 08 Oct.', 1)

# ---------- order ----------
lst = prs.slides._sldIdLst
ids = list(lst)
old, new = ids[:19], ids[19:]            # new = [A, B, C, D, E, F]
order = old[:4] + new[:3] + [old[4]] + [old[5]] + new[4:6] + old[7:13] + [new[3]] + old[13:]
drop = old[6]                            # old Overview slide, replaced by E and F
prs.part.drop_rel(drop.get(qn('r:id')))
for e in ids:
    lst.remove(e)
for e in order:
    lst.append(e)
prs.save(OUT)
print('saved', OUT, len(order), 'slides')
