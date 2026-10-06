"""Apply Yaqoob's comments to the KSA market discussion deck (Hatem's Canva version).

1. New slide before the channels slide: what we learned from Instagram, X and TikTok.
2. Channels slide: table rebuilt with a Dates column, dates in bold.
3. New slide after Jodel: an eMYAA takeover poll example on personal goals (Arabic, with English).
Source (read-only): input/saudi-entry/eMYAA_KSA_Market_Discussion.pptx
Run from emyaa-ceo-deck/:  python3 working/update_ksa_deck.py
"""
import copy
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = 'input/saudi-entry/eMYAA_KSA_Market_Discussion.pptx'
OUT = 'output/vision-bank/eMYAA_KSA_Market_Discussion_-_v2.pptx'
NAVY, GOLD, GOLDT, GREEN, BLUE = '132257', 'D4AF37', '9A7B1F', '1BA97C', '3C61B6'
PANEL, LINE, TEXT, MUTED, RED = 'EEF1F8', 'D9DFEA', '1F2A44', '6B7590', 'C53030'
JODEL = 'FF9908'
AR_FONT = 'Tajawal'

prs = Presentation(SRC)
slides = list(prs.slides)
E = 914400


def rgb(h):
    return RGBColor.from_string(h)


def box(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    if line:
        sh.line.color.rgb = rgb(line); sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    if radius is not None:
        sh.adjustments[0] = radius
    return sh


def text(s, x, y, w, h, paras, size=18, bold=False, color=TEXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         rtl=False, italic=False, space=0):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(paras, str):
        paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if space:
            p.space_after = Pt(space)
        if rtl:
            p._p.get_or_add_pPr().set('rtl', '1')
        runs = [(para, {})] if isinstance(para, str) else para
        for t, o in runs:
            r = p.add_run(); r.text = t
            b = o.get('bold', bold)
            f = r.font; f.size = Pt(o.get('size', size)); f.bold = b; f.italic = o.get('italic', italic)
            f.color.rgb = rgb(o.get('color', color))
            f.name = 'Exo 2 Bold' if b else 'Exo 2'
            rPr = r._r.get_or_add_rPr()
            for tag, face in (('a:ea', f.name), ('a:cs', AR_FONT if rtl else f.name)):
                el = rPr.find(qn(tag))
                if el is None:
                    el = etree.SubElement(rPr, qn(tag))
                el.set('typeface', face)
            if rtl:
                rPr.set('lang', 'ar-SA')
    return tb


# ---------- slide cloning (keeps the Canva header, footer and logos) ----------
TEMPLATE = slides[5]                       # "For discussion" slide
KEEP = {2, 4, 6, 7, 10, 12, 13, 14}        # band, gold rule, kicker, title, Ajyad, footer text, page no., eMYAA


def clone(kicker, title):
    new = prs.slides.add_slide(TEMPLATE.slide_layout)
    for sp in list(new.shapes._spTree)[2:]:
        new.shapes._spTree.remove(sp)
    for sh in TEMPLATE.shapes:
        if sh.shape_id not in KEEP:
            continue
        el = copy.deepcopy(sh._element)
        for node in el.iter():
            for att in (qn('r:embed'), qn('r:link'), qn('r:id')):
                rid = node.get(att)
                if rid and rid in TEMPLATE.part.rels:
                    rel = TEMPLATE.part.rels[rid]
                    node.set(att, new.part.relate_to(rel._target, rel.reltype))
        new.shapes._spTree.append(el)
    bg = TEMPLATE._element.find(qn('p:cSld')).find(qn('p:bg'))
    if bg is not None:
        new._element.find(qn('p:cSld')).insert(0, copy.deepcopy(bg))
    for sh in new.shapes:
        if sh.shape_id == 6:
            sh.text_frame.paragraphs[0].runs[0].text = kicker
        if sh.shape_id == 7:
            runs = sh._element.findall('.//' + qn('a:r'))
            runs[0].find(qn('a:t')).text = title
            for r in runs[1:]:
                r.getparent().remove(r)
    return new


def card(s, x, y, w, h, top):
    box(s, x, y, w, h, 'FFFFFF', LINE)
    box(s, x, y, w, 0.09, top)


# ---------- 1. Social media feedback ----------
soc = clone('WHAT WE HAVE LEARNED SO FAR', 'Instagram, X and TikTok: What Worked and Why')
cols = [
    ('Instagram', 'Our strongest channel', GREEN, [
        [('2,116 followers', {'bold': True, 'color': NAVY}), (', up from 774 at the end of June', {})],
        [('2.27M views in 90 days', {'bold': True, 'color': NAVY}), (', 72% of interactions from Reels', {})],
        [('Creator reels: ', {'bold': True, 'color': NAVY}), ('Musheera 5.3M and Ali Sabeel 1.3M views (paid boost)', {})],
        [('Mention & Win: ', {'bold': True, 'color': NAVY}), ('266K views and 22,580 clicks for USD 600, 6 new accounts', {})],
    ]),
    ('X (Twitter)', 'Reach, but not in Saudi', GOLDT, [
        [('738 followers', {'bold': True, 'color': NAVY}), (', 96% in Bahrain and only 4% in Saudi', {})],
        [('Saudi city tests ', {'bold': True, 'color': NAVY}), ('in Jeddah, Dammam and Khobar in June', {})],
        [('Engagement fell 42% ', {'bold': True, 'color': NAVY}), ('in June while followers grew', {})],
        [('Paid support paused ', {'bold': True, 'color': NAVY}), ('since July', {})],
    ]),
    ('TikTok', 'Paused', RED, [
        [('15 followers', {'bold': True, 'color': NAVY}), (', 176 total likes', {})],
        [('Bot traffic found ', {'bold': True, 'color': NAVY}), ('in May: 706K fake clicks through a partner ad network', {})],
        [('Paid ads stopped ', {'bold': True, 'color': NAVY}), ('on 22 June; traffic has been clean since', {})],
        [('Audience ', {'bold': True, 'color': NAVY}), ('not yet a fit for investing content', {})],
    ]),
]
for i, (name, verdict, col, items) in enumerate(cols):
    x = 0.92 + i * 6.15
    card(soc, x, 2.6, 5.95, 5.0, col)
    text(soc, x + 0.4, 2.95, 5.2, 0.6, name, 26, True, NAVY)
    text(soc, x + 0.4, 3.55, 5.2, 0.45, verdict.upper(), 14, True, col)
    text(soc, x + 0.4, 4.2, 5.2, 3.3, items, 17.5, False, TEXT, space=12)
box(soc, 0.92, 7.85, 18.2, 2.45, PANEL)
box(soc, 0.92, 7.85, 0.11, 2.45, GOLD)
text(soc, 1.35, 8.05, 6, 0.4, 'WHY INSTAGRAM OUTPERFORMED', 14, True, BLUE)
text(soc, 1.35, 8.55, 10.5, 1.7, [
    [('Real people, not animation: ', {'bold': True, 'color': NAVY}), ('presenter-led creator videos built more trust than graphic explainers.', {})],
    [('Paid boost on every key video: ', {'bold': True, 'color': NAVY}), ('reach came from sponsoring creator content to the right GCC cities and ages.', {})],
    [('Contests that ask for action: ', {'bold': True, 'color': NAVY}), ('Mention & Win turns views into follows and new accounts.', {})],
], 15.5, False, TEXT, space=8)

# reel thumbnails inside the "why" panel
reels = [('tile9_musheera.png', 'Musheera', '5.3M views'), ('tile9_ali.png', 'Ali Sabeel', '1.3M views'),
         ('ksa_reel_mw.png', 'Mention & Win', '186K views'), ('tile9_abdulelah.png', 'Abdulelah Al Harbi', 'October')]
th = 2.15
tw = th * 488 / 628
for k, (img, who, views) in enumerate(reels):
    rx = 19.0 - (4 - k) * tw - (3 - k) * 0.12
    soc.shapes.add_picture('working/charts/' + img, Inches(rx), Inches(7.98), Inches(tw), Inches(th))
    box(soc, rx, 7.98 + th - 0.55, tw, 0.55, NAVY)
    text(soc, rx + 0.08, 7.98 + th - 0.55, tw - 0.16, 0.55, [[(who, {'bold': True, 'size': 10.5, 'color': 'FFFFFF'})], [(views, {'size': 10, 'color': GOLD})]],
         10, False, 'FFFFFF', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

# ---------- 2. Channels slide: table with a Dates column ----------
ch = slides[2]
tree = ch.shapes._spTree
for sh in list(ch.shapes):
    if sh.shape_id >= 16 and sh.shape_id <= 128:
        tree.remove(sh._element)
rows = [
    ['Channel and market', 'Budget', 'Target', 'Dates', 'Status'],
    ['Google Search, Saudi', 'USD 9,000', 'Riyadh first', '22 Sep to 31 Dec', ('Live', GREEN)],
    ['LinkedIn, Saudi', 'USD 9,000', 'Lead generation forms', '6 Oct to 31 Dec', ('Live', GREEN)],
    ['Google Search, rest of GCC', 'USD 6,000', 'Qatar, UAE, Kuwait, Bahrain, Oman', 'Early Oct to 31 Dec', ('To be live', NAVY)],
    ['LinkedIn, rest of GCC', 'USD 6,000', 'Qatar, UAE, Kuwait, Bahrain, Oman', 'Early Oct to 31 Dec', ('To be live', NAVY)],
    ['Total', 'USD 30,000', '', '22 Sep to 31 Dec', ''],
]
colW = [3.85, 2.3, 5.05, 3.55, 3.42]
y = 2.61
for i, r in enumerate(rows):
    h = 0.7 if i == 0 else 0.87
    x = 0.92
    for j, c in enumerate(r):
        fill = NAVY if i == 0 else ('FFFFFF' if i % 2 else PANEL)
        box(ch, x, y, colW[j], h, fill, None if i == 0 else LINE)
        t, col = (c if isinstance(c, tuple) else (c, None))
        if i == 0:
            text(ch, x + 0.2, y, colW[j] - 0.3, h, t, 15.75, True, 'FFFFFF', anchor=MSO_ANCHOR.MIDDLE)
        else:
            bold = j in (0, 1, 3, 4) or i == len(rows) - 1
            color = col or (NAVY if j in (0, 3) else (GOLDT if False else TEXT))
            size = 19.5 if j == 3 else 18.75
            text(ch, x + 0.2, y, colW[j] - 0.3, h, t, size, bold, color, anchor=MSO_ANCHOR.MIDDLE)
        x += colW[j]
    y += h

# ---------- 3. Takeover poll example ----------
poll = clone('JODEL TAKEOVER POLL, EXAMPLE', 'An eMYAA Poll Built on Personal Goals')
# phone
px, py, pw, ph = 1.3, 2.55, 4.3, 7.55
box(poll, px - 0.12, py - 0.12, pw + 0.24, ph + 0.24, '1B1B2A', shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
box(poll, px, py, pw, ph, 'F4F5F7', shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
box(poll, px, py + 0.25, pw, 0.75, JODEL)
text(poll, px + 0.25, py + 0.25, 2.5, 0.75, 'Jodel  |  Riyadh', 15, True, 'FFFFFF', anchor=MSO_ANCHOR.MIDDLE)
text(poll, px + pw - 1.6, py + 0.25, 1.4, 0.75, 'Sponsored', 11, False, 'FFFFFF', PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
box(poll, px + 0.2, py + 1.2, pw - 0.4, 6.05, '3C8BE8', shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
text(poll, px + 0.4, py + 1.35, pw - 0.8, 0.35, 'eMYAA  |  Shari\'ah-compliant investing', 10.5, True, 'FFFFFF')
text(poll, px + 0.4, py + 1.8, pw - 0.8, 1.4, 'لو بديت تستثمر من اليوم، وش أول هدف تبي تحققه؟', 19, True, 'FFFFFF', PP_ALIGN.RIGHT, rtl=True)
opts = ['أسدد ديوني', 'أشتري سيارة', 'أكمل دراستي', 'أتزوج']
for k, o in enumerate(opts):
    oy = py + 3.1 + k * 0.82
    box(poll, px + 0.4, oy, pw - 0.8, 0.65, 'FFFFFF', shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    text(poll, px + 0.6, oy, pw - 1.2, 0.65, o, 16, True, NAVY, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE, rtl=True)
text(poll, px + 0.4, py + 6.6, pw - 0.8, 0.4, 'Live for 24 hours  |  tap to vote', 10.5, False, 'FFFFFF', PP_ALIGN.CENTER)
# right side
x0 = 6.4
text(poll, x0, 2.55, 12.5, 0.4, 'THE POLL, IN ENGLISH', 14, True, BLUE)
text(poll, x0, 3.0, 12.5, 0.6, '"If you started investing today, what is the first goal you would want to reach?"', 21, True, NAVY)
labels = ['Pay off my debts', 'Buy a car', 'Finish my studies', 'Get married']
for k, l in enumerate(labels):
    xx = x0 + (k % 2) * 6.25
    yy = 3.85 + (k // 2) * 0.75
    box(poll, xx, yy, 6.0, 0.6, PANEL)
    box(poll, xx, yy, 0.09, 0.6, JODEL)
    text(poll, xx + 0.3, yy, 5.6, 0.6, l, 17, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
text(poll, x0, 5.55, 12.5, 0.4, 'WHY IT WORKS', 14, True, BLUE)
text(poll, x0, 6.0, 12.5, 2.2, [
    [('Personal, not technical: ', {'bold': True, 'color': NAVY}), ('it links investing to goals young Saudis already think about every day.', {})],
    [('Everyone sees it: ', {'bold': True, 'color': NAVY}), ('a 24-hour takeover poll reaches about 2 million impressions across the Kingdom.', {})],
    [('Easy to answer, easy to share: ', {'bold': True, 'color': NAVY}), ('one tap to vote, and results and comments are live.', {})],
    [('A clear next step: ', {'bold': True, 'color': NAVY}), ('the display ads that follow point voters to eMYAA.', {})],
], 16, False, TEXT, space=8)
box(poll, x0, 8.45, 12.7, 1.85, PANEL)
box(poll, x0, 8.45, 0.11, 1.85, GOLD)
text(poll, x0 + 0.4, 8.6, 12.0, 0.35, 'FOLLOW-UP AD AFTER THE POLL', 13, True, BLUE)
text(poll, x0 + 0.4, 9.0, 12.0, 0.6, 'كل هدف يبدأ بخطوة. ابدأ استثمارك المتوافق مع الشريعة من 50 دولار مع eMYAA', 17, True, NAVY, PP_ALIGN.RIGHT, rtl=True)
text(poll, x0 + 0.4, 9.6, 12.0, 0.5, 'Every goal starts with a step. Start investing the Shari\'ah-compliant way from USD 50 with eMYAA.', 15, False, TEXT, italic=True)
poll.notes_slide.notes_text_frame.text = ('Example copy for a Jodel 24-hour takeover poll (Nadher Media offer, 16 Sep 2026: about 2,000,000 poll '
    'impressions). Arabic in Saudi dialect. To be reviewed by compliance before use: no promise of returns, and the '
    'debt option should not suggest borrowing to invest.')
soc.notes_slide.notes_text_frame.text = ('Sources: Instagram Professional Dashboard (Aug 2026 report: 1,765 followers, 2,271,132 views in 90 days, '
    '72.4% of interactions from Reels; Musheera and Ali Sabeel views are paid boosted); Instagram followers 2,116 (Sep); '
    'June 2026 report (TRACCS / Sprout Social): Instagram 774 followers at 30 Jun, X engagement -41.8%; Mention & Win '
    'pilot 11-18 Jun: USD 600, 266,218 views, 22,580 clicks, 6 accounts; X: 738 followers, 96.4% Bahrain, 3.6% Saudi, '
    'App Screens Saudi tests (Jeddah, Dammam and Khobar) 4-13 Jun, paid support lapsed July; TikTok: 15 followers, 176 likes, '
    'Pangle bot clicks 706,445 in May, paid TikTok removed 22 Jun, zero Pangle clicks July and August.')

# ---------- 4. Marketing activity so far (table) ----------
act = clone('WHAT WE HAVE DONE SO FAR', 'Marketing Activity Since Launch, May to October 2026')
arows = [
    ['Campaign', 'Platform', 'When', 'Spend (USD)', 'Result'],
    ['Long Form Launch', 'Instagram, LinkedIn', '19 May to 1 Jun', '1,700', '231,850 impressions and 3,471 clicks across KSA, Bahrain, Oman, Kuwait'],
    ['Astronaut (animation)', 'Instagram, TikTok, X', '19 May to 1 Jun', '3,700', 'Highest reach in May, but TikTok traffic was bots'],
    ['App Screens', 'X, Saudi only', '4 to 13 Jun', '500', 'Product test in Jeddah, Dammam and Khobar'],
    ['GCC Relaunch', 'Instagram, X, LinkedIn', '8 Jun to 8 Jul', '6,700', 'Followers grew on every channel in June'],
    ['Mention & Win', 'Instagram, GCC', 'Monthly since Jun', '500 prize + ads', 'Pilot: 266K views, 22,580 clicks, 6 new accounts for USD 600'],
    ['Creator videos', 'Instagram', 'Jul to Oct', '9,037', 'Musheera 5.3M and Ali Sabeel 1.3M views; Abdulelah in October'],
    ['USD 5,000 Prize Campaign', 'Instagram, landing pages', '1 Aug to 31 Dec', 'Prizes up to 5,000', 'Live; draw on 10 Jan 2027'],
    ['Google Ads and LinkedIn', 'Saudi, then GCC', 'From 22 Sep', '30,000', '70% of September website visits (see next slides)'],
]
aw = [3.6, 3.0, 2.6, 2.35, 6.62]
y = 2.55
for i, r in enumerate(arows):
    h = 0.62 if i == 0 else 0.8
    x = 0.92
    for j, c in enumerate(r):
        fill = NAVY if i == 0 else ('FFFFFF' if i % 2 else PANEL)
        box(act, x, y, aw[j], h, fill, None if i == 0 else LINE)
        if i == 0:
            text(act, x + 0.18, y, aw[j] - 0.3, h, c, 15, True, 'FFFFFF', anchor=MSO_ANCHOR.MIDDLE)
        else:
            text(act, x + 0.18, y, aw[j] - 0.3, h, c, 15.5 if j < 4 else 14.5, j in (0, 3), NAVY if j in (0, 3) else TEXT, anchor=MSO_ANCHOR.MIDDLE)
        x += aw[j]
    y += h
box(act, 0.92, 9.75, 18.17, 0.6, PANEL)
box(act, 0.92, 9.75, 0.11, 0.6, GOLD)
text(act, 1.3, 9.75, 17.6, 0.6, [[('Spent to date: USD 34,482 ', {'bold': True, 'color': NAVY}), ('(14 May to 6 Sep, from a USD 80,000 allocated budget), plus USD 30,000 for Google Ads and LinkedIn from Sep to Dec.', {})]],
     14.5, False, TEXT, anchor=MSO_ANCHOR.MIDDLE)
act.notes_slide.notes_text_frame.text = ('Sources: July 2026 monthly report (sponsored ads May to July: Long Form Launch, Astronaut, App Screens, '
    'GCC Relaunch, Mention & Win pilot); Board update 17 Sep 2026 (marketing budget: USD 80,000 allocated, USD 34,482 spent '
    '14 May to 6 Sep, influencer programme USD 9,037; USD 30,000 Google Ads and LinkedIn Sep to Dec); USD 5,000 campaign T&Cs; '
    'GA4 1-28 Sep. GCC Relaunch results were not reported separately.')

# ---------- order and page numbers ----------
lst = prs.slides._sldIdLst
ids = list(lst)
old, new = ids[:7], ids[7:]          # new = [social, poll, activity]
order = old[:2] + [new[0], new[2]] + old[2:5] + [new[1]] + old[5:]
for e in ids:
    lst.remove(e)
for e in order:
    lst.append(e)
for n, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().isdigit() and sh.top / E > 10.3:
            sh.text_frame.paragraphs[0].runs[0].text = str(n)
prs.save(OUT)
print('saved', OUT, len(order), 'slides')
