"""Make the CEO Progress Update (the Canva export of 01.10.2026) consistent and Canva-safe.

1. Every content slide gets the same header: navy gradient band with a gold rule, gold section
   label, white title and white Ajyad logo; a small blue subtitle line under the band; eMYAA logo bottom right.
2. Every SVG logo or icon is replaced with a high-resolution PNG (Canva shows SVG image fills as a blue box).
3. Every gradient fill on a shape is replaced with a PNG of the same gradient (Canva turns them light blue);
   gradient text becomes solid blue.
Source (read-only): input/source/eMYAA_Progress_Update_CEO_-_01.10.2026.pptx
Run from emyaa-ceo-deck/:  python3 working/canva_fix_progress_update.py
"""
import copy, io, math, re
import cairosvg
from PIL import Image
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

SRC = 'input/source/eMYAA_Progress_Update_CEO_-_01.10.2026.pptx'
OUT = 'output/ceo-update/eMYAA_Progress_Update_CEO_-_01.10.2026_-_Canva-ready.pptx'
EMU = 914400
TITLE_BLUE = '1F3B8C'
SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'

prs = Presentation(SRC)
slides = list(prs.slides)
REF = slides[6]                     # "Next App" slide: the deck's standard header
REF_IDS = {'ajyad': 2, 'title_a': 4, 'title': 6, 'sub': 9, 'emyaa': 54, 'bar': 55}


def shape_by_id(slide, sid):
    return next(sh for sh in slide.shapes if sh.shape_id == sid)


def copy_into(src_slide, dst_slide, el):
    """Deep-copy an element into another slide, carrying its image relationships."""
    el = copy.deepcopy(el)
    for node in el.iter():
        for att in (qn('r:embed'), qn('r:link'), qn('r:id')):
            rid = node.get(att)
            if rid and rid in src_slide.part.rels:
                rel = src_slide.part.rels[rid]
                node.set(att, dst_slide.part.relate_to(rel._target, rel.reltype))
    dst_slide.shapes._spTree.append(el)
    return el


def renumber_ids(slide):
    for i, e in enumerate(slide._element.iter(qn('p:cNvPr')), 2):
        e.set('id', str(i))


def set_group_text(el, text):
    runs = el.findall('.//' + qn('a:r'))
    runs[0].find(qn('a:t')).text = text
    for r in runs[1:]:
        r.getparent().remove(r)


# band header (navy gradient band, gold rule, gold kicker, white title, white Ajyad logo), from slide 18
BAND_SRC = slides[17]
BAND_EL = {k: copy.deepcopy(shape_by_id(BAND_SRC, i)._element)
           for k, i in {'band': 2, 'rule': 4, 'title': 6, 'ajyad': 9, 'kicker': 13}.items()}

# ---------- 1. one header style ----------
BAND = {  # slide number: (title, subtitle)
    14: ('Marketing Efforts', None),
    15: ('Marketing Results', None),
    16: ('English vs Arabic', None),
    17: ('Abdulelah Masterclass & Google Ads', 'INFLUENCER VIDEO AND CAMPAIGN UPDATE  |  FOR WEALTHTECH ONLY'),
    18: ('Google Ads SEO Ranking', 'SCREENSHOT OF THE GOOGLE ADS RANKING'),
    24: ('Google Ads and Installs', None),
}
for n, (title, sub) in BAND.items():
    s = slides[n - 1]
    tree = s.shapes._spTree
    headline = None
    for sh in list(s.shapes):
        top, bottom = sh.top / EMU, (sh.top + sh.height) / EMU
        is_footer_logo = abs(sh.left / EMU - 17.4) < 0.1 and top > 10.2
        if bottom <= 2.25 or is_footer_logo:
            if sh.shape_type == 6:
                txt = ' '.join(x.text_frame.text for x in sh.shapes if x.has_text_frame and x.text_frame.text.strip())
                if len(txt) > 20:
                    headline = txt.strip()
            tree.remove(sh._element)
    # move content up to where the standard header ends (the Abdulelah video is resized to fit)
    rest = [sh for sh in s.shapes]
    if rest:
        media = [sh for sh in rest if sh.shape_type == 16]
        for m in media:
            if m.top / EMU < 2.35:
                h = 10.15 * EMU - 2.35 * EMU
                m.width = int(m.width * h / m.height); m.height = int(h); m.top = int(2.35 * EMU)
        tops = [sh.top for sh in rest if sh.shape_type != 16 and sh.top / EMU < 10.2]
        if tops:
            delta = min(tops) - int(2.4 * EMU)
            if delta > 0.2 * EMU:
                for sh in rest:
                    if sh.shape_type != 16:
                        sh.top = sh.top - delta
    for key in ('ajyad', 'title_a', 'title', 'sub', 'emyaa', 'bar'):
        el = copy_into(REF, s, shape_by_id(REF, REF_IDS[key])._element)
        if key == 'title':
            set_group_text(el, title)
        if key == 'sub':
            set_group_text(el, (sub or headline or '').upper())
    renumber_ids(s)


# ---------- 1b. identical title, accent bar and subtitle on every content slide ----------
ref_title = shape_by_id(REF, REF_IDS['title'])._element
ref_bar = shape_by_id(REF, REF_IDS['bar'])._element
for s in slides:
    if s is REF:
        continue
    tree = s.shapes._spTree
    title = bar = None
    for sh in s.shapes:
        top = sh.top / EMU
        if sh.shape_type == 6 and top < 1.3 and sh.left / EMU < 1.2:
            if sh.width / EMU < 0.3:
                bar = sh
            elif any(x.has_text_frame and x.text_frame.text.strip() for x in sh.shapes):
                title = sh
    if title is None:
        continue                                   # title slide and section dividers
    text_ = ' '.join(x.text_frame.text.strip() for x in title.shapes if x.has_text_frame and x.text_frame.text.strip())
    new = copy_into(REF, s, ref_title)
    set_group_text(new, text_)
    title._element.addprevious(new); tree.remove(title._element)
    if bar is not None:
        nb = copy_into(REF, s, ref_bar)
        bar._element.addprevious(nb); tree.remove(bar._element)
    for sh in list(s.shapes):                      # subtitle drawn as a pill: make it the plain subtitle line
        if sh.shape_type == 6 and 1.4 < sh.top / EMU < 2.0 and sh.left / EMU < 1.5 and sh.width / EMU < 9:
            t = ' '.join(x.text_frame.text.strip() for x in sh.shapes if x.has_text_frame and x.text_frame.text.strip())
            if t and t.isupper():
                ns = copy_into(REF, s, shape_by_id(REF, REF_IDS['sub'])._element)
                set_group_text(ns, t)
                box_ = (sh.left, sh.top, sh.width)
                for other in list(s.shapes):   # the pill's background layers at the same place
                    if other.shape_type == 6 and abs(other.top - box_[1]) < 0.3 * EMU and abs(other.left - box_[0]) < 0.4 * EMU \
                            and abs(other.width - box_[2]) < 0.6 * EMU:
                        tree.remove(other._element)
    for sh in s.shapes:                            # subtitle line: 13 pt everywhere
        if sh.has_text_frame and 1.5 < sh.top / EMU < 2.2 and sh.text_frame.text.strip() and sh.text_frame.text.isupper():
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = shape_by_id(REF, REF_IDS['sub']).text_frame.paragraphs[0].runs[0].font.size
    renumber_ids(s)


# ---------- 1c. navy band header on every content slide ----------
SECTION = {2: 'PERFORMANCE', 3: 'PERFORMANCE', 4: 'PERFORMANCE',
           6: 'IT UPDATE', 7: 'IT UPDATE', 8: 'IT UPDATE',
           10: 'CLIENT EXPERIENCE', 11: 'CLIENT EXPERIENCE', 12: 'CLIENT EXPERIENCE',
           14: 'MARKETING', 15: 'MARKETING', 16: 'MARKETING', 17: 'MARKETING', 18: 'MARKETING',
           21: 'APPENDIX  |  IT', 22: 'APPENDIX  |  IT', 23: 'APPENDIX  |  IT', 24: 'APPENDIX  |  MARKETING',
           25: 'APPENDIX  |  IT', 26: 'APPENDIX  |  IT', 27: 'APPENDIX  |  CLIENT EXPERIENCE'}
for n, s in enumerate(slides, 1):
    if n not in SECTION:
        continue
    tree = s.shapes._spTree
    title_txt = None
    for sh in list(s.shapes):
        top, left = sh.top / EMU, sh.left / EMU
        if sh.shape_type == 6 and top < 1.3 and left < 1.2:          # title, its companion layer, accent bar
            t = ' '.join(x.text_frame.text.strip() for x in sh.shapes if x.has_text_frame and x.text_frame.text.strip())
            if t:
                title_txt = t
            tree.remove(sh._element)
        elif left > 16 and top < 1.2:                                 # dark Ajyad logo
            tree.remove(sh._element)
        elif left < 1.2 and top < 1.3 and sh.width / EMU < 0.3:       # loose accent bar
            tree.remove(sh._element)
    first = tree[2] if len(tree) > 2 else None                        # band goes to the back
    for key in ('band', 'rule', 'title', 'ajyad', 'kicker'):
        el = copy_into(BAND_SRC, s, BAND_EL[key])
        if key == 'title':
            set_group_text(el, title_txt)
        if key == 'kicker':
            el.find('.//' + qn('a:t')).text = SECTION[n]
        if key in ('band', 'rule') and first is not None:
            first.addprevious(el)
    renumber_ids(s)


# ---------- 2. SVG images to PNG ----------
svg_cache = {}
def svg_png(part, px):
    key = (part.partname, px)
    if key not in svg_cache:
        png = cairosvg.svg2png(bytestring=part.blob, output_width=px)
        svg_cache[key] = png
    return svg_cache[key]


def ext_px(blip):
    """Display width of the shape that holds this blip, in pixels at ~300 dpi."""
    sp = blip
    while sp is not None and etree.QName(sp).localname not in ('sp', 'pic'):
        sp = sp.getparent()
    ext = sp.find('.//' + qn('a:ext'))
    w = int(ext.get('cx')) / EMU if ext is not None else 2
    return max(600, min(3000, int(w * 300)))


# ---------- 3. gradients to PNG ----------
def hexc(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def grad_png(grad, w_emu, h_emu):
    stops = []
    for gs in grad.iter(qn('a:gs')):
        c = gs.find(qn('a:srgbClr'))
        stops.append((int(gs.get('pos')) / 100000, hexc(c.get('val'))))
    stops.sort()
    W = 600
    H = max(8, min(2400, int(W * h_emu / max(w_emu, 1))))
    if H > 1200:
        W = int(W * 1200 / H); H = 1200
    lin, path = grad.find(qn('a:lin')), grad.find(qn('a:path'))
    if lin is not None:
        ang = math.radians(int(lin.get('ang', '0')) / 60000)
        dx, dy = math.cos(ang), math.sin(ang)
        span = abs(dx) * W + abs(dy) * H
        x0 = 0 if dx >= 0 else W; y0 = 0 if dy >= 0 else H
        f = lambda x, y: ((x - x0) * dx + (y - y0) * dy) / span
    else:
        ftr = path.find(qn('a:fillToRect'))
        cx = int(ftr.get('l', 50000)) / 100000 if ftr is not None else 0.5
        cy = int(ftr.get('t', 50000)) / 100000 if ftr is not None else 0.5
        far = max(math.hypot(ax - cx, ay - cy) for ax in (0, 1) for ay in (0, 1))
        f = lambda x, y: math.hypot(x / W - cx, y / H - cy) / far
    img = Image.new('RGB', (W, H))
    px = img.load()
    for y in range(H):
        for x in range(W):
            t = min(1, max(0, f(x, y)))
            for i in range(len(stops) - 1):
                (p0, c0), (p1, c1) = stops[i], stops[i + 1]
                if t <= p1 or i == len(stops) - 2:
                    u = 0 if p1 == p0 else min(1, max(0, (t - p0) / (p1 - p0)))
                    px[x, y] = tuple(int(c0[k] + (c1[k] - c0[k]) * u) for k in range(3))
                    break
    buf = io.BytesIO(); img.save(buf, 'PNG'); buf.seek(0)
    return buf


counts = {'svg': 0, 'grad_shape': 0, 'grad_text': 0}
# the deck's card gradient (radial, light blue corner to navy): used for every card so they all match
STD = next(g for g in slides[1]._element.iter(qn('a:gradFill'))
           if etree.QName(g.getparent()).localname == 'spPr' and g.find(qn('a:path')) is not None)

for s in slides:
    part = s.part
    # SVG blips
    for blip in list(s._element.iter(qn('a:blip'))):
        svgb = blip.find('.//{%s}svgBlip' % SVG_NS)
        if svgb is None:
            continue
        svg_part = part.related_part(svgb.get(qn('r:embed')))
        png = svg_png(svg_part, ext_px(blip))
        _, rid = part.get_or_add_image_part(io.BytesIO(png))
        blip.set(qn('r:embed'), rid)
        ext_lst = blip.find(qn('a:extLst'))
        for ext in list(ext_lst):
            if ext.find('{%s}svgBlip' % SVG_NS) is not None:
                ext_lst.remove(ext)
        if len(ext_lst) == 0:
            blip.remove(ext_lst)
        counts['svg'] += 1
    # gradients
    for grad in list(s._element.iter(qn('a:gradFill'))):
        par = grad.getparent()
        if etree.QName(par).localname == 'rPr':
            sf = etree.Element(qn('a:solidFill'))
            etree.SubElement(sf, qn('a:srgbClr'), val=TITLE_BLUE)
            par.replace(grad, sf)
            counts['grad_text'] += 1
        elif etree.QName(par).localname == 'spPr':
            ext = par.find(qn('a:xfrm') + '/' + qn('a:ext'))
            if b'2B4AA0' in etree.tostring(grad):      # Client Experience cards: match the other cards
                src = STD
            else:
                src = grad
            _, rid = part.get_or_add_image_part(grad_png(src, int(ext.get('cx')), int(ext.get('cy'))))
            bf = etree.Element(qn('a:blipFill'))
            etree.SubElement(bf, qn('a:blip')).set(qn('r:embed'), rid)
            st = etree.SubElement(bf, qn('a:stretch')); etree.SubElement(st, qn('a:fillRect'))
            par.replace(grad, bf)
            counts['grad_shape'] += 1

# ---------- 4. slide 2 figures ----------
# Revenue to date is $140.21 (management, 30 Sep); previous update $107.52, so +30.4%.
# Every increase uses the green arrow (deposits and AUM were grey).
def walk(g):
    for sh in g.shapes:
        if sh.shape_type == 6:
            yield from walk(sh)
        else:
            yield sh
GREEN_UP = '3DDC84'
for sh in walk(slides[1]):
    if not sh.has_text_frame:
        continue
    t = sh.text_frame.text.strip()
    if t.startswith('\u25b2'):
        for r in sh.text_frame.paragraphs[0].runs:
            r._r.find('.//' + qn('a:srgbClr')).set('val', GREEN_UP)
    elif t == '$107.52':
        sh.text_frame.paragraphs[0].runs[0].text = '$140.21'
    elif t.startswith('From 14 May 2026'):
        p0 = sh.text_frame.paragraphs[0]
        r0 = p0.runs[0]
        r0.text = 'prev  '
        for txt, col in (('$107.52', 'FFFFFF'), ('     \u25b2 +30.4%', GREEN_UP)):
            nr = copy.deepcopy(r0._r)
            nr.find(qn('a:t')).text = txt
            nr.find('.//' + qn('a:srgbClr')).set('val', col)
            p0._p.append(nr)

prs.save(OUT)
print('saved', OUT, counts)
