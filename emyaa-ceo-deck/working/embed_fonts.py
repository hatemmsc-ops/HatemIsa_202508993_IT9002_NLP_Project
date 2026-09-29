"""Embed the Board deck's Exo 2 font files (EOT .fntdata, as PowerPoint stores them) into the built deck."""
import re, shutil, sys, zipfile, os
DECK = sys.argv[1] if len(sys.argv) > 1 else 'output/final-deck/eMYAA-Marketing-Update-01.10.2026-CEO-Redesigned.pptx'
FONTS = {'regular': 'Exo2.fntdata', 'bold': 'Exo2Bold.fntdata', 'italic': 'Exo2Italics.fntdata', 'boldItalic': 'Exo2BoldItalics.fntdata'}
SRC = 'input/brand-assets/fonts/'
tmp = DECK + '.tmp'
zin = zipfile.ZipFile(DECK)
names = zin.namelist()
rels = zin.read('ppt/_rels/presentation.xml.rels').decode()
pres = zin.read('ppt/presentation.xml').decode()
ct = zin.read('[Content_Types].xml').decode()
assert 'embeddedFontLst' not in pres, 'already embedded'
ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
nxt = max(ids) + 1
entries, rel_add, files = [], '', {}
for i, (kind, fn) in enumerate(FONTS.items()):
    rid = f'rId{nxt + i}'
    part = f'fonts/emyaa-font{i + 1}.fntdata'
    rel_add += f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="{part}"/>'
    entries.append(f'<p:{kind} r:id="{rid}"/>')
    files['ppt/' + part] = open(SRC + fn, 'rb').read()
rels = rels.replace('</Relationships>', rel_add + '</Relationships>')
lst = '<p:embeddedFontLst><p:embeddedFont><p:font typeface="Exo 2" charset="0"/>' + ''.join(entries) + '</p:embeddedFont></p:embeddedFontLst>'
# CT_Presentation order: ... sldSz, notesSz, smartTags?, embeddedFontLst?, custShowLst?, photoAlbum?, custDataLst?, kinsoku?, defaultTextStyle? ...
m = re.search(r'<p:notesSz[^>]*/>', pres)
assert m, 'notesSz not found'
pres = pres[:m.end()] + lst + pres[m.end():]
pres = re.sub(r'<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', pres, count=1) if 'embedTrueTypeFonts' not in pres else pres
if 'Extension="fntdata"' not in ct:
    ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
for n in names:
    data = zin.read(n)
    if n == 'ppt/_rels/presentation.xml.rels': data = rels.encode()
    elif n == 'ppt/presentation.xml': data = pres.encode()
    elif n == '[Content_Types].xml': data = ct.encode()
    zout.writestr(zin.getinfo(n), data)
for n, d in files.items(): zout.writestr(n, d)
zout.close(); zin.close(); os.replace(tmp, DECK)
print('embedded Exo 2 (4 styles) into', DECK)
