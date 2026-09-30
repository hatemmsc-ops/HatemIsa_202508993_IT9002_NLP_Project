"""Extract every report page of the eMYAA Daily Performance Dashboard PDFs into JSON (newest first)."""
import json, re, pymupdf
SOURCES = ['input/analytics/dashboards/eMYAA_-_Daily_Performance_Dashboard_34.pdf',
           'input/analytics/dashboards/eMYAA_-_Daily_Performance_Dashboard-1.pdf']
out = []
for src in SOURCES:
    for pno, page in enumerate(pymupdf.open(src)):
        spans = [s for b in page.get_text('dict')['blocks'] for l in b.get('lines', []) for s in l['spans'] if s['text'].strip()]
        labels = [s for s in spans if abs(s['size'] - 14.2) < 0.3]
        values = [s for s in spans if abs(s['size'] - 39.0) < 0.5 and 'Type3' not in s['font']]
        subs = [s for s in spans if abs(s['size'] - 11.2) < 0.3]
        cards = {}
        for lab in labels:
            lx, ly = lab['bbox'][0], lab['bbox'][1]
            v = [s for s in values if abs(s['bbox'][0] - lx) < 120 and 0 < ly - s['bbox'][1] < 90]
            v = min(v, key=lambda s: ly - s['bbox'][1])
            sb = [s for s in subs if abs(s['bbox'][0] - lx) < 30 and 0 < s['bbox'][1] - ly < 50]
            cards[lab['text'].strip()] = {'value': v['text'].strip(), 'sub': sb[0]['text'].strip() if sb else ''}
        t = page.get_text()
        date = re.search(r'Report Date:\s*([^\n]+)', t).group(1).strip()
        period = re.search(r'Reporting Period:\s*([^\n]+)', t).group(1).strip()
        period = re.sub(r'\s+', ' ', period)
        out.append({'source': src.split('/')[-1], 'page': pno + 1, 'date': date, 'period': period, 'cards': cards})
json.dump(out, open('working/audit-data/dashboard-pages.json', 'w'), indent=1)
print(len(out), out[0]['date'], out[-1]['date'])
print(json.dumps(out[0]['cards'], indent=0)[:1500])
bad = [(o['date'], len(o['cards'])) for o in out if len(o['cards']) != 15]
print('pages without 15 cards:', bad)
