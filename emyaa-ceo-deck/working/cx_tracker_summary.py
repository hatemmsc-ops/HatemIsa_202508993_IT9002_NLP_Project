"""Aggregate the Client Monthly Tracker into counts only (no names, phones or emails).

The tracker holds client personal data, so it is kept out of git (see .gitignore);
only this summary is committed. Run from emyaa-ceo-deck/:  python3 working/cx_tracker_summary.py
"""
import collections, csv, openpyxl

SRC = 'input/source/client-experience/eMYAA_Client_Monthly_Tracker_-_2026.xlsx'
OUT = 'working/audit-data/cx-tracker-weekly.csv'

wb = openpyxl.load_workbook(SRC, data_only=True)
rows = []
for ws in wb.worksheets[1:]:
    allr = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(allr) if r[1] == 'Customer')
    hdr = [h.strip() if isinstance(h, str) else h for h in allr[hi]]
    wk = next(r for r in allr if r[1] == 'Week Starting')
    data = [dict(zip(hdr, r)) for r in allr[hi + 1:] if any(r[i] for i in (1, 2, 3))]
    data = [d for d in data if d.get('User Segment')]          # rows with no client type are not counted
    seg = collections.Counter(d['User Segment'] for d in data)
    stage = collections.Counter(d.get('Current Stage') or 'Not logged' for d in data)
    status = collections.Counter(d.get('Action Status') or 'Not logged' for d in data)
    reason = collections.Counter(d.get('Blocker / Reason') for d in data if d.get('Blocker / Reason'))
    rows.append({
        'week': ws.title, 'week_ending': wk[7].date() if hasattr(wk[7], 'date') else wk[7],
        'clients': len(data),
        'potential_not_onboarded': seg.get('Potential Customer', 0),
        'funding_gap': seg.get('Funded Gap Customer', 0),
        'reached': stage.get('Contacted', 0), 'no_answer': stage.get('Contact Attempted', 0),
        'stage_other': sum(v for k, v in stage.items() if k not in ('Contacted', 'Contact Attempted', 'Not logged')),
        'stage_not_logged': stage.get('Not logged', 0),
        'completed': status.get('Completed', 0), 'closed': status.get('Closed', 0),
        'waiting': status.get('Waiting for Client', 0), 'status_not_logged': status.get('Not logged', 0),
        'closed_spam': reason.get('Spam', 0), 'closed_not_interested': reason.get('Not Interested', 0),
        'with_owner_or_due_date': sum(1 for d in data if d.get('Next Follow-Up')),
        'with_funding_value': sum(1 for d in data if d.get('Funding Value')),
    })
with open(OUT, 'w', newline='') as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0]))
    w.writeheader(); w.writerows(rows)
for r in rows:
    print(r)
