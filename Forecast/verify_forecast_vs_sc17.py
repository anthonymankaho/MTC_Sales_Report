# -*- coding: utf-8 -*-
"""VERIFY step (monthly SOP step 1): compare the built SO/PO combine forecast
against an SC17 / SCF017 outstanding report and report ANY quantity difference.

Usage:
  python verify_forecast_vs_sc17.py <forecast.xlsx> <sc17.xlsx>
  (defaults: newest MTC_FORECAST_*_UPDATED.xlsx in this folder; SC17 path required)

Handles both SC17 header styles:
  SC17_YYYYMMDD : ORD_NBR / PART_ID / IVR_QTY
  EPI_SCF017_*  : ord_nbr / part_id / due_qty
Aggregates by (SO, PART) and compares outstanding qty; lists matched-qty diffs,
forecast-only and SC17-only keys, and grand totals. Read-only.
"""
import zipfile, re, sys, os, glob
from xml.etree import ElementTree as ET
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
HERE = os.path.dirname(os.path.abspath(__file__))

def cidx(ref):
    L = re.match(r'([A-Z]+)\d+', ref).group(1); n = 0
    for ch in L: n = n*26 + (ord(ch)-64)
    return n-1

def read_rows(path):
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    RR = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rm = {x.get('Id'): x.get('Target') for x in rels.findall(RR+'Relationship')}
    tgt = None
    for s in wb.iter(M+'sheet'):
        t = rm.get(s.get(RID))
        if t and not t.startswith('xl/'): t = 'xl/'+t.lstrip('/')
        tgt = t; break  # first sheet
    ss = []
    try:
        sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst.findall(M+'si'):
            ss.append(''.join(t.text or '' for t in si.iter(M+'t')))
    except KeyError:
        pass
    r = ET.fromstring(z.read(tgt)); out = []
    for row in r.iter(M+'row'):
        cells = {}
        for c in row.findall(M+'c'):
            ref = c.get('r')
            if not ref: continue
            t = c.get('t'); v = c.find(M+'v'); ise = c.find(M+'is')
            if ise is not None: val = ''.join(x.text or '' for x in ise.iter(M+'t'))
            elif v is not None: val = ss[int(v.text)] if t == 's' else v.text
            else: val = None
            cells[cidx(ref)] = val
        if cells: out.append(cells)
    z.close(); return out

def num(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0

def header_map(row):
    return {str(v).strip().lower(): i for i, v in row.items() if v}

def pick(h, *names):
    for n in names:
        if n.lower() in h: return h[n.lower()]
    return None

def agg(rows, so_i, part_i, qty_i):
    d = defaultdict(float)
    for r in rows[1:]:
        so = str(r.get(so_i) or '').strip(); part = str(r.get(part_i) or '').strip().upper()
        if so: d[(so, part)] += num(r.get(qty_i))
    return d

def main():
    fc_path = sys.argv[1] if len(sys.argv) > 1 else None
    sc_path = sys.argv[2] if len(sys.argv) > 2 else None
    if not fc_path:
        cands = sorted(glob.glob(os.path.join(HERE, 'MTC_FORECAST_*_UPDATED.xlsx')), key=os.path.getmtime)
        if not cands: sys.exit('no MTC_FORECAST_*_UPDATED.xlsx found')
        fc_path = cands[-1]
    if not sc_path: sys.exit('usage: verify_forecast_vs_sc17.py [forecast.xlsx] <sc17.xlsx>')
    fr = read_rows(fc_path); fh = header_map(fr[0])
    fso = pick(fh, 'SO_NO'); fpart = pick(fh, 'Part_id'); fqty = pick(fh, 'SKU_qty')
    sr = read_rows(sc_path); sh = header_map(sr[0])
    sso = pick(sh, 'ORD_NBR', 'ord_nbr'); spart = pick(sh, 'PART_ID', 'part_id'); sqty = pick(sh, 'IVR_QTY', 'due_qty')
    if None in (fso, fpart, fqty, sso, spart, sqty): sys.exit('could not locate SO/PART/QTY columns')
    fc = agg(fr, fso, fpart, fqty); sc = agg(sr, sso, spart, sqty)
    both = set(fc) & set(sc); fonly = set(fc) - set(sc); sonly = set(sc) - set(fc)
    qd = [(k, sc[k], fc[k]) for k in sorted(both) if abs(sc[k]-fc[k]) > 0.01]
    print('FORECAST: %s  keys=%d qty=%.0f' % (os.path.basename(fc_path), len(fc), sum(fc.values())))
    print('SC17    : %s  keys=%d qty=%.0f' % (os.path.basename(sc_path), len(sc), sum(sc.values())))
    print('MATCHED=%d  qty-diff=%d  forecast-only=%d  SC17-only=%d' % (len(both), len(qd), len(fonly), len(sonly)))
    for k, a, b in qd[:40]:
        print('  QTY-DIFF SO %s PART %s: SC17=%.0f FC=%.0f diff=%+.0f' % (k[0], k[1], a, b, b-a))
    for k in sorted(sonly, key=lambda k: -sc[k])[:40]:
        print('  SC17-ONLY SO %s PART %s qty=%.0f (missing from forecast)' % (k[0], k[1], sc[k]))
    for k in sorted(fonly, key=lambda k: -fc[k])[:40]:
        print('  FC-ONLY   SO %s PART %s qty=%.0f (not in SC17)' % (k[0], k[1], fc[k]))
    ok = (not qd and not fonly and not sonly)
    print('RESULT: %s' % ('IDENTICAL - no quantity difference' if ok else 'DIFFERENCES FOUND (see above)'))
    return 0 if ok else 1

if __name__ == '__main__':
    sys.exit(main())
