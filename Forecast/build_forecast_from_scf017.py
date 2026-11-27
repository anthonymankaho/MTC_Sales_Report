# -*- coding: utf-8 -*-
"""Rebuild the 98-col forecast data sheet FROM the current-time SCF017 (Nov-27),
the authoritative complete SO/PO outstanding report, so forecast == SC17.
Maps SCF017 fields -> 98-col layout; derived lookups from Sku/Part/BRAND/Customer/Vendor.
Writes MTC_FORECAST_20261127_UPDATED.xlsx (data sheet) + prints comparison vs SCF017 & old forecast.
"""
import zipfile, re, sys, json, datetime, os, glob
from xml.etree import ElementTree as ET
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')
M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'; RR='{http://schemas.openxmlformats.org/package/2006/relationships}'
FC=os.path.dirname(os.path.abspath(__file__)).replace('\\','/')
# args: [1]=SCF017 path (required), [2]=output xlsx (optional)
if len(sys.argv)<2:
    sys.exit('usage: build_forecast_from_scf017.py <SCF017.xlsx> [output.xlsx]')
SCF017_PATH=sys.argv[1]
def cidx(ref):
    L=re.match(r'([A-Z]+)\d+',ref).group(1); n=0
    for ch in L: n=n*26+(ord(ch)-64)
    return n-1
def colletter(n):
    s=''; n+=1
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
def openwb(p):
    z=zipfile.ZipFile(p); wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rm={x.get('Id'):x.get('Target') for x in rels.findall(RR+'Relationship')}
    sh={}
    for s in wb.iter(M+'sheet'):
        t=rm.get(s.get(RID))
        if t and not t.startswith('xl/'): t='xl/'+t.lstrip('/')
        sh[s.get('name')]=t
    ss=[]
    try:
        sst=ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst.findall(M+'si'): ss.append(''.join(t.text or '' for t in si.iter(M+'t')))
    except KeyError: pass
    return z,sh,ss
def cv(c,ss):
    t=c.get('t'); v=c.find(M+'v'); ise=c.find(M+'is')
    if ise is not None: return ''.join(x.text or '' for x in ise.iter(M+'t'))
    if v is not None: return ss[int(v.text)] if t=='s' else v.text
    return None
def rows(z,tgt,ss):
    r=ET.fromstring(z.read(tgt)); out=[]
    for row in r.iter(M+'row'):
        cells={}
        for c in row.findall(M+'c'):
            ref=c.get('r')
            if ref: cells[cidx(ref)]=cv(c,ss)
        if cells: out.append(cells)
    return out
def num(v):
    try: return float(v)
    except: return 0.0
def ser2date(v):
    try: return (datetime.date(1899,12,30)+datetime.timedelta(days=int(float(v)))).isoformat()
    except: return ''
MONTHS={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

# FX from fx_report_island (latest populated <= 2026-10)
fxj=json.load(open(FC+'/fx_report_island.json',encoding='utf-8'))
fh=fxj['header']; mi=fh.index('yyyy-mm'); best={}
for r in fxj['rows']:
    ym=str(r[mi])
    if ym>'2026-10': continue
    for j,cur in enumerate(fh):
        if j==mi: continue
        v=r[j]
        if v in (None,'',' '): continue
        if cur not in best or ym>best[cur][0]: best[cur]=(ym,float(v))
FX={c:v for c,(ym,v) in best.items()}
def fx(cur):
    c=str(cur or 'USD').strip().upper() or 'USD'
    return FX.get(c,1.0)

# lookups from UPDATED workbook Sku/Part/BRAND/Customer/Vendor
# (skip data-only rebuild outputs, which have no Sku sheet; skip our own output)
_outpath=os.path.abspath(sys.argv[2]) if len(sys.argv)>2 else os.path.abspath(FC+'/MTC_FORECAST_%s_UPDATED.xlsx'%datetime.date.today().strftime('%Y%m%d'))
_lu=sorted(glob.glob(FC+'/MTC_FORECAST_*_UPDATED.xlsx'), key=os.path.getmtime, reverse=True)
_pick=None
for _p in _lu:
    if os.path.abspath(_p)==_outpath: continue
    _z,_sh,_ss=openwb(_p)
    if 'Sku' in _sh and 'Part' in _sh and 'BRAND' in _sh:
        _pick=(_z,_sh,_ss); break
    _z.close()
if not _pick: sys.exit('no MTC_FORECAST_*_UPDATED.xlsx with Sku/Part/BRAND lookup sheets found')
zu,shu,ssu=_pick
sku={}
for r in rows(zu,shu['Sku'],ssu)[1:]:
    k=str(r.get(0) or '').strip().upper()
    if k and k not in sku: sku[k]=[r.get(i) for i in range(5)]
part={}
for r in rows(zu,shu['Part'],ssu):
    k=str(r.get(0) or '').strip().upper()
    if k and k not in part: part[k]=[r.get(i) for i in range(5)]
brand={}
for r in rows(zu,shu['BRAND'],ssu)[1:]:
    if r.get(0): brand[str(r.get(0)).strip().upper()]=r.get(1)
cust={}
for r in rows(zu,shu['Customer Details'],ssu)[2:]:
    if r.get(0): cust[str(r.get(0)).strip().upper()]=r
vend={}
for r in rows(zu,shu['Vendor'],ssu)[1:]:
    if r.get(0): vend[str(r.get(0)).strip().upper()]=r
zu.close()

# SCF017 current
zs,shs,sss=openwb(SCF017_PATH)
s=rows(zs,shs['Page1'],sss)
h={str(v).strip():i for i,v in s[0].items() if v}
zs.close()

HEADERS=['Key','Invoice','Date','Max_inv','Ship_Code','Ship_to','Division Code','Division Desc','Factory','Division','Cust_id','Cust_name','Region','Cust_Countries','Terms','Open a/c ; L/C ; T/T','FOB','Related Party','Subsidiary','Family','Family1','Part_id','Part_desc1','Part_desc2','Webspec_Desc','Part_desc3','Comm_code','Part No = ?????-?','SKU_qty','no of pans','Ttl  Pans','Web UOM','Ttl Pans (UOM)','Unit_Price_Inv','Net_Price_Inv','curr','rate','Unit_Price_USD','Net_Unit_Price_USD','Sales_Inv','Ttl_Commission1_Inv','Ttl_Commission2_Inv','Ttl_Direct Commission_Inv','Ttl_Defective_Allowance_Inv','Ttl_Return_Rebate_Inv','Ttl_Disc_Rebate_Inv','Ttl_Royalty_Inv','Net_Sales_Inv','Sales_USD','Net_Sales_USD','Matl_USD','Labor_USD','Matl_brdn_USD','Labor_brdn_USD','Subcon_cost_USD','Prime_USD','Gross_Cont_USD','Gross_Cont_%','Net_Cont_USD','Net_Cont_%','Dupont Rebate (USD/unit)','Ttl Dupont Rebate (USD)','Alum Rebate (USD/unit)','Ttl Alum Rebate (USD)','Prime Cost After Rebates (USD)','Net Cont After Rebates (USD)','Net Cont After Rebates (%)','Family Category','Meyer Brand','Database','yyyy-mm','Year','Ttl PO USD','VEND_ID','VEND_NAME','VEND_NAME 1','PO_CURR','LST_RCP_PR (ORI CUR)','TTLLST_RCP_PR (ORI CUR)','SO_NO','PO_NO','PO Rate','U / PO USD','TTL PO USD','TTL PO GROSS CONT USD','TTL PO NET CONT USD','TTL PO NET CONT %','Net Net Price (Fcy)','Total Net Net Sales (USD)','Total Net Net Margin (USD)','Commodity','Month','Years','Combined Cust','Combined Vendor','Combined Brand','Vendor Country','Ship By Date']
assert len(HEADERS)==98

out=[]
for r in s[1:]:
    so=str(r.get(h['ord_nbr']) or '').strip(); ln=str(r.get(h['ord_ln_nbr']) or '').strip()
    pid=str(r.get(h['part_id']) or '').strip(); up=pid.upper()
    if not so: continue
    due=num(r.get(h['due_qty'])); osq=num(r.get(h['os_qty']))
    socur=str(r.get(h['so_curr_cd']) or 'USD'); pocur=str(r.get(h['po_curr_cd']) or 'USD')
    sup=num(r.get(h['so_uprice'])); pup=num(r.get(h['po_uprice']))
    ship=ser2date(r.get(h['so_due_dt']))
    ym=ship[:7]; yr=ship[:4]; mo=MONTHS.get(int(ship[5:7]),'') if len(ship)>=7 else ''
    sfr=fx(socur); pfr=fx(pocur)
    sales_inv=due*sup; sales_usd=sales_inv*sfr
    po_doc=osq*pup; po_usd=po_doc*pfr
    is_sku=len(pid)==7
    sk=sku.get(up); pt=part.get(up)
    famcat=(sk[2] if is_sku and sk and sk[2] else (pt[1] if pt and pt[1] else 'MISC'))
    family=(sk[3] if is_sku and sk and sk[3] else (pt[2] if pt and pt[2] else 'NO BRAND'))
    webspec=(sk[1] if is_sku and sk and sk[1] else (pt[4] if pt and pt[4] else ''))
    comm='SKU-' if is_sku else (pt[3] if pt and pt[3] else 'MISC')
    uom=(sk[4] if is_sku and sk and sk[4] else 0)
    meyer=brand.get(str(famcat).strip().upper(),'')
    cid=str(r.get(h['cust_id']) or '').strip().upper(); cr=cust.get(cid)
    vid=str(r.get(h['vend_id']) or '').strip().upper(); vr=vend.get(vid)
    R=[None]*98
    R[0]='MTC-%s-%s-%s'%(so,ln,pid)
    R[2]=ship; R[6]='65'; R[7]='MTC-SOURCING'; R[9]='SOURCING'
    R[10]=str(r.get(h['cust_id']) or ''); R[11]=str(r.get(h['cust_name']) or '')
    R[12]=cr[3] if cr and len(cr)>3 else ''
    R[13]=cr[2] if cr and len(cr)>2 else ''
    R[19]=family; R[20]=family; R[21]=pid
    R[22]=str(r.get(h['part_desc']) or ''); R[23]=R[22]; R[24]=webspec
    R[26]=comm; R[28]=due; R[31]=uom
    R[33]=sup; R[35]=socur; R[36]=sfr; R[37]=sup*sfr; R[39]=sales_inv; R[48]=sales_usd
    R[50]=po_usd; R[55]=po_usd; R[56]=sales_usd-po_usd
    R[57]=(sales_usd-po_usd)/sales_usd if sales_usd else None
    R[67]=famcat; R[68]=meyer; R[69]='Forecast'
    R[70]=ym; R[71]=yr; R[72]=po_usd
    R[73]=str(r.get(h['vend_id']) or ''); R[74]=str(r.get(h['vend_name']) or ''); R[75]=R[74]
    R[76]=pocur; R[77]=pup; R[78]=po_doc
    R[79]=so; R[80]=str(r.get(h['po_no']) or ''); R[81]=pfr; R[82]=pup*pfr; R[83]=po_usd
    R[84]=sales_usd-po_usd; R[90]='SKU-' if (len(pid)>=6 and pid[5]=='-') else comm
    R[91]=mo; R[92]=yr
    R[93]=str(r.get(h['cust_name']) or ''); R[94]=str(r.get(h['vend_name']) or '')
    R[95]=famcat; R[96]=vr[3] if vr and len(vr)>3 else ''
    R[97]=ship
    out.append(R)

# write xlsx (single data sheet)
def esc(v): return str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')
def cell(ref,v):
    if v is None or v=='': return '<c r="%s"/>'%ref
    if isinstance(v,(int,float)): return '<c r="%s"><v>%s</v></c>'%(ref,v)
    return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>'%(ref,esc(v))
parts=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
parts.append('<row r="1">%s</row>'%''.join(cell('%s1'%colletter(i),HEADERS[i]) for i in range(98)))
for ri,R in enumerate(out,start=2):
    parts.append('<row r="%d">%s</row>'%(ri,''.join(cell('%s%d'%(colletter(i),ri),R[i]) for i in range(98))))
parts.append('</sheetData></worksheet>')
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
rr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
wbx='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data" sheetId="1" r:id="rId1"/></sheets></workbook>'
wbr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
OUT=sys.argv[2] if len(sys.argv)>2 else FC+'/MTC_FORECAST_%s_UPDATED.xlsx'%datetime.date.today().strftime('%Y%m%d')
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as zo:
    zo.writestr('[Content_Types].xml',ct); zo.writestr('_rels/.rels',rr)
    zo.writestr('xl/workbook.xml',wbx); zo.writestr('xl/_rels/workbook.xml.rels',wbr)
    zo.writestr('xl/worksheets/sheet1.xml',''.join(parts))
print('WROTE', OUT, 'rows=', len(out))
tot_sales=sum(R[48] or 0 for R in out); tot_prime=sum(R[55] or 0 for R in out); tot_gross=sum(R[56] or 0 for R in out)
tot_qty=sum(R[28] or 0 for R in out)
print('NEW forecast(from SCF017 Nov-27): rows=%d qty=%.0f salesUSD=%.2f primeUSD=%.2f grossUSD=%.2f' % (len(out),tot_qty,tot_sales,tot_prime,tot_gross))
# compare vs SCF017 due qty
sc_due=sum(num(r.get(h['due_qty'])) for r in s[1:])
print('SCF017 due qty=%.0f | new forecast qty=%.0f | match=%s' % (sc_due, tot_qty, abs(sc_due-tot_qty)<0.01))
