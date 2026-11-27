# -*- coding: utf-8 -*-
"""
Build ONE "data" sheet (98 columns = the Demo template's 97 + 'Ship By Date')
from the raw Epicor cache (_work_mtc_details_raw.json), joining each Sales Order
1:1 to its Purchase Order via the authoritative link SO.PonumE27 == PO.PonumJ16.

Database = 'Forecast' (outstanding / future on-hand orders).
Rule: OUTSTANDING ONLY - COMPLETED and CANCELLED SO/PO are excluded (status filter).

FX: fx_report_island.json = the Exchange Rate island extracted verbatim from
FIN_MTC_Performance_Report_html_TEMP.html (verified identical to the _work copy).
Rate = USD per 1 unit of currency; each side (SO / PO) converts with its OWN
currency's latest populated month <= FX_CUTOFF (2026-10, last month of the report
period). Oct-Dec rows carry only USD/HKD, so EUR etc. carry forward from 2026-09.
Unknown currency = abort (no silent 1.0 fallback).

Period: yyyy-mm / Year / Month / Years derive from SHIP BY DATE (SO line NeedByDate;
PO-only rows use the PO line DueDate), NOT the order date - this is a forward-looking
forecast. 'Date' (col 2) keeps the order date as the transaction date.
Line pairing: FIFO per SKU so duplicate-SKU SO/PO lines pair 1:1 (no orphan rows).
Quantity basis: OUTSTANDING ONLY - both sales and purchase amounts use Epidueqty
(unshipped qty), never the full order qty, so the shipped portion (which lives in the
actuals) is never shown or double-counted. Per-unit price cols are unaffected.

Lookups (from the Demo template): FAMILY, Commodity, Combined, Vendor Country, DESNEY.
Output: MTC_Data_20261120_R7.xlsx
"""
import zipfile, re, sys, json, os
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

ROOT = r"C:/Users/User/Desktop/MTC Performance Report Claude/Forecast"
DEMO = r"C:/Users/User/Desktop/Meyer/MTC/3. MTC Sales Forecast for Peggy/3. MTC Sales Forecast for Peggy/Sales_Cont_Forecast_Report-MTC/Sales_Cont_Forecast_Report-MTC-Program (Normal) TEMP.xlsx"
RAW = ROOT + "/_work_mtc_details_raw.json"

# ---------------- FX from the MTC Performance Report (Exchange Rate tab) ----------------
# fx_report_island.json was extracted verbatim from the report HTML's exrate island
# (SHEETS -> key 'exrate' -> var HEADER/ROWS) and verified identical to
# HTML_Report_Generator/_work/exrate_usd.json. Convention: USD per 1 unit of currency
# (USD=1, HKD=0.1282=1/7.8). Latest populated month <= cutoff wins per currency;
# 2026-10..12 rows only carry USD/HKD, so EUR etc. carry forward from 2026-09.
FX_CUTOFF = '2026-10'
_fx_island = json.load(open(ROOT + "/fx_report_island.json", encoding='utf-8'))
_fx_header = _fx_island['header']
_fx_month_col = _fx_header.index('yyyy-mm')
_fx_best = {}
for _r in _fx_island['rows']:
    _ym = str(_r[_fx_month_col])
    if _ym > FX_CUTOFF:
        continue
    for _j, _cur in enumerate(_fx_header):
        if _j == _fx_month_col:
            continue
        _v = _r[_j]
        if _v in (None, '', ' '):
            continue
        if _cur not in _fx_best or _ym > _fx_best[_cur][0]:
            _fx_best[_cur] = (_ym, float(_v))
FX = {c: v for c, (ym, v) in _fx_best.items()}
print("FX (USD per unit, latest <= %s): %s" % (
    FX_CUTOFF, {c: '%s (%s)' % (FX[c], _fx_best[c][0]) for c in sorted(FX)}))


def fx_rate(currency):
    cur = s(currency).upper() or 'USD'
    if cur not in FX:
        sys.exit('FX ABORT: no rate for currency %r in report island (<= %s)' % (cur, FX_CUTOFF))
    return FX[cur]


def col_idx(ref):
    m = re.match(r'([A-Z]+)\d+', ref); L = m.group(1); n = 0
    for ch in L:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_shared(z):
    shared = []
    try:
        sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst.findall(M + 'si'):
            shared.append(''.join(t.text or '' for t in si.iter(M + 't')))
    except KeyError:
        pass
    return shared


def read_ref(z, xml_bytes, shared):
    root = ET.fromstring(xml_bytes); rows = []
    for row in root.iter(M + 'row'):
        cells = {}
        for c in row.findall(M + 'c'):
            ref = c.get('r'); ci = col_idx(ref) if ref else None
            if ci is None:
                continue
            t = c.get('t'); v = c.find(M + 'v'); ise = c.find(M + 'is')
            if v is not None:
                val = shared[int(v.text)] if t == 's' else v.text
            elif ise is not None:
                val = ''.join(x.text or '' for x in ise.iter(M + 't'))
            else:
                val = None
            cells[ci] = val
        if cells:
            n = max(cells) + 1
            rows.append([cells.get(i) for i in range(n)])
    return rows


def num(v):
    if v is None or v == '':
        return None
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return v


def s(v):
    return '' if v is None else str(v).strip()


def date_only(v):
    return s(v)[:10] if s(v) else ''


MONTHS = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
          7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}


def period_from_date(d):
    d = date_only(d)
    yyyymm = d[:7] if len(d) >= 7 else ''
    year = d[:4] if len(d) >= 4 else ''
    month = ''
    try:
        month = MONTHS.get(int(d[5:7]), '')
    except (ValueError, IndexError):
        pass
    return yyyymm, year, month, year


# ---------------- load raw cache ----------------
raw = json.load(open(RAW, encoding='utf-8'))
_all_so = list(raw['so'].values())
_all_po = list(raw['po'].values())

# ---- OUTSTANDING-ONLY RULE --------------------------------------------------
# The report is "outstanding" (in-system, not yet shipped). COMPLETED orders are
# fully shipped and CANCELLED orders are voided — both out of scope, so drop them.
# Filter on STATUS, never on quantity: cancelled lines still carry a positive
# outstanding qty and would wrongly survive a qty>0 filter.
NON_OUTSTANDING = {'COMPLETED', 'CANCELLED'}
so_entities = [e for e in _all_so if str(e.get('StatusE27') or '').strip().upper() not in NON_OUTSTANDING]
po_entities = [e for e in _all_po if str(e.get('StatusJ16') or '').strip().upper() not in NON_OUTSTANDING]
print("SO entities: %d outstanding (excluded %d COMPLETED/CANCELLED)" % (len(so_entities), len(_all_so) - len(so_entities)))
print("PO entities: %d outstanding (excluded %d COMPLETED/CANCELLED)" % (len(po_entities), len(_all_po) - len(po_entities)))

# ---------------- load Demo lookups ----------------
z = zipfile.ZipFile(DEMO); shared = read_shared(z)
wb = ET.fromstring(z.read('xl/workbook.xml'))
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
RR = '{http://schemas.openxmlformats.org/package/2006/relationships}'
relmap = {r.get('Id'): r.get('Target') for r in rels.findall(RR + 'Relationship')}
sheetmap = {}
for s_ in wb.iter(M + 'sheet'):
    sheetmap[s_.get('name')] = relmap.get(s_.get(R))


def sheet_rows(name):
    tgt = sheetmap.get(name)
    if not tgt:
        return []
    if not tgt.startswith('xl/'):
        tgt = 'xl/' + tgt.lstrip('/')
    try:
        return read_ref(z, z.read(tgt), shared)
    except KeyError:
        return []


family = {}
for r in sheet_rows("FAMILY")[1:]:
    if r and r[0]:
        family[str(r[0]).strip().upper()] = r
commodity = {}
for r in sheet_rows("Commodity")[1:]:
    if r and r[0]:
        commodity[str(r[0]).strip().upper()] = r[3] if len(r) > 3 else None
combined_cust = {}
combined_vend = {}
for r in sheet_rows("Combined")[1:]:
    if not r:
        continue
    if len(r) > 0 and r[0]:
        combined_cust[str(r[0]).strip().upper()] = r[2] if len(r) > 2 else None
    if len(r) > 4 and r[4]:
        combined_vend[str(r[4]).strip().upper()] = r[6] if len(r) > 6 else None
vendor_country = {}
for r in sheet_rows("Vendor Country")[1:]:
    if r and r[0] and len(r) > 3 and r[3]:
        vendor_country[str(r[0]).strip().upper()] = r[3]
disney = {}
for r in sheet_rows("DESNEY")[1:]:
    if r and r[0] and len(r) > 2 and r[2]:
        disney[str(r[0]).strip().upper()] = r[2]
z.close()


def fg(row, i):
    if row is None or i >= len(row):
        return ''
    return s(row[i])


def lookup_family(sku):
    return family.get(s(sku).upper())


# ---------------- 97-column header ----------------
HEADERS = ['Key','Invoice','Date','Max_inv','Ship_Code','Ship_to','Division Code','Division Desc',
 'Factory','Division','Cust_id','Cust_name','Region','Cust_Countries','Terms','Open a/c ; L/C ; T/T',
 'FOB','Related Party','Subsidiary','Family','Family1','Part_id','Part_desc1','Part_desc2','Webspec_Desc',
 'Part_desc3','Comm_code','Part No = ?????-?','SKU_qty','no of pans','Ttl  Pans','Web UOM','Ttl Pans (UOM)',
 'Unit_Price_Inv','Net_Price_Inv','curr','rate','Unit_Price_USD','Net_Unit_Price_USD','Sales_Inv',
 'Ttl_Commission1_Inv','Ttl_Commission2_Inv','Ttl_Direct Commission_Inv','Ttl_Defective_Allowance_Inv',
 'Ttl_Return_Rebate_Inv','Ttl_Disc_Rebate_Inv','Ttl_Royalty_Inv','Net_Sales_Inv','Sales_USD','Net_Sales_USD',
 'Matl_USD','Labor_USD','Matl_brdn_USD','Labor_brdn_USD','Subcon_cost_USD','Prime_USD','Gross_Cont_USD',
 'Gross_Cont_%','Net_Cont_USD','Net_Cont_%','Dupont Rebate (USD/unit)','Ttl Dupont Rebate (USD)',
 'Alum Rebate (USD/unit)','Ttl Alum Rebate (USD)','Prime Cost After Rebates (USD)','Net Cont After Rebates (USD)',
 'Net Cont After Rebates (%)','Family Category','Meyer Brand','Database','yyyy-mm','Year','Ttl PO USD',
 'VEND_ID','VEND_NAME','VEND_NAME 1','PO_CURR','LST_RCP_PR (ORI CUR)','TTLLST_RCP_PR (ORI CUR)','SO_NO','PO_NO',
 'PO Rate','U / PO USD','TTL PO USD','TTL PO GROSS CONT USD','TTL PO NET CONT USD','TTL PO NET CONT %',
 'Net Net Price (Fcy)','Total Net Net Sales (USD)','Total Net Net Margin (USD)','Commodity','Month','Years',
 'Combined Cust','Combined Vendor','Combined Brand','Vendor Country','Ship By Date']
assert len(HEADERS) == 98, len(HEADERS)


def blank_row():
    return [None] * 98


def so_header_fields(so):
    return {
        'so_no': s(so.get('OrdernumE27')),
        'cust_id': s(so.get('CustidE27')),
        'cust_name': s(so.get('Name')),
        'customer_po': s(so.get('PoNum')),          # the customer's PO (base field)
        'linked_po': s(so.get('PonumE27')),          # the linked Meyer PO number
        'order_date': date_only(so.get('OrderDate')),
        'ship_code': s(so.get('ShptonumE27')),
        'terms': s(so.get('TermsCode')),
        'fob': s(so.get('Fob')),
        'currency': s(so.get('CurrencyCode')),
        'div': s(so.get('CodvcdE27')),
        'factory': s(so.get('FactcdE27')),
        'order_type': s(so.get('OrdtyE27')),
        'poe': s(so.get('PoeE27')),
        'coo': s(so.get('CooE27')),
        'vendor_id': s(so.get('VendoridE27')),
    }


def po_header_fields(po):
    return {
        'po_no': s(po.get('PonumJ16')),
        'vendor_id': s(po.get('VendoridJ16')),
        'vendor_name': s(po.get('Name')),
        'so_no': s(po.get('OrdernumJ16')),
        'order_date': date_only(po.get('OrderDate')),
        'terms': s(po.get('TermsCode')),
        'fob': s(po.get('FobJ16') or po.get('Fob')),
        'currency': s(po.get('CurrencyCode')),
        'div': s(po.get('CodvcdJ16')),
        'factory': s(po.get('FactcdJ16')),
    }


def build_row(so_hdr, so_ln, po_hdr, po_ln):
    R = blank_row()
    sku = s(so_ln.get('CurpartnumE28')) if so_ln else s(po_ln.get('CurpartnumJ10'))
    fam = lookup_family(sku)

    # ----- sales side -----
    if so_hdr is not None:
        ship_by = date_only(so_ln.get('CurneedbydateE28'))   # Ship By Date drives the forecast period
        yyyymm, year, month, years = period_from_date(ship_by)
        R[0]  = 'MTC-%s-%s-%s' % (so_hdr['so_no'], s(so_ln.get('OrderlineE28')), sku)
        R[2]  = so_hdr['order_date']
        R[4]  = so_hdr['ship_code']
        R[6]  = so_hdr['div'] or '65'
        R[7]  = 'MTC-SOURCING'
        R[8]  = so_hdr['factory']
        R[9]  = 'SOURCING'
        R[10] = so_hdr['cust_id']
        R[11] = so_hdr['cust_name']
        R[14] = so_hdr['terms']
        R[16] = so_hdr['fob']
        R[21] = sku
        R[22] = fg(fam, 1) or s(so_ln.get('DescE01') or so_ln.get('LineDesc'))
        R[23] = fg(fam, 2) or s(so_ln.get('DescE01') or so_ln.get('LineDesc'))
        R[24] = fg(fam, 3)
        R[25] = fg(fam, 2)
        so_out_qty = num(so_ln.get('Epidueqty'))
        if so_out_qty is None:
            so_out_qty = num(so_ln.get('CurorderqtyE28'))
        R[28] = so_out_qty   # outstanding (unshipped) qty only
        R[33] = num(so_ln.get('CurdocunitpriceE28'))
        R[35] = so_hdr['currency'] or 'USD'
        so_rate = fx_rate(so_hdr['currency'])
        R[36] = so_rate
        R[37] = (num(so_ln.get('CurdocunitpriceE28')) or 0) * so_rate
        R[39] = so_out_qty * (num(so_ln.get('CurdocunitpriceE28')) or 0)   # value outstanding qty only
        R[48] = R[39] * so_rate if R[39] is not None else None
        R[19] = fg(fam, 4) or s(so_ln.get('FamilynameE01'))
        R[20] = R[19]
        R[67] = fg(fam, 5)
        R[68] = fg(fam, 6)
        R[70] = yyyymm; R[71] = year; R[91] = month; R[92] = years
        R[79] = so_hdr['so_no']
        R[90] = commodity.get(sku.upper())
        R[93] = combined_cust.get(so_hdr['cust_id'].upper()) or so_hdr['cust_name']
        R[95] = disney.get(sku.upper()) or fg(fam, 5)
        R[97] = ship_by   # Ship By Date (Need By Date)

    # ----- purchase side -----
    if po_hdr is not None:
        po_rate = fx_rate(po_hdr['currency'])
        po_out_qty = num(po_ln.get('Epidueqty'))
        if po_out_qty is None:
            po_out_qty = num(po_ln.get('CurorderqtyJ10'))
        cost = po_out_qty * (num(po_ln.get('CurdocunitcostJ10')) or 0)   # outstanding qty only
        cost_usd = cost * po_rate if cost is not None else None
        R[28] = R[28] if R[28] is not None else po_out_qty
        R[35] = R[35] or po_hdr['currency'] or 'USD'
        R[36] = R[36] if R[36] is not None else po_rate   # PO-only rows: rate = PO currency
        R[50] = cost_usd
        R[55] = cost_usd
        R[72] = cost_usd
        R[73] = po_hdr['vendor_id']
        R[74] = po_hdr['vendor_name']
        R[75] = po_hdr['vendor_name']
        R[76] = po_hdr['currency'] or 'USD'
        R[77] = num(po_ln.get('CurdocunitcostJ10'))   # LST_RCP_PR (ORI CUR) - doc unit cost
        R[78] = cost                                  # TTLLST_RCP_PR (ORI CUR) - doc total
        R[80] = po_hdr['po_no']
        R[81] = po_rate
        R[82] = (num(po_ln.get('CurdocunitcostJ10')) or 0) * po_rate   # U / PO USD
        R[83] = cost_usd
        R[94] = combined_vend.get(po_hdr['vendor_id'].upper()) or po_hdr['vendor_name']
        R[96] = vendor_country.get(po_hdr['vendor_name'].upper())
        if so_hdr is None:
            po_ship_by = date_only(po_ln.get('CurduedateJ10')) or po_hdr['order_date']
            yyyymm, year, month, years = period_from_date(po_ship_by)
            R[0]  = 'MTC-PO-%s-%s-%s' % (po_hdr['po_no'], s(po_ln.get('PolineJ10')), sku)
            R[2]  = po_hdr['order_date']
            R[6]  = po_hdr['div'] or '65'
            R[7]  = 'MTC-SOURCING'
            R[8]  = po_hdr['factory']
            R[9]  = 'SOURCING'
            R[14] = po_hdr['terms']
            R[16] = po_hdr['fob']
            R[19] = fg(fam, 4) or s(po_ln.get('FamilynameE01'))
            R[20] = R[19]
            R[21] = sku
            R[22] = fg(fam, 1) or s(po_ln.get('DescE01') or po_ln.get('LineDesc'))
            R[23] = fg(fam, 2) or s(po_ln.get('DescE01') or po_ln.get('LineDesc'))
            R[24] = fg(fam, 3)
            R[67] = fg(fam, 5)
            R[68] = fg(fam, 6)
            R[70] = yyyymm; R[71] = year; R[91] = month; R[92] = years
            R[79] = po_hdr['so_no']
            R[90] = commodity.get(sku.upper())
            R[95] = disney.get(sku.upper()) or fg(fam, 5)
            R[97] = po_ship_by   # PO Due Date as ship-by proxy (no SO)

    R[69] = 'Forecast'
    sales_usd = R[48]; prime_usd = R[55]
    if (sales_usd is not None and prime_usd is not None
            and isinstance(sales_usd, (int, float)) and isinstance(prime_usd, (int, float))):
        gross = sales_usd - prime_usd
        R[56] = gross
        R[84] = gross
        if sales_usd:
            R[57] = gross / sales_usd
    return R


# ---------------- join: SO.PonumE27 == PO.PonumJ16 ----------------
po_by_number = {}
for po in po_entities:
    po_by_number[s(po.get('PonumJ16'))] = po

# set of PO numbers referenced by SOs
linked_po_numbers = set(s(so.get('PonumE27')) for so in so_entities)

rows = []
matched_so = 0
matched_lines = 0
unmatched_so = []
for so in so_entities:
    so_hdr = so_header_fields(so)
    linked_po_no = so_hdr['linked_po']
    po = po_by_number.get(linked_po_no)
    so_lines = so.get('DetailList') or []
    if po is None:
        unmatched_so.append((so_hdr['so_no'], linked_po_no))
        for ln in so_lines:
            rows.append(build_row(so_hdr, ln, None, None))
        continue
    matched_so += 1
    po_hdr = po_header_fields(po)
    po_lines = po.get('DetailList') or []
    # index PO lines by SKU
    po_line_by_sku = {}
    for pl in po_lines:
        po_line_by_sku.setdefault(s(pl.get('CurpartnumJ10')).upper(), []).append(pl)
    consumed_po_lines = set()
    for ln in so_lines:
        sku = s(ln.get('CurpartnumE28')).upper()
        plist = po_line_by_sku.get(sku, [])
        if plist:
            pl = plist.pop(0)   # FIFO: pair SO line i with PO line i for duplicate SKUs
            consumed_po_lines.add(id(pl))
            rows.append(build_row(so_hdr, ln, po_hdr, pl))
            matched_lines += 1
        else:
            rows.append(build_row(so_hdr, ln, po_hdr, None))
    # any PO lines not consumed -> PO-only rows (same linked PO)
    for pl in po_lines:
        if id(pl) not in consumed_po_lines:
            rows.append(build_row(None, None, po_hdr, pl))

# POs not referenced by any SO -> PO-only rows
extra_po = 0
for po in po_entities:
    pno = s(po.get('PonumJ16'))
    if pno in linked_po_numbers:
        continue
    po_hdr = po_header_fields(po)
    for pl in (po.get('DetailList') or []):
        rows.append(build_row(None, None, po_hdr, pl))
        extra_po += 1

print("SO entities: %d" % len(so_entities))
print("SOs with linked PO found in cache: %d ; still missing: %d" % (matched_so, len(unmatched_so)))
if unmatched_so:
    print("  still-missing SOs (SO -> linked PO):", unmatched_so[:15])
print("matched SO lines (SO+PO): %d" % matched_lines)
print("extra PO-only rows (unlinked POs): %d" % extra_po)
print("TOTAL data rows: %d" % len(rows))


def write_xlsx(path, headers, rows):
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             '<sheetData>']
    def esc(v):
        return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))
    def cell(v):
        if v is None or v == '':
            return '<c/>'
        if isinstance(v, (int, float)):
            return '<c><v>%s</v></c>' % v
        return '<c t="inlineStr"><is><t>%s</t></is></c>' % esc(str(v))
    def rowxml(r, values):
        return '<row r="%d">%s</row>' % (r, "".join(cell(v) for v in values))
    sheet.append(rowxml(1, headers))
    for i, r in enumerate(rows, start=2):
        sheet.append(rowxml(i, r))
    sheet.append('</sheetData></worksheet>')
    sheet_xml = "".join(sheet)
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                     '</Types>')
    rels_root = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
               '</Relationships>')
    ws_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", ws_rels)


out = ROOT + "/MTC_Data_20261120_R7.xlsx"
write_xlsx(out, HEADERS, rows)
print("wrote %s" % out)
