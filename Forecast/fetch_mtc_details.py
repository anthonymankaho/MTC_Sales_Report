# -*- coding: utf-8 -*-
"""
Fetch MTC Sales Order + Purchase Order detail lines from the Epicor portal
(portal.meyer-msc.com/EPICOR-SUB) and combine with the header report to
reproduce an SCF017-style detail report.

Login: alex@meyer-msc.com (Team Leader, has MTC company access).
For each SO # in "EPI_SO Header_*.xlsx" and each PO # in "EPI_PO Header_*.xlsx",
Retrieve the extend entity (which returns DetailList) and flatten header+lines.

Outputs (in project root):
  EPI_SO Detail_<date>_R0.xlsx
  EPI_PO Detail_<date>_R0.xlsx
Raw JSON is cached at _work_mtc_details_raw.json so a rebuild needs no refetch.
"""
import json, os, re, sys, time, zipfile, glob
from xml.etree import ElementTree as ET
import urllib.request, urllib.parse, http.cookiejar

BASE = "https://portal.meyer-msc.com/EPICOR-SUB"
USER = os.environ.get("EPICOR_USER")
PASS = os.environ.get("EPICOR_PASS")
ROOT = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(ROOT, "_work_mtc_details_raw.json")
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def xlsx_rows(path):
    """Read first worksheet of an xlsx, return list of rows (list of cell values)."""
    z = zipfile.ZipFile(path)
    shared = []
    try:
        sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in sst.findall(M + "si"):
            shared.append("".join(t.text or "" for t in si.iter(M + "t")))
    except KeyError:
        pass
    root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    out = []
    for row in root.iter(M + "row"):
        cells = []
        for c in row.findall(M + "c"):
            v = c.find(M + "v")
            if v is None:
                cells.append(None)
                continue
            t = c.get("t")
            if t == "s":
                cells.append(shared[int(v.text)])
            elif t == "inlineStr":
                ise = c.find(M + "is")
                cells.append("".join(x.text or "" for x in ise.iter(M + "t")) if ise is not None else None)
            else:
                cells.append(v.text)
        out.append(cells)
    return out


def find_header_file(prefix):
    hits = glob.glob(os.path.join(ROOT, prefix + "*.xlsx"))
    if not hits:
        sys.exit("missing header file: " + prefix)
    return sorted(hits)[-1]


def extract_order_numbers():
    so_rows = xlsx_rows(find_header_file("EPI_SO Header"))
    po_rows = xlsx_rows(find_header_file("EPI_PO Header"))
    # locate SO # / PO # column by header text
    def col_index(hdr, *names):
        for i, h in enumerate(hdr):
            if h and any(n.lower() == str(h).strip().lower() for n in names):
                return i
        return None

    so_hdr = so_rows[0]
    po_hdr = po_rows[0]
    so_col = col_index(so_hdr, "SO #", "SO#", "SO No")
    po_col = col_index(po_hdr, "PO #", "PO#", "PO No")
    if so_col is None or po_col is None:
        sys.exit("could not find SO # / PO # column")
    so_nums = [int(float(r[so_col])) for r in so_rows[1:] if r[so_col] is not None]
    po_nums = [int(float(r[po_col])) for r in po_rows[1:] if r[po_col] is not None]
    return so_nums, po_nums


class Epicor:
    def __init__(self):
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))
        self.csrf = None

    def _req(self, url, data=None):
        headers = {
            "X-Requested-With": "XMLHttpRequest",
            "User-Agent": "Mozilla/5.0",
        }
        if self.csrf:
            headers["X-CSRF-TOKEN"] = self.csrf
        body = None
        if data is not None:
            body = json.dumps(data).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=body, headers=headers, method="POST" if data is not None else "GET")
        try:
            with self.opener.open(req, timeout=60) as r:
                return r.status, r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            return e.code, e.read().decode("utf-8", "replace")

    def _cookie(self, name):
        for c in self.jar:
            if c.name == name:
                return c.value
        return None

    def login(self):
        # GET login page to seed CSRF-TOKEN cookie
        self._req(BASE + "/Account/Login")
        self.csrf = self._cookie("CSRF-TOKEN")
        if not self.csrf:
            sys.exit("no CSRF-TOKEN cookie from login page")
        status, body = self._req(BASE + "/Account/Login", {"Username": USER, "Password": PASS})
        # success returns {} ; failure returns Error
        try:
            j = json.loads(body)
            if j.get("Error"):
                sys.exit("login failed: " + json.dumps(j["Error"]))
        except json.JSONDecodeError:
            pass
        # refresh CSRF after login
        self._req(BASE + "/")
        self.csrf = self._cookie("CSRF-TOKEN")

    def retrieve(self, kind, entity_id):
        url = BASE + "/Services/" + kind + "/Retrieve"
        status, body = self._req(url, {"EntityId": entity_id})
        try:
            j = json.loads(body)
        except json.JSONDecodeError:
            return None, "non-JSON (%d): %s" % (status, body[:200])
        if j.get("Error"):
            return None, j["Error"]
        return j.get("Entity"), None


def norm(v):
    if v is None:
        return ""
    return v


def date_only(v):
    s = str(v)
    return s[:10] if s else ""


def build_so_rows(entities):
    rows = []
    for e in entities:
        hdr = [
            norm(e.get("CompanyE27")),              # Company
            norm(e.get("OrdernumE27")),             # SO #
            str(e.get("VernumE27") or "0").zfill(2),  # REV.
            norm(e.get("CustidE27")),               # Customer
            norm(e.get("Name")),                    # Customer Name
            norm(e.get("PonumE27")),                # Customer PO #
            date_only(e.get("OrderDate")),          # Order Date
            norm(e.get("StatusE27")),               # Status
            norm(e.get("ShptonumE27")),             # Ship Code
            norm(e.get("TermsCode")),               # Terms
            norm(e.get("Fob")),                     # FOB
            norm(e.get("CurrencyCode")),            # Currency
            norm(e.get("CodvcdE27")),               # Company DIV.
            norm(e.get("FactcdE27")),               # Factory
            norm(e.get("OrdtyE27")),                # Order Type
            norm(e.get("PoeE27")),                  # Port of Entry
            norm(e.get("CooE27")),                  # Country of Origin
        ]
        for ln in e.get("DetailList") or []:
            qty = ln.get("CurorderqtyE28") or 0
            price = ln.get("CurdocunitpriceE28") or 0
            rows.append(hdr + [
                norm(ln.get("OrderlineE28")),            # Line #
                norm(ln.get("CurpartnumE28")),           # SKU
                norm(ln.get("DescE01") or ln.get("LineDesc")),  # Description
                norm(ln.get("FamilynameE01")),           # Family
                qty,                                     # Order Qty
                ln.get("Epidueqty"),                     # Outstanding Qty
                price,                                   # Unit Price
                round(qty * price, 2) if qty else 0.0,   # Amount
                norm(ln.get("Ium")),                     # UOM
                norm(ln.get("UnitPack")),                # Unit Pack
                date_only(ln.get("CurrequestdateE28")),  # Request Date
                date_only(ln.get("CurneedbydateE28")),   # Need By Date
                norm(ln.get("LineDesc")),                # Line Description
            ])
    return rows


def build_po_rows(entities):
    rows = []
    for e in entities:
        hdr = [
            norm(e.get("CompanyJ16")),              # Company
            norm(e.get("PonumJ16")),                # PO #
            str(e.get("VernumJ16") or "0").zfill(2),  # REV.
            norm(e.get("VendoridJ16")),             # Supplier
            norm(e.get("Name")),                    # Supplier Name
            norm(e.get("OrdernumJ16")),             # SO #
            norm(e.get("CustidJ16")),               # Customer
            norm(e.get("CustpoJ16")),               # Customer PO
            norm(e.get("BuyerId")),                 # Buyer
            date_only(e.get("OrderDate")),          # Order Date
            norm(e.get("StatusJ16")),               # Status
            norm(e.get("TermsCode")),               # Terms
            norm(e.get("FobJ16") or e.get("Fob")),  # FOB
            norm(e.get("CurrencyCode")),            # Currency
            norm(e.get("CodvcdJ16")),               # Company DIV.
            norm(e.get("FactcdJ16")),               # Factory
            norm(e.get("RejreaJ16")),               # Reject Reason
        ]
        for ln in e.get("DetailList") or []:
            qty = ln.get("CurorderqtyJ10") or 0
            cost = ln.get("CurdocunitcostJ10") or 0
            rows.append(hdr + [
                norm(ln.get("PolineJ10")),               # Line #
                norm(ln.get("CurpartnumJ10")),           # SKU
                norm(ln.get("DescE01") or ln.get("LineDesc")),  # Description
                norm(ln.get("FamilynameE01")),           # Family
                qty,                                     # Order Qty
                ln.get("Epidueqty"),                     # Outstanding Qty
                cost,                                    # Unit Cost
                round(qty * cost, 2) if qty else 0.0,    # Amount
                norm(ln.get("Ium")),                     # UOM
                date_only(ln.get("CurduedateJ10")),      # Due Date
                norm(ln.get("LineDesc")),                # Line Description
            ])
    return rows


SO_HEADERS = ["Company", "SO #", "REV.", "Customer", "Customer Name", "Customer PO #",
              "Order Date", "Status", "Ship Code", "Terms", "FOB", "Currency",
              "Company DIV.", "Factory", "Order Type", "Port of Entry", "Country of Origin",
              "Line #", "SKU", "Description", "Family", "Order Qty", "Outstanding Qty",
              "Unit Price", "Amount", "UOM", "Unit Pack", "Request Date", "Need By Date",
              "Line Description"]

PO_HEADERS = ["Company", "PO #", "REV.", "Supplier", "Supplier Name", "SO #", "Customer",
              "Customer PO", "Buyer", "Order Date", "Status", "Terms", "FOB", "Currency",
              "Company DIV.", "Factory", "Reject Reason",
              "Line #", "SKU", "Description", "Family", "Order Qty", "Outstanding Qty",
              "Unit Cost", "Amount", "UOM", "Due Date", "Line Description"]


def write_xlsx(path, headers, rows):
    """Minimal single-sheet xlsx using inline strings (no sharedStrings)."""
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             '<sheetData>']
    def esc(v):
        return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace('"', "&quot;"))
    def cell(v, style=None):
        s = ' s="%d"' % style if style else ""
        if v is None:
            return '<c%s/>' % s
        sv = str(v)
        if isinstance(v, (int, float)):
            return '<c%s><v>%s</v></c>' % (s, sv)
        return '<c%s t="inlineStr"><is><t>%s</t></is></c>' % (s, esc(sv))
    def rowref(r, values):
        return '<row r="%d">%s</row>' % (r, "".join(cell(v) for v in values))
    sheet.append(rowref(1, headers))
    for i, r in enumerate(rows, start=2):
        sheet.append('<row r="%d">%s</row>' % (i, "".join(cell(v) for v in r)))
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
                '<sheets><sheet name="Detail" sheetId="1" r:id="rId1"/></sheets></workbook>')
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


def main():
    if not USER or not PASS:
        sys.exit("set EPICOR_USER and EPICOR_PASS environment variables")
    so_nums, po_nums = extract_order_numbers()
    print("SO #s from header: %d" % len(so_nums))
    print("PO #s from header: %d" % len(po_nums))

    api = Epicor()
    api.login()
    print("logged in as %s" % USER)

    result = {"so": {}, "po": {}}
    if os.path.exists(RAW):
        try:
            result = json.load(open(RAW))
            print("loaded cache (%d SO, %d PO)" % (len(result["so"]), len(result["po"])))
        except Exception:
            result = {"so": {}, "po": {}}

    def fetch(kind, entity_ids, key):
        data = result[key]
        missing = [i for i in entity_ids if str(i) not in data]
        print("%s: %d total, %d already cached, %d to fetch" % (kind, len(entity_ids), len(entity_ids) - len(missing), len(missing)))
        for idx, i in enumerate(missing):
            ent, err = api.retrieve(kind, "MTC-%s" % i)
            if ent is None:
                print("  WARN %s %s -> %s" % (kind, i, err))
                data[str(i)] = None
            else:
                data[str(i)] = ent
            if (idx + 1) % 20 == 0:
                print("  ... %d/%d" % (idx + 1, len(missing)))
                json.dump(result, open(RAW, "w"))
            time.sleep(0.05)

    fetch("SalesOrder/SalesOrder", so_nums, "so")
    fetch("PurchaseOrder/PurchaseOrder", po_nums, "po")
    json.dump(result, open(RAW, "w"))

    so_entities = [result["so"][str(i)] for i in so_nums if result["so"].get(str(i))]
    po_entities = [result["po"][str(i)] for i in po_nums if result["po"].get(str(i))]
    print("retrieved SO entities: %d/%d" % (len(so_entities), len(so_nums)))
    print("retrieved PO entities: %d/%d" % (len(po_entities), len(po_nums)))

    so_rows = build_so_rows(so_entities)
    po_rows = build_po_rows(po_entities)
    print("SO detail rows: %d" % len(so_rows))
    print("PO detail rows: %d" % len(po_rows))

    stamp = "20261120"
    so_out = os.path.join(ROOT, "EPI_SO Detail_%s_R0.xlsx" % stamp)
    po_out = os.path.join(ROOT, "EPI_PO Detail_%s_R0.xlsx" % stamp)
    write_xlsx(so_out, SO_HEADERS, so_rows)
    write_xlsx(po_out, PO_HEADERS, po_rows)
    print("wrote %s" % so_out)
    print("wrote %s" % po_out)


if __name__ == "__main__":
    main()
