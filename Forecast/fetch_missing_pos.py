# -*- coding: utf-8 -*-
"""
Fetch the POs referenced by the SOs (via PonumE27) but missing from the raw cache.
Every SO has a linked PO in Epicor (1:1); the earlier PO header export was a stale
snapshot (cutoff ~PO 1004329 / Oct 30), so ~139 POs are missing.

Reads EPICOR_USER / EPICOR_PASS from env vars. Merges results into
_work_mtc_details_raw.json so a rebuild needs no refetch.
"""
import json, os, sys, time
import urllib.request, urllib.error, http.cookiejar

BASE = "https://portal.meyer-msc.com/EPICOR-SUB"
USER = os.environ.get("EPICOR_USER")
PASS = os.environ.get("EPICOR_PASS")
ROOT = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(ROOT, "_work_mtc_details_raw.json")


class Epicor:
    def __init__(self):
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))
        self.csrf = None

    def _req(self, url, data=None):
        headers = {"X-Requested-With": "XMLHttpRequest", "User-Agent": "Mozilla/5.0"}
        if self.csrf:
            headers["X-CSRF-TOKEN"] = self.csrf
        body = None
        if data is not None:
            body = json.dumps(data).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=body, headers=headers,
                                     method="POST" if data is not None else "GET")
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
        self._req(BASE + "/Account/Login")
        self.csrf = self._cookie("CSRF-TOKEN")
        if not self.csrf:
            sys.exit("no CSRF-TOKEN cookie from login page")
        status, body = self._req(BASE + "/Account/Login", {"Username": USER, "Password": PASS})
        try:
            j = json.loads(body)
            if j.get("Error"):
                sys.exit("login failed: " + json.dumps(j["Error"]))
        except json.JSONDecodeError:
            pass
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


def main():
    if not USER or not PASS:
        sys.exit("set EPICOR_USER and EPICOR_PASS environment variables")
    raw = json.load(open(RAW, encoding="utf-8"))
    so_dict = raw["so"]; po_dict = raw["po"]

    missing = sorted(set(
        str(so.get("PonumE27") or "").strip()
        for so in so_dict.values()
        if str(so.get("PonumE27") or "").strip() and str(so.get("PonumE27")).strip() not in po_dict
    ))
    print("missing POs to fetch: %d" % len(missing))

    api = Epicor()
    api.login()
    print("logged in as %s" % USER)

    ok = 0; fail = []
    for idx, p in enumerate(missing):
        ent, err = api.retrieve("PurchaseOrder/PurchaseOrder", "MTC-%s" % p)
        if ent is None:
            fail.append((p, str(err)[:120]))
            print("  FAIL %s -> %s" % (p, str(err)[:120]))
        else:
            po_dict[p] = ent
            ok += 1
        if (idx + 1) % 25 == 0:
            print("  ... %d/%d" % (idx + 1, len(missing)))
            json.dump(raw, open(RAW, "w", encoding="utf-8"))
        time.sleep(0.05)

    json.dump(raw, open(RAW, "w", encoding="utf-8"))
    print("fetched OK: %d ; failed: %d" % (ok, len(fail)))
    print("total POs in cache now: %d" % len(po_dict))
    if fail:
        print("FAILED list:")
        for p, e in fail:
            print("   %s : %s" % (p, e))


if __name__ == "__main__":
    main()
