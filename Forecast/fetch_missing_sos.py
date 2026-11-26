# -*- coding: utf-8 -*-
"""
Fetch the SOs referenced by the POs (via OrdernumJ16) but missing from the raw cache.
The SO header export was also a partial snapshot (235 confirmed SOs); the PO export
references 209 further SOs. Fetch them so the 1:1 SO<->PO picture is complete.

Reads EPICOR_USER / EPICOR_PASS from env vars. Merges into _work_mtc_details_raw.json.
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

    existing_so = set(str(v.get("OrdernumE27")) for v in so_dict.values())
    # SO numbers referenced by any PO via OrdernumJ16
    referenced = set()
    for p in po_dict.values():
        r = str(p.get("OrdernumJ16")).strip()
        if r:
            referenced.add(r)
    missing = sorted(int(x) for x in referenced if x not in existing_so and x.isdigit())
    print("SO#s referenced by POs but missing from cache: %d" % len(missing))

    api = Epicor()
    api.login()
    print("logged in as %s" % USER)

    ok = 0; fail = []
    for idx, sn in enumerate(missing):
        ent, err = api.retrieve("SalesOrder/SalesOrder", "MTC-%s" % sn)
        if ent is None:
            fail.append((sn, str(err)[:120]))
            print("  FAIL %s -> %s" % (sn, str(err)[:120]))
        else:
            so_dict[str(sn)] = ent
            ok += 1
        if (idx + 1) % 25 == 0:
            print("  ... %d/%d" % (idx + 1, len(missing)))
            json.dump(raw, open(RAW, "w", encoding="utf-8"))
        time.sleep(0.05)

    json.dump(raw, open(RAW, "w", encoding="utf-8"))
    print("fetched OK: %d ; failed: %d" % (ok, len(fail)))
    print("total SOs in cache now: %d" % len(so_dict))
    if fail:
        print("FAILED list:", fail)


if __name__ == "__main__":
    main()
