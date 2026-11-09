#!/usr/bin/env python3
"""Health check for the sites in sites.json.

Usage: check.py SITES_JSON PREV_STATUS_JSON OUT_STATUS_JSON
Writes the new status file. A site counts as down after 2 failed checks in a row;
main() returns the list of sites that just went down (used for alerts).
On the office Synology this is run by agent.py, which also uploads the result.
"""
import json, socket, ssl, sys, time, urllib.error, urllib.request
from datetime import datetime, timezone
from urllib.parse import urlparse

KEEP_SECONDS = 7 * 24 * 3600
TIMEOUT = 20
UP_CODES_4XX = {401, 403, 405, 429}  # server answered; just needs login / blocks bots
UA = "Mozilla/5.0 (compatible; us-sm-healthcheck/1.0; +https://us-sm.github.io/us-sm/health/)"


def probe(url):
    t0 = time.time()
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            r.read(2048)
            code = r.status
    except urllib.error.HTTPError as e:
        code = e.code
    except Exception as e:  # DNS, timeout, TLS, refused...
        return {"up": False, "code": None, "ms": None, "error": short_err(e)}
    ms = int((time.time() - t0) * 1000)
    up = code < 400 or code in UP_CODES_4XX
    return {"up": up, "code": code, "ms": ms, "error": None if up else f"HTTP {code}"}


def short_err(e):
    r = getattr(e, "reason", e)
    if isinstance(r, socket.timeout) or "timed out" in str(r):
        return "Timeout"
    if isinstance(r, ssl.SSLCertVerificationError) or "CERTIFICATE" in str(r).upper():
        return "SSL certificate error"
    if isinstance(r, socket.gaierror):
        return "DNS lookup failed"
    if isinstance(r, ConnectionRefusedError):
        return "Connection refused"
    return str(r)[:120]


def cert_days(url):
    u = urlparse(url)
    if u.scheme != "https":
        return None
    try:
        ctx = ssl.create_default_context()
        with socket.create_connection((u.hostname, u.port or 443), timeout=TIMEOUT) as s:
            with ctx.wrap_socket(s, server_hostname=u.hostname) as ss:
                exp = ssl.cert_time_to_seconds(ss.getpeercert()["notAfter"])
        return int((exp - time.time()) // 86400)
    except Exception:
        return None


def main(sites_path, prev_path, out_path):
    sites = json.load(open(sites_path, encoding="utf-8"))
    try:
        prev = {s["url"]: s for s in json.load(open(prev_path, encoding="utf-8"))["sites"]}
    except Exception:
        prev = {}
    now = int(time.time())
    out, alert = [], []
    for site in sites:
        url = site["url"]
        p = prev.get(url, {})
        r = probe(url)
        if not r["up"]:  # retry once to avoid false alarms
            time.sleep(5)
            r2 = probe(url)
            if r2["up"]:
                r = r2
        fails = 0 if r["up"] else p.get("fails", 0) + 1
        history = [h for h in p.get("history", []) if h[0] > now - KEEP_SECONDS]
        history.append([now, r["ms"] if r["up"] else None])

        incidents = p.get("incidents", [])
        open_inc = incidents and incidents[-1].get("end") is None
        if fails == 2 and not open_inc:  # confirmed down
            incidents.append({"start": history[-2][0] if len(history) > 1 else now, "end": None, "error": r["error"]})
            alert.append(f"{site['name']} is DOWN ({r['error']})")
        elif r["up"] and open_inc:
            incidents[-1]["end"] = now
        incidents = incidents[-20:]

        certs = p.get("certDays")
        if p.get("certChecked", 0) < now - 6 * 3600:  # check certificate a few times a day
            certs = cert_days(url)
            cert_checked = now
        else:
            cert_checked = p.get("certChecked")

        out.append({
            "name": site["name"], "url": url, "up": r["up"], "code": r["code"], "ms": r["ms"],
            "error": r["error"], "fails": fails, "checked": now,
            "certDays": certs, "certChecked": cert_checked,
            "history": history, "incidents": incidents,
        })
        print(f"{site['name']}: {'UP' if r['up'] else 'DOWN'} code={r['code']} ms={r['ms']} err={r['error']}", file=sys.stderr)

    json.dump({"updated": datetime.now(timezone.utc).isoformat(timespec="seconds"), "sites": out},
              open(out_path, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    return alert


if __name__ == "__main__":
    alerts = main(*sys.argv[1:4])
    if alerts:
        print("alert=1")
        print("message=" + "; ".join(alerts))
