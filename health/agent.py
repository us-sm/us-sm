#!/usr/bin/env python3
"""Office health-check agent (runs on the Synology every 10 minutes).

Setup (see health/README.md):
  1. Put this file and token.txt (a GitHub fine-grained token with
     Contents: Read and write on us-sm/us-sm only) in one folder on the NAS.
  2. Task Scheduler -> user-defined script, every 10 minutes:
         python3 /volume1/scripts/healthcheck/agent.py
     and tick "send run details by email ... only when the script terminates abnormally".

Each run downloads the latest check.py and sites.json from the repo (so editing
sites.json on GitHub is enough to add a site), checks the sites from inside the
office, and uploads status.json to the status-data branch that the status page reads.
Exit code 1 = a site just went down (Synology emails you); 2 = the agent itself failed.
"""
import base64, json, os, sys, urllib.error, urllib.request

REPO = "us-sm/us-sm"
DATA_BRANCH = "status-data"
HERE = os.path.dirname(os.path.abspath(__file__))
API = f"https://api.github.com/repos/{REPO}/contents/status.json"
RAW = f"https://raw.githubusercontent.com/{REPO}/main/health/"


def token():
    t = os.environ.get("GH_TOKEN")
    if not t:
        with open(os.path.join(HERE, "token.txt"), encoding="utf-8") as f:
            t = f.read().strip()
    return t


def gh(method, url, body=None):
    req = urllib.request.Request(url, method=method, data=json.dumps(body).encode() if body else None, headers={
        "Authorization": "Bearer " + token(),
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "us-sm-health-agent",
        "Content-Type": "application/json",
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def download(name):
    with urllib.request.urlopen(RAW + name, timeout=30) as r:
        data = r.read()
    with open(os.path.join(HERE, name), "wb") as f:
        f.write(data)


def run():
    # 1. latest checker + site list (fall back to the local copy if GitHub is unreachable)
    for name in ("check.py", "sites.json"):
        try:
            download(name)
        except Exception as e:
            if not os.path.exists(os.path.join(HERE, name)):
                raise
            print(f"warning: could not update {name}: {e}", file=sys.stderr)
    sys.path.insert(0, HERE)
    import check  # noqa: E402

    # 2. previous results (straight from the API, so there is no CDN cache)
    prev_path, out_path = os.path.join(HERE, "prev.json"), os.path.join(HERE, "status.json")
    sha = None
    try:
        cur = gh("GET", f"{API}?ref={DATA_BRANCH}")
        sha = cur["sha"]
        with open(prev_path, "wb") as f:
            f.write(base64.b64decode(cur["content"]))
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
        with open(prev_path, "w") as f:
            f.write('{"sites":[]}')

    # 3. check and upload
    alerts = check.main(os.path.join(HERE, "sites.json"), prev_path, out_path)
    with open(out_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    body = {"message": "Health check from office", "content": content, "branch": DATA_BRANCH}
    if sha:
        body["sha"] = sha
    gh("PUT", API, body)
    return alerts


if __name__ == "__main__":
    try:
        alerts = run()
    except Exception as e:
        print(f"health agent failed: {e}", file=sys.stderr)
        sys.exit(2)
    if alerts:
        print("DOWN: " + "; ".join(alerts))
        sys.exit(1)
    print("ok")
