# System Status (health check)

ALPS2 and Hanjin ERP only accept connections from the office IP, so the check runs
**on the office Synology**, and the page at `health/` shows the results.

```
Synology (every 10 min) --agent.py--> GitHub branch status-data/status.json <-- health/index.html
```

## Add or change a site
Edit `health/sites.json` on GitHub. The Synology picks it up on its next run.

## Synology setup (one time)
1. **GitHub token**: GitHub → Settings → Developer settings → Fine-grained tokens → Generate.
   - Resource owner: `us-sm`, Repository access: only `us-sm/us-sm`
   - Permissions: **Contents: Read and write** (nothing else)
   - Set the expiration you want and put a reminder to renew it.
2. In File Station, create a folder, e.g. `scripts/healthcheck`, and upload:
   - `agent.py` (from this folder)
   - `token.txt` containing only the token
3. Control Panel → Task Scheduler → Create → Scheduled Task → User-defined script
   - General: user `root`
   - Schedule: Daily, first run 00:00, repeat **every 10 minutes**, last run 23:50
   - Task Settings → tick "Send run details by email" and "only when the script terminates abnormally"
   - Run command: `python3 /volume1/scripts/healthcheck/agent.py`
4. Select the task → **Run**, then open the status page. Data shows up within a minute.

Exit codes: `0` ok, `1` a site just went down (Synology sends the email, once per outage), `2` the agent itself failed
(no internet, expired token...).

If the NAS has no `python3`: Package Center → install **Python 3.9**, then use `python3.9` in the command.
