import os, urllib.request, json
import build_board as cfg

print("=" * 60)
print("STEP 1 — what URL override does build_board.py hold?")
print("  ESPN_URL_OVERRIDE =", repr(cfg.ESPN_URL_OVERRIDE))
print("  USE_ESPN          =", cfg.USE_ESPN)

url = cfg.ESPN_URL_OVERRIDE or (
    "https://lm-api-reads.fantasy.espn.com/apis/v3/games/flb/"
    f"seasons/{cfg.ESPN_SEASON}/segments/0/leagues/{cfg.ESPN_LEAGUE_ID}"
    "?view=mTeam&view=mRoster")
print("\nSTEP 2 — the address the build will actually call:")
print(" ", url)
if not cfg.ESPN_URL_OVERRIDE:
    print("  >>> WARNING: override is EMPTY, so it is calling ESPN directly.")
elif "workers.dev" not in cfg.ESPN_URL_OVERRIDE:
    print("  >>> WARNING: override does not look like a workers.dev relay URL.")

print("\nSTEP 3 — calling that address now...")
try:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (diagnose)"})
    with urllib.request.urlopen(req, timeout=25) as r:
        body = r.read().decode("utf-8", "replace")
    print("  HTTP status: 200 OK")
    print("  first 200 characters of the response:")
    print("  " + body[:200])
    try:
        data = json.loads(body)
        teams = data.get("teams", [])
        print(f"\n  Parsed JSON OK. teams found: {len(teams)}")
        if teams:
            entries = (teams[0].get("roster") or {}).get("entries", [])
            print(f"  first team roster entries: {len(entries)}")
            print("  >>> SUCCESS: this address returns real roster data.")
    except Exception as e:
        print("  Could NOT parse as JSON:", e)
except Exception as e:
    print("  CALL FAILED:", type(e).__name__, "-", e)
print("=" * 60)
