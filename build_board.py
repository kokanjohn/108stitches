#!/usr/bin/env python3
"""
108 Stitches — keeper board generator
-------------------------------------
Rosters come from ESPN (live source of truth for who's on each team now).
The Pre-Draft Worksheet is the overlay: keeper prices, price history, and how
each drafted player was acquired. Every current player is tagged:

    Kept    - drafted and on that team's keeper list (carries price history)
    Auction - drafted at auction, not kept (fresh price, no history)
    Traded  - acquired from another team via trade (inherits their sheet price)
    FA      - added from free agency mid-season (no auction value)

    python3 build_board.py

If ESPN can't be reached, it falls back to the draft-day rosters from the sheet.
"""
import unicodedata, re, json, sys
from pathlib import Path
from openpyxl import load_workbook

# ---- CONFIG -----------------------------------------------------------------
HERE      = Path(__file__).parent
WORKSHEET = HERE / "2026_Pre-Draft_Worksheet.xlsx"
TEMPLATE  = HERE / "template.html"
OUTPUT    = HERE / "108-stitches-keeper-board.html"
LEAGUE, SEASON, BUYIN = "108 Stitches", 2026, 290

USE_ESPN       = True
ESPN_LEAGUE_ID = 35759
ESPN_SEASON    = 2026
ESPN_S2        = None            # public league -> None
ESPN_SWID      = None            # public league -> None
ESPN_JSON_FILE = None            # optional: path to saved league JSON (offline/testing)
ESPN_URL_OVERRIDE = "https://espn-relay.baseball-gm.workers.dev/"   # your Cloudflare relay
OWNER_ALIAS    = {}              # {"ESPN Name": "Sheet Owner Name"} if a person's name differs

NAME_FIX = {"Jak Caglianone":"Jac Caglianone", "Sam Basallo":"Samuel Basallo",
            "Agustin Ramirez":"Agustín Ramírez", "Augustin Ramirez":"Agustín Ramírez"}
DRAFT_LABEL = {"Waller":"Justin Waller","T. Carter":"Travis Carter","Linthicum":"Brandon Linthicum",
 "Schottmiller":"Matt Schottmiller","Amos":"Jonathan Amos","Bonham":"Andrew Bonham",
 "Walker":"Eric Walker","Eanes":"Casey Eanes","Martinez":"Sam Martinez","Hines":"Matt Hines",
 "Moore":"John Moore","Brady":"Doug Brady","J. Carter":"Jeff Carter","Mitchell":"Chris Mitchell",
 "Cooper":"Scott Cooper","Stone":"Jonathan Stone"}
KEEPER_TABS = set(DRAFT_LABEL) | {"Travis","Brandon","Jeff"}  # tab names differ from Draft labels
KEEPER_TABS = {"Waller","Travis","Brandon","Schottmiller","Amos","Bonham","Walker","Brady",
 "Hines","Moore","Martinez","Jeff","Mitchell","Cooper","Stone","Eanes"}
# -----------------------------------------------------------------------------

def norm(s): return ("" if s is None else str(s)).replace("\xa0"," ").strip()
def keyof(s):
    s = unicodedata.normalize("NFKD", norm(s)).encode("ascii","ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)
def toint(v):
    try: return int(v) if v not in (None,"") else None
    except (TypeError, ValueError): return None

def parse_worksheet(wb):
    """Return draft-day records (kept + auction), keyed index, and owner->team."""
    HIST, META, OWNER_TEAM = {}, {}, {}
    for i, r in enumerate(wb["Keeper Prices"].iter_rows(values_only=True)):
        if i == 0 or r is None: continue
        player = norm(r[2]) if len(r) > 2 else ""
        if not player: continue
        k = keyof(player)
        HIST[k] = {y: toint(r[idx]) for y, idx in zip((2022,2023,2024,2025), (5,6,7,8))}
        META[k] = {"name": NAME_FIX.get(player, player), "mlb": norm(r[3]) or "FA", "pos": norm(r[4])}
        owner, team = norm(r[1]), norm(r[0])
        if owner and owner not in OWNER_TEAM: OWNER_TEAM[owner] = team

    KEPT, KEEP_PRICE = set(), {}
    for tab in KEEPER_TABS:
        for r in list(wb[tab].iter_rows(values_only=True))[3:]:
            if r is None: continue
            name = norm(r[1]) if len(r) > 1 else ""
            if name and name.lower() != "name":
                KEPT.add(keyof(name)); KEEP_PRICE[keyof(name)] = toint(r[2]) if len(r) > 2 else None

    rows = list(wb["Draft"].iter_rows(values_only=True))
    cols = {c: DRAFT_LABEL[norm(rows[2][c])] for c in range(1, len(rows[2]), 2)
            if norm(rows[2][c]) in DRAFT_LABEL}
    records, index = [], {}
    for r in rows[6:32]:
        if r is None: continue
        slot = norm(r[0])
        for c, owner in cols.items():
            if c >= len(r): continue
            name = norm(r[c]); price = toint(r[c+1]) if c+1 < len(r) else None
            if not name or name.upper() == "FILL IN" or name.startswith("#"): continue
            k = keyof(name); kept = k in KEPT; m = META.get(k)
            disp = m["name"] if m else NAME_FIX.get(name, name)
            cur = price if price is not None else (KEEP_PRICE.get(k) if kept else None)
            hist = HIST.get(k, {}) if kept else {}
            rec = {"team": OWNER_TEAM.get(owner, owner), "owner": owner, "player": disp,
                   "mlb": (m["mlb"] if m else "FA"), "pos": slot, "kept": kept,
                   "tag": "kept" if kept else "auction",
                   "p": {**{y: (hist.get(y) if kept else None) for y in (2022,2023,2024,2025)}, 2026: cur}}
            records.append(rec)
            index[k] = rec
    return records, index, OWNER_TEAM

def classify(acq, sp, owner):
    a = acq or ""
    if a == "TRADE": return "traded"
    if a in ("ADD", "WAIVER", "FREEAGENT", "ADD_WAIVER"): return "fa"
    if a in ("DRAFT", "DRAFTED"):
        if sp and keyof(sp["owner"]) == keyof(owner): return "kept" if sp["kept"] else "auction"
        if sp: return "traded"          # drafted flag but sheet shows another owner -> treated as trade
        return "auction"                # drafted but not on the sheet
    if sp: return "kept" if sp["kept"] else "auction"
    return "fa"

def roster_rank(e, idx):
    """ESPN roster display order: active batters C->UT, bench/IL batters,
       active pitchers P/SP/RP, bench/IL pitchers."""
    BAT = {0:0, 1:1, 2:2, 3:3, 4:4, 6:5, 7:6, 5:7, 8:7, 9:7, 10:7, 11:8, 12:8}
    PIT = {13:0, 14:1, 15:2}
    s = e.get("slot_id")
    if s in (13, 14, 15):
        return (2, PIT.get(s, 9), idx)
    if s in (16, 17):
        sub = 0 if s == 16 else 1          # bench before IL
        return ((3 if e.get("is_pitcher") else 1), sub, idx)
    return (0, BAT.get(s, 9), idx)          # active batter

def build_from_espn(rosters, index, OWNER_TEAM):
    from espn_live import key as ekey, fuzzy
    keys = list(index.keys())
    # order every roster spot the way ESPN displays it, within each team
    ordered = sorted(enumerate(rosters),
                     key=lambda t: ((t[1]["team"] or t[1]["owner"] or ""), roster_rank(t[1], t[0])))
    records, teams, matched = [], {}, 0
    for _, e in ordered:
        owner, name, acq = e["owner"], e["player"], e["acq"]
        k = ekey(name); sp = index.get(k)
        if sp is None:
            nk = fuzzy(k, keys)
            if nk: sp = index.get(nk)
        if sp: matched += 1
        tag = classify(acq, sp, owner)
        team = e["team"] or OWNER_TEAM.get(owner, owner)     # ESPN's current team name wins
        s = e.get("slot_id"); is_pit = e.get("is_pitcher")
        grp = "pit" if (s in (13,14,15) or (s in (16,17) and is_pit)) else "bat"
        p = {**{y: (sp["p"][y] if sp else None) for y in (2022,2023,2024,2025)},
             2026: (sp["p"][2026] if sp else None)}
        rec = {"team": team, "owner": owner,
               "player": (sp["player"] if sp else name),
               "mlb": (sp["mlb"] if sp else (e["mlb"] or "FA")),
               "pos": (e["pos"] or (sp["pos"] if sp else "")),  # ESPN's current roster slot
               "elig": e.get("elig", ""), "grp": grp,
               "kept": tag == "kept", "tag": tag, "p": p}
        records.append(rec)
        t = teams.setdefault(team, {"team": team, "owner": owner, "count": 0, "kept": 0, "total2026": 0})
        t["count"] += 1; t["kept"] += 1 if tag == "kept" else 0
        if p[2026]: t["total2026"] += p[2026]
    return records, teams, matched

def team_meta(records):
    teams = {}
    for r in records:
        t = teams.setdefault(r["team"], {"team": r["team"], "owner": r["owner"], "count": 0, "kept": 0, "total2026": 0})
        t["count"] += 1; t["kept"] += 1 if r["tag"] == "kept" else 0
        if r["p"][2026]: t["total2026"] += r["p"][2026]
    return teams

def build():
    if not WORKSHEET.exists(): sys.exit(f"Can't find the worksheet: {WORKSHEET}")
    if not TEMPLATE.exists():  sys.exit("Can't find template.html next to this script.")
    wb = load_workbook(WORKSHEET, read_only=True, data_only=True)
    draft_records, index, OWNER_TEAM = parse_worksheet(wb)

    from datetime import datetime, timezone
    built_at = datetime.now(timezone.utc).strftime("%b %d, %Y · %H:%M UTC")
    records, teams, live = draft_records, team_meta(draft_records), False
    live_error = live_hint = target = ""
    if USE_ESPN:
        target = ESPN_URL_OVERRIDE or (
            f"https://lm-api-reads.fantasy.espn.com/apis/v3/games/flb/seasons/{ESPN_SEASON}"
            f"/segments/0/leagues/{ESPN_LEAGUE_ID}?view=mTeam&view=mRoster")
        try:
            import urllib.error
            from espn_live import fetch_league, current_rosters
            raw = Path(ESPN_JSON_FILE).read_text(encoding="utf-8") if ESPN_JSON_FILE else None
            league = fetch_league(ESPN_LEAGUE_ID, ESPN_SEASON, ESPN_S2, ESPN_SWID,
                                  local_json=raw, url=(ESPN_URL_OVERRIDE or None))
            rosters = current_rosters(league, OWNER_ALIAS)
            records, teams, matched = build_from_espn(rosters, index, OWNER_TEAM)
            live = True
            counts = {t: sum(1 for r in records if r["tag"] == t) for t in ("kept","auction","traded","fa")}
            print(f"  live: {len(records)} current roster spots — {counts['kept']} kept, "
                  f"{counts['auction']} auction, {counts['traded']} traded, {counts['fa']} free agents "
                  f"({matched} matched to sheet prices)")
        except Exception as e:
            live_error = f"{type(e).__name__}: {e}"
            if isinstance(e, urllib.error.HTTPError):
                body = ""
                try: body = e.read().decode("utf-8", "replace")
                except Exception: pass
                msg = ""
                if body:
                    try: msg = (json.loads(body).get("messages") or [""])[0]
                    except Exception: msg = body[:160]
                live_error = f"HTTP {e.code} from ESPN" + (f' — "{msg}"' if msg else "")
                if e.code in (401, 403):
                    live_hint = ("ESPN refused the request. If the league is private, the relay's espn_s2 "
                                 "cookie is missing or expired — refresh it in the Cloudflare Worker and Deploy. "
                                 "If you switched the league to public, double-check that setting saved.")
            elif isinstance(e, urllib.error.URLError):
                live_hint = "Couldn't reach the URL below — check the relay address is right and deployed."
            print(f"  live: ESPN unavailable ({live_error}) — using draft-day rosters from the sheet")

    # 2027 keeper cost: FA (never auctioned) = flat $5; otherwise the price-grows-by-1
    # escalation (2*2026 - 2025 + 1), or a freshly-auctioned player's first keep = 2026 + 2.
    for r in records:
        if r.get("tag") == "fa":
            r["p"][2027] = 5
        else:
            p26, p25 = r["p"].get(2026), r["p"].get(2025)
            r["p"][2027] = None if p26 is None else (2*p26 - p25 + 1 if p25 is not None else p26 + 2)

    data = {"league": LEAGUE, "season": SEASON, "buyin": BUYIN, "live": live,
            "built_at": built_at, "live_error": live_error, "live_hint": live_hint,
            "live_target": target, "via_relay": bool(ESPN_URL_OVERRIDE),
            "records": records, "teams": sorted(teams.values(), key=lambda x: -x["total2026"])}
    html = TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/", json.dumps(data, ensure_ascii=False))
    OUTPUT.write_text(html, encoding="utf-8")
    src = "ESPN live rosters" if live else "draft-day rosters (ESPN off)"
    print(f"✓ Built {OUTPUT.name} — {len(records)} players across {len(teams)} teams · {src}")

if __name__ == "__main__":
    build()
