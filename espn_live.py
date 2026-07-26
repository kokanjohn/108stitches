#!/usr/bin/env python3
"""
espn_live.py — read current rosters from ESPN's (public) fantasy baseball API.

Returns each team's CURRENT roster with ESPN's acquisitionType per player, which
build_board.py turns into Kept / Auction / Traded / Free Agency tags. No auth is
needed for a public league; pass espn_s2 + swid for a private one. The API is
undocumented, so callers wrap this in try/except and fall back to the static board.
"""
import json, re, unicodedata, urllib.request, difflib

HOST = "https://lm-api-reads.fantasy.espn.com"

# best-effort ESPN id maps (used only for players NOT in the spreadsheet)
PRO_TEAM = {0:"FA",1:"BAL",2:"BOS",3:"LAA",4:"CHW",5:"CLE",6:"DET",7:"KC",8:"MIL",
 9:"MIN",10:"NYY",11:"ATH",12:"SEA",13:"TEX",14:"TOR",15:"ATL",16:"CHC",17:"CIN",
 18:"HOU",19:"LAD",20:"WSH",21:"NYM",22:"PHI",23:"PIT",24:"STL",25:"SD",26:"SF",
 27:"COL",28:"MIA",29:"ARI",30:"TB"}
SLOT = {0:"C",1:"1B",2:"2B",3:"3B",4:"SS",5:"OF",6:"MI",7:"CI",11:"DH",12:"UT",
 13:"P",14:"SP",15:"RP",16:"BE",17:"IL"}

def key(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii","ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)

def fetch_league(league_id, season, espn_s2=None, swid=None, timeout=25, local_json=None):
    if local_json is not None:
        return json.loads(local_json) if isinstance(local_json, str) else local_json
    url = (f"{HOST}/apis/v3/games/flb/seasons/{season}/segments/0/leagues/{league_id}"
           f"?view=mTeam&view=mRoster")
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://fantasy.espn.com/",
        "X-Fantasy-Source": "kona", "X-Fantasy-Platform": "kona",
    })
    if espn_s2 and swid:
        req.add_header("Cookie", f"espn_s2={espn_s2}; SWID={swid}")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)

def _team_name(t):
    return (t.get("name") or " ".join(x for x in (t.get("location"), t.get("nickname")) if x)).strip()

def current_rosters(league, owner_alias=None):
    """-> list of dicts, one per current roster spot:
       {owner, team, player, player_id, acq, mlb, pos}"""
    owner_alias = owner_alias or {}
    members = {}
    for m in league.get("members", []):
        full = f"{(m.get('firstName') or '').strip()} {(m.get('lastName') or '').strip()}".strip()
        members[m.get("id")] = owner_alias.get(full, full) or (m.get("displayName") or "")
    out = []
    for t in league.get("teams", []):
        owners = t.get("owners") or []
        person = members.get(owners[0]) if owners else _team_name(t)
        for e in (t.get("roster") or {}).get("entries", []):
            pl = (e.get("playerPoolEntry") or {}).get("player") or {}
            name = pl.get("fullName") or ""
            if not name:
                continue
            out.append({
                "owner": person, "team": _team_name(t), "player": name,
                "player_id": pl.get("id"),
                "acq": (e.get("acquisitionType") or "").upper(),
                "mlb": PRO_TEAM.get(pl.get("proTeamId"), ""),
                "pos": SLOT.get(e.get("lineupSlotId"), ""),
            })
    return out

def fuzzy(k, keys, cutoff=0.88):
    near = difflib.get_close_matches(k, keys, n=1, cutoff=cutoff)
    return near[0] if near else None
