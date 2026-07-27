# Fix: route ESPN through a free Cloudflare relay

ESPN blocks GitHub's servers, so the build can't reach it directly. This creates a tiny
free "relay" on Cloudflare that ESPN *will* answer, and points your build at it. All in
the browser, no credit card. Takes about 10 minutes.

You'll do three things: (A) create the relay on Cloudflare, (B) paste its address into
your build, (C) re-upload two files and run it.

---

## A. Create the Cloudflare relay

1. Go to https://dash.cloudflare.com and sign up for a free account (verify your email).
2. In the left sidebar click **Compute (Workers)** (older menus call it **Workers & Pages**).
3. Click **Create** → **Create Worker** (or **Start with Hello World**).
4. It suggests a name like `worker-abc123`. Change it to something clear like
   **`espn-relay`**, then click **Deploy** (a placeholder deploys — that's fine).
5. Click **Edit code** (top right).
6. Select **all** the code in the editor and delete it. Then paste in the entire contents
   of the `espn-relay-worker.js` file I gave you.
7. Click **Deploy** (top right).
8. Now find your relay's address. Go back to the Worker's overview page — under
   **Settings → Domains & Routes** (or right on the overview) you'll see a URL like
   **`https://espn-relay.kokanjohn.workers.dev`**. Copy it.

**Test it:** paste that URL into a new browser tab. You should see a wall of JSON starting
with `{"draftDetail"...` — the same kind you saw earlier. If you do, the relay works. 🎉
(If you get an error instead, stop here and send me what it says.)

---

## B. Point your build at the relay

1. Go to your GitHub repo → click the **`build_board.py`** file → click the pencil (**Edit**).
2. Near the top, find this line:

   `ESPN_URL_OVERRIDE = ""           # if ESPN blocks GitHub, paste your Cloudflare relay URL here`

3. Paste your relay URL between the quotes, so it reads (with your actual address):

   `ESPN_URL_OVERRIDE = "https://espn-relay.kokanjohn.workers.dev"`

4. Click **Commit changes** → **Commit changes**.

---

## C. Update the two changed files, then run

I updated `build_board.py` and `espn_live.py` so they know how to use the relay. If you
did step B by editing on GitHub, `build_board.py` is already current — you only need to
replace **`espn_live.py`**:

1. Repo → **Add file → Upload files** → drag in the new **`espn_live.py`** → **Commit changes**.
   (Uploading a file with the same name replaces the old one — that's what we want.)
2. Go to the **Actions** tab → **Build & publish keeper board** → **Run workflow** → **Run workflow**.
3. Wait for the green check, then hard-refresh your site.

**Check the line under the "108 Stitches" title:** it should now read **"live from ESPN
rosters."** Your team, Noble Tiger, should no longer show players you've dropped.

---

## Good to know
- The relay only forwards your league's roster data — nothing private, no login.
- Cloudflare's free plan allows 100,000 relay hits a day; your board uses a handful. No cost.
- If you ever change leagues or seasons, update the `target` URL inside the Worker code
  (the league number `35759` and season `2026`) and redeploy.
