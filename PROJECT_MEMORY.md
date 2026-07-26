# Project Memory & Account-Migration Runbook

> **Purpose.** Everything Claude needs to know about this project that is *not* obvious
> from reading the code, plus the exact steps to reconstitute the project on a **new
> Claude (Teams) account without losing any work.** This file is committed to Git, so it
> travels with the repo and is account-agnostic. If a fact about how this project should
> behave lives only in a Claude account's private "memory" store, it belongs here instead.

---

## 0. TL;DR — Why nothing is at risk

**All durable work lives in Git, not in a Claude account.** Switching Claude accounts does
not touch the repository. The learnings, rules, and history are in these committed files:

- `CLAUDE.md` — the 24 project Rules (scorecard, DCF, bank/EVS regimes, auto-push, deferral discipline). This is the single most important artifact.
- `fixes_log.md` — running history of fixes and decisions.
- `handoff/HANDOFF.md`, `handoff/MUFF_DESIGN_NOTES.md` — design + handoff notes.
- `pending_reruns.json` — FMP-blocked deferral ledger.
- `PROJECT_MEMORY.md` — this file.

The only things that are account-side (and must be re-established manually on the new
account) are: the GitHub connection, the Claude Code web **environment** config, the
**secrets**, and the **MCP connectors**. None of those are code, and none are lost — they
just have to be re-entered. See §2.

---

## 1. What lives where

| Asset | Location | Moves with repo? | Action on new account |
|---|---|---|---|
| Rules / learnings (`CLAUDE.md`, `fixes_log.md`) | Git | ✅ automatically | none |
| Engine, bridge, server, templates | Git | ✅ automatically | none |
| Deferral ledger (`pending_reruns.json`) | Git | ✅ automatically | none |
| Render deployment | Render account (deploys from `origin/main`) | independent of Claude | none — keeps running |
| GitHub app authorization | Claude account | ❌ | re-authorize, grant repo access |
| Web environment (network policy, setup) | Claude account | ❌ | recreate |
| Secrets (API keys, tokens, passwords) | Render + Claude env (git-ignored) | ❌ (by design) | re-enter |
| MCP connectors (Gmail, IBKR, GitHub) | Claude account | ❌ | reconnect |
| Chat / session history | Claude account | ❌ | export beforehand if wanted |

---

## 2. Migration runbook (personal repo stays under `jaysang2908`)

**Decision on record:** the repo stays `jaysang2908/Investment-Automation`. The new Teams
Claude account is simply granted access to it. Render is untouched — it deploys from
`origin/main` and does not know or care which Claude account is used, so there is **zero
deployment downtime**.

### Step-by-step

1. **On the old account** — confirm the tree is clean and pushed:
   `git status` (clean) and `git rev-list --left-right --count origin/main...HEAD` → `0 0`.
2. **Capture account memory** — make sure anything you told Claude to "remember" is written
   into this file (§4) and pushed. After this, the learnings are 100% in Git.
3. **Record secrets** — copy the values (not into Git) from Render's dashboard for:
   `FMP_API_KEY`, `GEMINI_KEY`, `GITHUB_TOKEN`, `APP_PASSWORD`, `GITHUB_REPO`,
   and `RENDER_INTERNAL_URL`.
4. **New Teams account → GitHub** — authorize the Claude GitHub app and grant it access to
   `jaysang2908/Investment-Automation`.
5. **New Teams account → web environment** — create a Claude Code on the web environment
   pointed at the repo. Use the same network policy the old environment used, keep the
   committed `.devcontainer/`, and re-enter the secrets from step 3 as environment
   variables.
6. **New Teams account → MCP connectors** — reconnect: **GitHub**, **Gmail**, **IBKR**
   (Interactive Brokers), and any FMP/data connectors used interactively.
7. **Verify** — open a session on the new account, confirm `CLAUDE.md` loads (Rules appear),
   run a read-only smoke check (e.g. `python _score_audit.py`), and confirm the GitHub tool
   can see the repo.

### Secrets & env vars (names only — values are NOT in Git)

From `render.yaml`:

| Var | Used by | Notes |
|---|---|---|
| `FMP_API_KEY` | web, daily-news cron | FMP data — free tier, quota-limited (Rule 24) |
| `GEMINI_KEY` | web | qualitative commentary only (Rule 3) |
| `APP_PASSWORD` | web, daily-long-term-bets cron | site auth |
| `GITHUB_TOKEN` | web, news/price crons | pushes refreshed `outputs.csv` back to the repo |
| `GITHUB_REPO` | web, crons | target repo path |
| `RENDER_INTERNAL_URL` | daily-long-term-bets cron | `https://investment-automation.onrender.com` |

### Render services that keep running (no Claude dependency)

- `investment-automation` (web) — `gunicorn server:app`
- `daily-news-refresh` (cron `0 6 * * *`) — `python daily_news.py`
- `daily-price-refresh` (cron `30 21 * * 1-5`) — `python daily_prices.py` (post-close, pushes `outputs.csv`)
- `daily-long-term-bets` (cron `0 2 * * *`) — `python daily_long_term_bets.py`

---

## 3. Operating context Claude should carry across accounts

- **The user only ever views the live Render URL** (auto-deployed from `origin/main`). An
  edit that is coded but not pushed does not exist from their point of view. This is why
  Rule 23 (auto-push by default) exists — reproduced from `CLAUDE.md`.
- **Audience is professional investors.** Institutional-standard accuracy; never dumb it
  down. (See CLAUDE.md "End User & Mindset".)
- **FMP free tier is quota-limited.** Ancillary features run on cron, not live. Deferred
  re-runs are logged in `pending_reruns.json` (Rule 24), never left in silent limbo.
- **The 24 Rules in `CLAUDE.md` are authoritative** and override default behavior. Read them
  first in any new session.

---

## 4. Account-side memory to carry over

> This section is the portable home for anything currently held only in a Claude account's
> private memory store. That store cannot be exported programmatically, so paste such items
> here (then push) before switching accounts. Add dated entries; keep durable rules in
> `CLAUDE.md` and fix history in `fixes_log.md`.

_(none captured yet — add items below as `- YYYY-MM-DD: …`)_
