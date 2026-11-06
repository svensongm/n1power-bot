# N1 POWER bot

Runs by itself on GitHub every day and posts to Telegram (@N1POWERbot):

| When (CET) | What |
|---|---|
| first auction result (~09:35) | **CBC album**: CBC SEE and CBC CH pictures. Under it: the delivery day, CBC SEE x/34, CBC CH x/8 and when the prices follow. New borders update the same pictures; a short reply lists what came in. |
| first price (~10:30, ME) | **Full album** (one message to forward): CBC SEE, CBC CH, DA prices table and the two calendar pictures, with the counts – CBC SEE, CBC CH, DA prices x/25, Calendar x/14. The morning CBC album is deleted. |
| every new result | The full album is updated within about 30 seconds and a short reply lists exactly what came in, e.g. "CBC SEE: HU-RS", "UA price", "Spreads: GRHU". If a source corrects a value that was already posted, the pictures are redrawn and the reply says e.g. "Corrected: DE price". |
| everything in (latest 17:00) | Caption shows ✅ complete (or final + what is n/a), then the month's Excel file `N1POWER_YYYY-MM.xlsx`. |

The Excel file has two sheets that grow every day and look like the pictures: **Prices** (calendar cards with flags,
B / P / MTD per day, week column with the spread to HU, the week recap, then the hourly prices of every day, newest
first) and **CBC** (daily average per border, then the hourly auction prices of every day, newest first).

How it runs: GitHub starts the bot every 20 minutes (07:07–16:47 CET); each run watches for 48 minutes, so runs overlap.
Missing data is checked every 30 seconds while something is due – CBC from 09:30, and each market from 5 minutes before
its usual time (ME 10:30, MK 10:45, CH and RS 11:10, Core and all others 12:55, set in `config.toml`). In between it sleeps.
Everything still missing at 17:00 is shown as n/a.
Borders are written with a dash: HU-RS = from HU to RS.

Calendars: blue = spread to HU (HUxx = HU-xx; DEHU, GRHU, ATHU = country-HU) – the chip in each card header
(month to date), a line in every day, the week column and the week recap. "MTD vs previous day" = how much the
month-to-date base average moved with the newest day.

Data: JAO (capacity auction prices, when the `JAO_API_KEY` secret is set), ENTSO-E (all other capacity auctions,
and prices that Volue does not have) and Volue (prices; IT = PUN). DE and AT are always the EPEX day-ahead prices
(SDAC 12:00 auction, ~12:55), never the EXAA auction.

---

## Set-up (once, about 10 minutes)

1. **Telegram** – open `t.me/N1POWERbot` and press **Start**.
2. **GitHub** – sign up for free at github.com → **+** → **New repository** → name `n1power-bot` → **Public** → **Create repository**.
3. **Upload** – click *uploading an existing file* → drag in everything from the `n1power-bot` folder, **including the `.github` folder** → **Commit changes**.
   Check that `.github/workflows/n1power.yml` is there. If it is not: **Add file → Create new file**, type the name
   `.github/workflows/n1power.yml`, paste the content of that file, **Commit changes**.
4. **Keys** – **Settings → Secrets and variables → Actions → New repository secret**, one at a time:
   `ENTSOE_API_KEY`, `VOLUE_CLIENT_ID`, `VOLUE_CLIENT_SECRET`, `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID` (step 5),
   and `JAO_API_KEY` (optional – the JAO Market Data token; without it the bot uses ENTSO-E only).
5. **Chat id** – **Actions** tab → **N1 POWER bot** → **Run workflow** → choose *show my Telegram chat id* → **Run workflow**.
   The bot sends you your chat id in Telegram → save it as the secret `TELEGRAM_CHAT_ID`.
   More than one chat (for example a group): separate the ids with commas. A group id starts with `-`.
6. **Test** – **Run workflow** → *send report now* → within about 2 minutes everything for tomorrow arrives in Telegram.

7. **JAO check** (only with `JAO_API_KEY`) – **Run workflow** → *test JAO connection* → the log shows which borders
   JAO delivers and example prices. Nothing is sent to Telegram.

That's it: from now on it runs by itself every day.

## Changing things
* `config.toml` (pencil icon on GitHub): markets, borders, times, calendar countries, spread list, Excel on/off.
* Pause the bot: **Actions → N1 POWER bot → ··· → Disable workflow**.

## If something looks wrong
* `⚠️ N1 POWER bot problem: ENTSO-E refused the API key` → check the `ENTSOE_API_KEY` secret.
* If the Volue login fails, the bot uses ENTSO-E (without PUN) and adds a note under the table.
* If JAO refuses the token or is down, the bot quietly uses ENTSO-E; the CBC footer shows the source, e.g.
  "source JAO 9 + ENTSO-E 21". Check the `JAO_API_KEY` secret and run *test JAO connection*.
* **n/a** in a column → neither Volue nor ENTSO-E has data (e.g. no auction on that border that day).
  If a Volue curve has a special name, add it under `volue_curves` in `config.toml`.

Free: public repositories get unlimited GitHub Actions minutes. Keys stay in GitHub's encrypted secrets and are
never shown in the code or in the logs.
