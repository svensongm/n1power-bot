# N1 POWER bot

Runs by itself on GitHub every day and posts to Telegram (@N1POWERbot):

| When (CET) | What |
|---|---|
| from ~09:35 | **CBC daily auction prices for tomorrow** – two pictures, **SEE first, then CH**: hourly tables in the desk price-sheet layout. They fill in as more borders are published; when all are known (or by 12:30) they are marked final and a short "CBC final" message follows. |
| from ~13:00 | **Day-ahead prices for tomorrow – three pictures**: (1) the hourly table for 25 markets; (2) calendars HU, HR, SI, RS, RO, BG, GR, ME, MK; (3) calendars IT, DE, AT, FR, ES and the week recap of spreads to HU. Each day shows **B** (base), **P** (peak 08–20) and **MTD** (month average of base up to that day); the week column shows the Mon–Sun B and P and, in small blue, the spread to HU (HUxx = HU − xx; DEHU, GRHU, ATHU = country − HU). Posted when HU is out; the pictures fill in until all markets are known. |
| when prices are final | **The month's Excel file** (`N1POWER_YYYY-MM.xlsx`), two sheets that grow every day: **Prices** (calendars + hourly prices of every day, newest first) and **CBC** (daily average per border + hourly auction prices of every day, newest first). |

Data: Volue (prices; IT = PUN) and ENTSO-E (capacity auctions, and prices that Volue does not have).

---

## Set-up (once, about 10 minutes)

1. **Telegram** – open `t.me/N1POWERbot` and press **Start**.
2. **GitHub** – sign up for free at github.com → **+** → **New repository** → name `n1power-bot` → **Public** → **Create repository**.
3. **Upload** – click *uploading an existing file* → drag in everything from the `n1power-bot` folder, **including the `.github` folder** → **Commit changes**.
   Check that `.github/workflows/n1power.yml` is there. If it is not: **Add file → Create new file**, type the name
   `.github/workflows/n1power.yml`, paste the content of that file, **Commit changes**.
4. **Keys** – **Settings → Secrets and variables → Actions → New repository secret**, one at a time:
   `ENTSOE_API_KEY`, `VOLUE_CLIENT_ID`, `VOLUE_CLIENT_SECRET`, `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID` (step 5).
5. **Chat id** – **Actions** tab → **N1 POWER bot** → **Run workflow** → choose *show my Telegram chat id* → **Run workflow**.
   The bot sends you your chat id in Telegram → save it as the secret `TELEGRAM_CHAT_ID`.
   More than one chat (for example a group): separate the ids with commas. A group id starts with `-`.
6. **Test** – **Run workflow** → *send report now* → within about 2 minutes everything for tomorrow arrives in Telegram.

That's it: from now on it runs by itself every day.

## Changing things
* `config.toml` (pencil icon on GitHub): markets, borders, times, calendar countries, spread list, Excel on/off.
* Pause the bot: **Actions → N1 POWER bot → ··· → Disable workflow**.

## If something looks wrong
* `⚠️ N1 POWER bot problem: ENTSO-E refused the API key` → check the `ENTSOE_API_KEY` secret.
* If the Volue login fails, the bot uses ENTSO-E (without PUN) and adds a note under the table.
* **n/a** in a column → neither Volue nor ENTSO-E has data (e.g. no auction on that border that day).
  If a Volue curve has a special name, add it under `volue_curves` in `config.toml`.

Free: public repositories get unlimited GitHub Actions minutes. Keys stay in GitHub's encrypted secrets and are
never shown in the code or in the logs.
