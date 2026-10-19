# Restock Watch

Watches products on **Amazon, Flipkart, Myntra, Croma, Reliance Digital, AJIO, Nykaa** and most other
shopping sites, and alerts you on **Telegram, WhatsApp, Google Chat, email and a PC pop-up** the moment a product:

- ✅ **becomes available**
- 🔴 **goes out of stock**
- 🟢 **comes back in stock** (with how long it was gone)
- 💰 drops to your target price (optional)

Every change is saved with its exact date and time in **Excel, PDF and CSV**.

## Start in 3 steps (Windows)

1. Install **Python 3.10+** from python.org – tick **"Add Python to PATH"** during install.
2. Double-click **`1_INSTALL.bat`** (once).
3. Double-click **`2_START.bat`**. The dashboard opens at **http://localhost:8787** – always the same address, so bookmark it.

**It starts by itself with Windows.** From now on Restock Watch starts automatically (minimised, with a small
"Restock Watch is running" pop-up) every time you turn on / restart your PC and sign in – no need to click anything.
If it ever stops unexpectedly it restarts itself within 10 seconds. If port 8787 is briefly busy at start-up it waits
and tries again every minute. Turn this off on the **Alerts** page ("Start automatically with Windows") or with
`8_START_WITH_WINDOWS.bat`.

Then:
- Paste product links at the top and press **Start watching**.
- Open **Alerts** in the left menu and set up Telegram / WhatsApp / Google Chat / email (step-by-step instructions are on that page, each with a **Send test** button).
- Press **Turn on browser alerts** for a sound + pop-up while the dashboard is open.

Keep the black window open (minimise it). Tracking stops when it's closed or the PC sleeps.

Mac/Linux: `pip install -r requirements.txt && python -m playwright install chromium`, then `python stock_tracker.py web`.

## Updating to a new version (your data is kept)

Your products, alert settings and full history live in **`config.yaml`** and the **`data`** folder.
Updates never touch them.

**Easiest:** double-click **`6_UPDATE.bat`**. It
1. copies `config.yaml` and `data` into `backup\<date>` first,
2. downloads the latest version from GitHub (program files only),
3. installs anything new.
Then double-click `2_START.bat` – it closes the old version by itself.

**Or by hand:** unzip the update zip *into the same folder* and choose "Replace". The update zip contains
no `config.yaml` and no `data` folder, so nothing of yours can be overwritten.

**If you cloned with git:** `git pull` (your `config.yaml` and `data` are never uploaded or replaced).

## Products or history missing? Bring back old data

If you ever unzip a new version into a *different* folder, that folder starts empty – your old folder still has
everything. To bring it over:

- The dashboard offers it by itself when it starts empty ("Your old products and history are still on this PC"), or
- **Reports → Find my old data**, or
- double-click **`7_RESTORE_OLD_DATA.bat`**, or run `python stock_tracker.py restore "C:\path\to\old\folder"`.

It searches Downloads, Desktop, Documents and OneDrive, and **merges**: the other folder's products, alert settings and
full history are added; nothing here is deleted; a backup goes to `backup\` first; the old folder isn't changed.

**Run only one copy.** Two copies from two folders each keep their own products and history. If that happens the
dashboard shows "Two copies of Restock Watch are running" with both folder paths.

## How fast will I know?

| Product is… | Checked every (default) | Why |
|---|---|---|
| Out of stock | **2 minutes** | so you hear the moment it's back |
| Available | **5 minutes** | so you hear when it sells out |

Change both on the **Alerts** page. Alerts are sent the instant a change is detected, before anything else.
Going below 1–2 minutes risks Amazon/Flipkart temporarily blocking your connection. No tool can see a restock
before the shop's page shows it – this checks the page repeatedly, like refreshing it for you.

## What alerts look like

**Telegram** – product photo, price, time and an **Open on Amazon** button.
**Google Chat** – a card with photo, price, time and an open button.
**WhatsApp** – same details as text with the product link.

```
🟢 BACK IN STOCK - available again!

Sony WH-1000XM5 Wireless Headphones
💰 ₹26,990
🛒 Flipkart
🕒 17 Oct 2026, 02:41:28 PM
Was out of stock for 3h 12m

https://www.flipkart.com/...
```

## Setting up each app (also shown inside the app)

**Telegram (recommended – free, instant, most reliable)**
1. In Telegram open **@BotFather**, send `/newbot`, pick any name → copy the **token**.
2. Paste it on the Alerts page.
3. Open your new bot, press **Start**, send "hi".
4. Press **Find my chat ID** → **Save** → **Send test**.

**WhatsApp (free via CallMeBot, to your own number)**
1. Open CallMeBot's WhatsApp API page (callmebot.com) and save their number in your phone.
2. Send them: `I allow callmebot to send me messages`
3. Paste the API key they reply with, plus your number (+91…) → **Save** → **Send test**.
CallMeBot is a free personal service and can occasionally be slow – keep Telegram on too.
For an official business WhatsApp, use Twilio (`provider: twilio` in `config.yaml`).

**Google Chat**
Space → space name → **Apps & integrations → Webhooks → Add webhook** → copy URL → paste → **Save** → **Send test**.
Needs a Google Workspace (work/school) account; personal @gmail.com accounts can't add webhooks.

**Email (Gmail)** – turn on 2-Step Verification, create an **App password** at myaccount.google.com/apppasswords,
enter your Gmail + that 16-letter password. Alerts arrive with the Excel and PDF attached.

**PC pop-up** – on by default. **ntfy** – free phone push app; pick a topic name only you know.

## Reports – item wise

Every report is organised **product by product**, so you can read one item's story at a glance:

- **Now:** available or out of stock, and since when
- **Came back in stock X times, went out Y times**, total time available, % of time available
- A **green/red timeline bar** (green = available, red = out of stock)
- A table of every period, latest first: **Status · From · To · How long it lasted · Price**

Where to find it:
- **Reports** page in the dashboard (pick one product or see all)
- **Excel** (`data/stock_report.xlsx`): *Summary* (one line per product), *Item-wise history* (one block per product), *All checks (raw)*
- **PDF** (`data/stock_report.pdf`): an at-a-glance table, then one section per product
- `data/events.csv`: the raw change log

**Removed products:** when you remove a product, its history is **hidden** from the dashboard, reports, Excel and PDF.
It is not deleted – tick **Show removed products** on the Reports page to see it again, download it, or press **Watch again**.

**Download one product (or a few):** on the Reports page each product has **Excel · PDF · CSV** buttons, or tick several
products and use **Download** in the bar that appears. On the Products page: ••• → **Download history**.
A daily summary is sent at 9:00 AM (change `daily_summary_time` in `config.yaml`).

## Port 8787

Restock Watch always runs on **http://localhost:8787**.
- If an older Restock Watch is already using 8787, it is closed automatically and the new one takes over (its data stays on disk).
- If a different program is using 8787, the black window tells you which one; close it (or restart the PC) and start again.
  `5_CHECK_PORTS.bat` shows what is using 8787. To use another port permanently, change `web_port:` in `config.yaml`.
- `4_RUN_WITHOUT_DASHBOARD.bat` tracks and alerts with no port at all.
- Phone on the same Wi-Fi: set `web_host: 0.0.0.0` in `config.yaml` and open `http://<PC-IP>:8787` (anyone on that Wi-Fi can then open it).

## Commands

```
python stock_tracker.py web [--port 9123]   # dashboard + tracking
python stock_tracker.py run                 # tracking only
python stock_tracker.py add "<link>" --name "Boat headphones" --target-price 1499
python stock_tracker.py list | remove 2 | once | report | summary | test-notify
python stock_tracker.py ports --from 8000 --to 9000
```

## Good to know

- If a site shows a captcha, the app opens a hidden Chrome window to read the page. If it still can't, that check is
  marked "Can't read page" – it never sends a false "out of stock". After 6 blocked checks in a row you get a warning.
- Unusual site wording? Add to that product in `config.yaml`: `in_stock_text: "Add to Bag"` or `out_of_stock_text: "Notify Me"`.
- Start with Windows: Win+R → `shell:startup` → put a shortcut to `2_START.bat` there.
- Your tokens and passwords live only in `config.yaml` on your PC; it is excluded from GitHub by `.gitignore`.
- For personal use at a sensible rate – shopping sites' terms restrict heavy automated access.
