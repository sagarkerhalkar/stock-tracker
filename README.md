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
3. Double-click **`2_START.bat`** and press Enter. The dashboard opens in your browser.

Then:
- Paste product links at the top and press **Start watching**.
- Open **Alerts** in the left menu and set up Telegram / WhatsApp / Google Chat / email (step-by-step instructions are on that page, each with a **Send test** button).
- Press **Turn on browser alerts** for a sound + pop-up while the dashboard is open.

Keep the black window open (minimise it). Tracking stops when it's closed or the PC sleeps.

Mac/Linux: `pip install -r requirements.txt && python -m playwright install chromium`, then `python stock_tracker.py web`.

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

## Reports

`data/stock_report.xlsx` – sheets: Current Status, Stock Events (every change with date & time),
Availability Periods (came in → went out → how long), All Checks.
`data/stock_report.pdf` – same, printable. `data/events.csv` – raw history.
Download them from the **Reports** page too. A daily summary is sent at 9:00 AM (change `daily_summary_time` in `config.yaml`).

## Ports

- `2_START.bat` asks for a port; Enter = 8787. **If it's busy, the next free port is used automatically** and opened in your browser.
- Or: `python stock_tracker.py web --port 9123`. To keep a port, set `web_port:` in `config.yaml`.
- `4_RUN_WITHOUT_DASHBOARD.bat` tracks and alerts with no port at all.
- See which ports are busy: **`5_CHECK_PORTS.bat`**, or `netstat -ano | findstr LISTENING` then `tasklist /FI "PID eq <number>"`.
- Phone on the same Wi-Fi: set `web_host: 0.0.0.0` and open `http://<PC-IP>:<port>`.

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
