# Stock Tracker – Amazon, Flipkart & any shopping site

Give it product links. Every few minutes it checks each page and records the exact
**date and time** a product comes **IN STOCK**, goes **OUT OF STOCK**, and comes **BACK IN STOCK**.
Every change is saved to Excel, CSV and PDF, and you get an alert on Telegram, WhatsApp,
Google Chat, Email, ntfy (phone push), Discord or Slack.

## What you get

| File (in `data/`) | What's inside |
|---|---|
| `stock_report.xlsx` | 4 sheets: **Current Status**, **Stock Events** (every in/out change with date & time), **Availability Periods** (came in stock → went out, and how long it stayed), **All Checks** |
| `stock_report.pdf` | Same report, printable |
| `events.csv` / `checks.csv` | Raw history (opens in Excel / Google Sheets) |

Example alert:
```
🟢 BACK IN STOCK
iPhone 15 128GB Black
Site: Flipkart
Status: IN STOCK
Time: 17 Oct 2026, 02:41:09 PM
Price: ₹58,999
Previous status: OUT OF STOCK since 2026-10-15 09:12:44
https://www.flipkart.com/...
```

## Setup on Windows (10 minutes)

1. Install **Python 3.10+** from python.org – tick **"Add Python to PATH"** during install.
2. Unzip this folder, double-click **`1_INSTALL.bat`**.
3. Double-click **`2_START.bat`**, press Enter (or type a port number). The dashboard opens in your browser.
4. Click **Turn on pop-up alerts** and allow notifications.
5. Paste product links into the box at the top and click **Start watching**.
6. Optional: turn on Telegram / WhatsApp / Google Chat / Email in **`config.yaml`** (see below),
   then click **Send a test alert** on the dashboard.

Keep the black window open (you can minimise it) – tracking runs as long as it's open.

Mac/Linux: `pip install -r requirements.txt && playwright install chromium`, copy `config.example.yaml` to `config.yaml`, then `python stock_tracker.py web`.

## The dashboard

- Every product is a box on a shelf. **Lit in saffron = in stock.** Dimmed = out of stock.
- The box **flips over** the moment its stock changes, with a sound, an on-screen message,
  a pop-up on your PC, and a message on every channel you've turned on.
- You're alerted for all three moments: **available**, **went out of stock**, **back in stock**
  (and a price drop if you set "Alert below ₹").
- "What changed" shows every change with exact date and time. Excel / PDF / CSV buttons download the full history.
- The top pill shows the countdown to the next check, or live progress while checking.
- Updates appear live – no page reloads. Different shopping sites are checked at the same time, so a full check takes seconds.

## Ports

- `2_START.bat` asks for a port; press Enter for 8787. **If that port is busy, it picks the next free one automatically**
  and opens it in your browser.
- Or: `python stock_tracker.py web --port 9123`. To make a port permanent, set `web_port:` in `config.yaml`.
- `4_RUN_WITHOUT_DASHBOARD.bat` tracks and alerts with **no port at all**.

Find out which ports are busy:
- Double-click **`5_CHECK_PORTS.bat`**, or run `python stock_tracker.py ports --from 8000 --to 9000`
- Windows: `netstat -ano | findstr LISTENING` → last column is the program ID → `tasklist /FI "PID eq 1234"` shows its name
- Mac/Linux: `lsof -iTCP -sTCP:LISTEN -P -n`

To open the dashboard on your phone (same Wi-Fi), set `web_host: 0.0.0.0` and open `http://<your-PC-IP>:<port>`.
(Browser pop-ups only work on the PC itself; your phone gets alerts through Telegram/WhatsApp/ntfy.)

## Commands

```
python stock_tracker.py add "https://www.amazon.in/dp/XXXX" --name "Boat headphones" --target-price 1499
python stock_tracker.py list          # show all products and their status
python stock_tracker.py remove 2      # remove product number 2
python stock_tracker.py once          # check everything once now
python stock_tracker.py web           # dashboard + tracking
python stock_tracker.py run           # tracking only, no dashboard
python stock_tracker.py report        # rebuild Excel + PDF
python stock_tracker.py summary       # send summary + reports now
python stock_tracker.py test-notify
```
You can add products while it's running – it re-reads `config.yaml` before every round.

## Notification setup

**Telegram (recommended – free, instant)**
1. In Telegram open **@BotFather** → `/newbot` → pick a name → copy the **bot token**.
2. Open your new bot and press **Start**, send it "hi".
3. Open `https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates` in a browser and copy the number after `"chat":{"id":` → that's your **chat_id**.

**WhatsApp – free (CallMeBot)**
1. Go to callmebot.com → WhatsApp API page, save their number in your phone and send the activation message shown there.
2. You'll get an **apikey** back on WhatsApp. Put your number (with +91) and apikey in config.
   *(For business use, choose `provider: twilio` and fill in your Twilio details.)*

**Google Chat**
In a Google Chat space → space name → **Apps & integrations → Webhooks → Add** → copy the URL.
(Needs a Google Workspace account; personal Gmail accounts can't create webhooks.)

**Email (Gmail)**
Turn on 2-Step Verification, then create an **App Password** at myaccount.google.com → Security → App passwords. Use that 16-letter password in config. Alerts come with the Excel and PDF attached.

**ntfy (free phone push, no signup)** – install the ntfy app, subscribe to a unique topic name, put the same name in config.

## How it decides "in stock"

1. **Amazon** – the availability box, "Currently unavailable", and the Add to Cart / Buy Now buttons.
2. **Flipkart** – "Sold Out", "Currently Unavailable", "Coming Soon" vs. Add to Cart / Buy Now.
3. **Any other site** (Myntra, Croma, Reliance Digital, Ajio, Nykaa, brand stores…) – the product data most shops publish for Google (`schema.org availability`), then button text.
4. If a site uses unusual words, add per product: `in_stock_text: "Add to Bag"` or `out_of_stock_text: "Notify Me"`.

If a page can't be read (site shows a captcha), the check counts as **UNKNOWN** and never triggers a false "out of stock" alert; after 6 failures in a row you get a warning.

## Tips

- Keep `check_every_minutes` at 10–15. Checking every minute gets you blocked by Amazon/Flipkart.
- For limited "flash sale" drops, 5 minutes is the practical minimum.
- To run it 24×7 without your PC, put the folder on any always-on machine (old laptop, Raspberry Pi, or a cheap cloud VM) and run `python stock_tracker.py run`.
- To start automatically with Windows: press Win+R → `shell:startup` → put a shortcut to `2_START.bat` there.
- Use it for personal tracking at a sensible rate; shopping sites' terms restrict heavy automated access.
