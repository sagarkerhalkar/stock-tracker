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
3. Open **`config.yaml`** in Notepad:
   - Under `products:` paste your product links (delete the two examples).
   - Turn on the notification channels you want (`enabled: true`) – see below.
4. Double-click **`3_TEST_NOTIFICATIONS.bat`** – you should receive a test message.
5. Double-click **`2_START_TRACKING.bat`** and keep that window open. Done.

Mac/Linux: `pip install -r requirements.txt && playwright install chromium`, copy `config.example.yaml` to `config.yaml`, then `python stock_tracker.py run`.

## Web dashboard & ports

`2_START_TRACKING.bat` uses **no port at all**. A port is only needed for the optional
web dashboard (add/remove products, see status, download Excel/PDF in your browser):

- Double-click **`4_START_DASHBOARD.bat`**, type a port (or press Enter for 8787), open the link it prints.
- Or: `python stock_tracker.py web --port 9123`
- **If the port is busy, it automatically picks the next free one** and prints the new link.
- To make a port permanent, set `web_port:` in `config.yaml`.

Find out which ports are busy:
- Double-click **`5_CHECK_PORTS.bat`**, or run `python stock_tracker.py ports --from 8000 --to 9000`
- Windows: `netstat -ano | findstr LISTENING` → last column is the program ID → `tasklist /FI "PID eq 1234"` shows its name
- Mac/Linux: `lsof -iTCP -sTCP:LISTEN -P -n`

To open the dashboard on your phone (same Wi-Fi), set `web_host: 0.0.0.0` and open `http://<your-PC-IP>:<port>`.

## Commands

```
python stock_tracker.py add "https://www.amazon.in/dp/XXXX" --name "Boat headphones" --target-price 1499
python stock_tracker.py list          # show all products and their status
python stock_tracker.py remove 2      # remove product number 2
python stock_tracker.py once          # check everything once now
python stock_tracker.py run           # keep checking forever
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
- To start automatically with Windows: press Win+R → `shell:startup` → put a shortcut to `2_START_TRACKING.bat` there.
- Use it for personal tracking at a sensible rate; shopping sites' terms restrict heavy automated access.
