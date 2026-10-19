#!/usr/bin/env python3
"""
Stock Tracker - watches Amazon, Flipkart and other shopping sites and tells you
when a product comes IN STOCK or goes OUT OF STOCK, with date and time.

  python stock_tracker.py add <product-url> --name "My product"
  python stock_tracker.py list
  python stock_tracker.py once          # check all products one time
  python stock_tracker.py run           # keep checking forever (every N minutes)
  python stock_tracker.py report        # rebuild Excel + PDF reports
  python stock_tracker.py test-notify   # send a test message to every enabled channel

History is kept in data/  ->  events.csv, checks.csv, stock_report.xlsx, stock_report.pdf
"""
import argparse
import csv
import json
import os
import random
import re
import smtplib
import sys
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote, urlparse
from zoneinfo import ZoneInfo

import requests
import yaml
from bs4 import BeautifulSoup

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE, "config.yaml")
DATA = os.path.join(BASE, "data")
STATE_FILE = os.path.join(DATA, "state.json")
EVENTS_CSV = os.path.join(DATA, "events.csv")
CHECKS_CSV = os.path.join(DATA, "checks.csv")
XLSX_FILE = os.path.join(DATA, "stock_report.xlsx")
PDF_FILE = os.path.join(DATA, "stock_report.pdf")

IN, OUT, UNKNOWN = "IN STOCK", "OUT OF STOCK", "UNKNOWN"
APP_VERSION = 8

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:131.0) Gecko/20100101 Firefox/131.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
]

EVENT_FIELDS = ["date", "time", "product", "site", "event", "old_status", "new_status", "price", "url", "note"]
CHECK_FIELDS = ["date", "time", "product", "site", "status", "price", "detail", "url"]


# --------------------------------------------------------------------------- config / state
def load_config():
    if not os.path.exists(CONFIG_FILE):
        # first run: create config.yaml from the example, with an empty product list
        with open(os.path.join(BASE, "config.example.yaml"), encoding="utf-8") as f:
            text = f.read()
        m = re.search(r"(?m)^products:\s*$", text)
        if m:
            end = text.find("\n# ----", m.end())
            text = text[:m.start()] + "products: []\n" + (text[end:] if end != -1 else "")
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(text)
    with open(CONFIG_FILE, encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    cfg.setdefault("settings", {})
    cfg.setdefault("products", [])
    cfg.setdefault("notifications", {})
    return cfg


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, allow_unicode=True)


def tz(cfg):
    return ZoneInfo(cfg["settings"].get("timezone", "Asia/Kolkata"))


def now(cfg):
    return datetime.now(tz(cfg))


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state):
    os.makedirs(DATA, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def append_csv(path, fields, row):
    os.makedirs(DATA, exist_ok=True)
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        if new:
            w.writeheader()
        w.writerow({k: row.get(k, "") for k in fields})


def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def site_name(url):
    host = urlparse(url).netloc.lower().replace("www.", "")
    for key, name in [("amazon", "Amazon"), ("flipkart", "Flipkart"), ("myntra", "Myntra"),
                      ("croma", "Croma"), ("reliancedigital", "Reliance Digital"), ("ajio", "AJIO"),
                      ("meesho", "Meesho"), ("nykaa", "Nykaa"), ("tatacliq", "Tata CLiQ"),
                      ("vijaysales", "Vijay Sales"), ("snapdeal", "Snapdeal")]:
        if key in host:
            return name
    return host


# --------------------------------------------------------------------------- fetching
def fetch_html(url, settings):
    """Download the product page. Uses a real browser (Playwright) if enabled or if plain request is blocked."""
    use_browser = settings.get("use_browser", "auto")  # auto | always | never
    html, err = None, None
    if use_browser != "always":
        try:
            headers = {
                "User-Agent": random.choice(USER_AGENTS),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-IN,en;q=0.9",
                "Cache-Control": "no-cache",
            }
            r = requests.get(url, headers=headers, timeout=30)
            if r.status_code == 200 and not looks_blocked(r.text):
                return r.text, None
            err = f"HTTP {r.status_code}" + (" (blocked/captcha)" if looks_blocked(r.text) else "")
            html = r.text
        except Exception as e:
            err = str(e)
    if use_browser in ("auto", "always"):
        b_html, b_err = fetch_with_browser(url)
        if b_html:
            return b_html, None
        err = f"{err or ''} | browser: {b_err}".strip(" |")
    return html, err


def fetch_with_browser(url):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None, "playwright not installed (optional: pip install playwright && playwright install chromium)"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(user_agent=random.choice(USER_AGENTS[:2]), locale="en-IN")
            page.goto(url, timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(4000)
            html = page.content()
            browser.close()
            if looks_blocked(html):
                return None, "blocked/captcha"
            return html, None
    except Exception as e:
        return None, str(e)


def looks_blocked(html):
    t = (html or "")[:200000].lower()
    return any(s in t for s in [
        "enter the characters you see below", "api-services-support@amazon.com",
        "to discuss automated access to amazon data", "are you a human", "captcha-delivery",
        "access denied", "request blocked", "please verify you are a human",
    ])


# --------------------------------------------------------------------------- stock detection
SCHEMA_IN = ("instock", "limitedavailability", "onlineonly", "instoreonly", "preorder", "presale", "backorder")
SCHEMA_OUT = ("outofstock", "soldout", "discontinued")

OUT_WORDS = ["currently unavailable", "out of stock", "sold out", "notify me", "coming soon",
             "temporarily unavailable", "not available", "no longer available", "item is unavailable"]
IN_WORDS = ["add to cart", "add to bag", "buy now", "add to basket", "in stock"]


def text_of(el):
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)) if el else ""


def inr(value):
    """'44990' / '₹69,900.00' -> '₹44,990' / '₹69,900' (Indian comma style)."""
    m = re.search(r"[\d,]+(?:\.\d+)?", str(value or ""))
    if not m:
        return ""
    try:
        n = float(m.group(0).replace(",", ""))
    except ValueError:
        return ""
    whole, paise = int(n), round(n - int(n), 2)
    t = str(whole)
    if len(t) > 3:
        head, tail = t[:-3], t[-3:]
        head = ",".join([head[max(0, i - 2):i] for i in range(len(head), 0, -2)][::-1])
        t = head + "," + tail
    return "₹" + t + (f"{paise:.2f}"[1:] if paise else "")


def find_price(soup, html):
    return inr(_find_price(soup, html))


def _find_price(soup, html):
    for sel in ["#corePriceDisplay_desktop_feature_div .a-price .a-offscreen", "#corePrice_feature_div .a-offscreen",
                "#priceblock_dealprice", "#priceblock_ourprice", ".a-price .a-offscreen",
                "div.Nx9bqj", "div._30jeq3", "[itemprop=price]", "meta[property='product:price:amount']"]:
        el = soup.select_one(sel)
        if el:
            val = el.get("content") or text_of(el)
            m = re.search(r"[\d,]+(?:\.\d+)?", val or "")
            if m:
                return "₹" + m.group(0) if "₹" not in val else val.strip()
    m = re.search(r'"price"\s*:\s*"?([\d.,]+)', " ".join(s.string or "" for s in
                                                       soup.find_all("script", type="application/ld+json")))
    if m:
        return "₹" + m.group(1)
    m = re.search(r"₹\s?[\d,]{2,}", soup.get_text(" ", strip=True)[:20000])
    return m.group(0).replace(" ", "") if m else ""


def schema_availability(soup):
    """Read schema.org JSON-LD / microdata availability - works on most shopping sites."""
    found = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(s.string or "")
        except Exception:
            continue
        stack = [data]
        while stack:
            d = stack.pop()
            if isinstance(d, list):
                stack.extend(d)
            elif isinstance(d, dict):
                if "availability" in d:
                    found.append(str(d["availability"]))
                stack.extend(v for v in d.values() if isinstance(v, (dict, list)))
    for el in soup.select("[itemprop=availability]"):
        found.append(el.get("href") or el.get("content") or text_of(el))
    for el in soup.select("meta[property='product:availability'], meta[property='og:availability']"):
        found.append(el.get("content", ""))
    for raw in found:
        v = raw.lower().replace(" ", "").replace("_", "")
        if any(k in v for k in SCHEMA_OUT):
            return OUT, f"schema: {raw}"
        if any(k in v for k in SCHEMA_IN):
            return IN, f"schema: {raw}"
    return None, ""


def check_amazon(soup):
    avail = text_of(soup.select_one("#availability")) or text_of(soup.select_one("#outOfStock"))
    a = avail.lower()
    if any(w in a for w in ["currently unavailable", "out of stock", "unavailable"]):
        return OUT, avail
    if soup.select_one("#add-to-cart-button, #buy-now-button, input[name='submit.add-to-cart']"):
        return IN, avail or "Add to Cart button present"
    if "in stock" in a or "left in stock" in a or "usually dispatched" in a:
        return IN, avail
    if soup.select_one("#outOfStock, #unqualifiedBuyBox"):
        return OUT, avail or "no buy box"
    return None, avail


def check_flipkart(soup):
    body = soup.get_text(" ", strip=True).lower()
    for w in ["sold out", "currently unavailable", "coming soon", "this item is currently out of stock"]:
        if w in body:
            return OUT, w
    for b in soup.find_all(["button", "li", "div"]):
        t = text_of(b).lower()
        if t in ("add to cart", "buy now", "buy now at", "go to cart") or t.startswith("buy now"):
            return IN, f"'{text_of(b)}' button"
    return None, ""


def check_generic(soup):
    buttons = " | ".join(text_of(b).lower() for b in soup.find_all(["button", "a", "input"])[:600]
                         if len(text_of(b)) < 40) + " " + " ".join(
        (i.get("value") or "").lower() for i in soup.find_all("input"))
    page = soup.get_text(" ", strip=True).lower()[:60000]
    for w in OUT_WORDS:
        if w in page and not any(iw in buttons for iw in IN_WORDS[:4]):
            return OUT, w
    for w in IN_WORDS:
        if w in buttons:
            return IN, f"'{w}' button"
    return None, ""


def check_product(product, settings):
    url = product["url"]
    site = site_name(url)
    html, err = fetch_html(url, settings)
    if err or not html:
        return {"status": UNKNOWN, "price": "", "detail": err or "no page", "site": site}
    soup = BeautifulSoup(html, "html.parser")
    title = text_of(soup.select_one("#productTitle")) or text_of(soup.select_one("h1")) or \
        (soup.title.string.strip() if soup.title and soup.title.string else "")
    status, detail = None, ""
    if site == "Amazon":
        status, detail = check_amazon(soup)
    elif site == "Flipkart":
        status, detail = check_flipkart(soup)
    if not status:
        status, detail = schema_availability(soup)
    if not status:
        status, detail = check_generic(soup)
    # extra rule: "must contain" / "must not contain" text the user can set per product
    page = soup.get_text(" ", strip=True).lower()
    if product.get("in_stock_text") and product["in_stock_text"].lower() in page:
        status, detail = IN, f"found '{product['in_stock_text']}'"
    if product.get("out_of_stock_text") and product["out_of_stock_text"].lower() in page:
        status, detail = OUT, f"found '{product['out_of_stock_text']}'"
    return {"status": status or UNKNOWN, "price": find_price(soup, html), "detail": detail or err or "could not decide",
            "site": site, "title": title[:150], "image": find_image(soup)}


def find_image(soup):
    el = soup.select_one("#landingImage, #imgBlkFront")
    if el:
        for k in ("data-old-hires", "src"):
            v = el.get(k) or ""
            if v.startswith("http"):
                return v
    for sel in ["meta[property='og:image']", "meta[name='og_image']", "meta[name='twitter:image']", "[itemprop=image]"]:
        el = soup.select_one(sel)
        v = el and (el.get("content") or el.get("src") or el.get("href"))
        if v and v.startswith("http"):
            return v
    m = re.search(r'"image"\s*:\s*\[?\s*"(https?://[^"]+)"', " ".join(
        x.string or "" for x in soup.find_all("script", type="application/ld+json")))
    return m.group(1) if m else ""


# --------------------------------------------------------------------------- notifications
REQUIRED = {"telegram": ["bot_token", "chat_id"], "whatsapp": ["phone", "apikey"], "google_chat": ["webhook_url"],
            "email": ["username", "password"], "ntfy": ["topic"], "discord": ["webhook_url"], "slack": ["webhook_url"]}
CHANNELS = ["desktop", "telegram", "whatsapp", "google_chat", "email", "ntfy", "discord", "slack"]


def notify_all(cfg, subject, message, attach=None, alert=None, only=None):
    """Send to every enabled channel (or just `only`). `alert` carries product details for rich messages."""
    n = cfg["notifications"]
    sent, errors = [], {}
    fns = {"telegram": send_telegram, "whatsapp": send_whatsapp, "google_chat": send_google_chat,
           "email": send_email, "discord": send_discord, "slack": send_slack, "ntfy": send_ntfy,
           "desktop": send_desktop}
    for name in ([only] if only else CHANNELS):
        c = n.get(name) or ({"enabled": True} if name == "desktop" else {})
        if not only and not c.get("enabled"):
            continue
        missing = [f for f in REQUIRED.get(name, []) if not str(c.get(f) or "").strip()]
        if name == "whatsapp" and c.get("provider") == "twilio":
            missing = [f for f in ("account_sid", "auth_token", "from_number", "to_number") if not c.get(f)]
        if missing:
            labels = {"username": "your Gmail address", "password": "app password", "apikey": "API key",
                      "phone": "your WhatsApp number", "bot_token": "bot token", "chat_id": "chat ID",
                      "webhook_url": "webhook URL", "topic": "topic name"}
            errors[name] = "Please fill in " + " and ".join(labels.get(m, m.replace("_", " ")) for m in missing)
            continue
        try:
            fns[name](c, subject, message, alert or {}, attach)
            sent.append(name)
        except Exception as e:
            errors[name] = friendly_error(name, e)
            print(f"   ! {name} failed: {errors[name]}")
    return sent, errors


def friendly_error(name, e):
    msg = str(e)
    if isinstance(e, KeyError):
        return f"Missing setting: {e.args[0]}"
    if "401" in msg or "Unauthorized" in msg:
        return "Wrong token / key - copy it again"
    if "400" in msg and name == "telegram":
        return "Chat ID is wrong, or you haven't pressed Start in your bot yet"
    if "403" in msg and name == "telegram":
        return "Open your bot in Telegram and press Start first"
    if "Username and Password not accepted" in msg or "535" in msg:
        return "Gmail rejected the login - use a 16-letter App Password, not your normal password"
    if "404" in msg:
        return "Link / webhook URL not found - copy it again"
    return msg[:200]


def send_desktop(c, subject, msg, alert, attach=None):
    """Pop-up notification on this computer (Windows/Mac/Linux). Needs: pip install plyer"""
    try:
        from plyer import notification
    except ImportError:
        return
    body = f"{alert['product']}\n{alert.get('price', '')}  {alert.get('site', '')}" if alert.get("product") else msg
    notification.notify(title=subject[:63], message=body[:250], app_name="Restock Watch", timeout=20)


def _html(s):
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def send_telegram(c, subject, msg, alert, attach=None):
    api = f"https://api.telegram.org/bot{str(c['bot_token']).strip()}"
    chat = str(c["chat_id"]).strip()
    if alert.get("product"):
        lines = [f"<b>{_html(subject)}</b>", "", f"<b>{_html(alert['product'])}</b>"]
        if alert.get("price"):
            lines.append(f"💰 {_html(alert['price'])}")
        lines.append(f"🛒 {_html(alert.get('site', ''))}")
        lines.append(f"🕒 {_html(alert.get('when', ''))}")
        if alert.get("previous"):
            lines.append(f"<i>{_html(alert['previous'])}</i>")
        text = "\n".join(lines)
        markup = {"inline_keyboard": [[{"text": f"Open on {alert.get('site', 'site')}", "url": alert["url"]}]]} \
            if alert.get("url") else None
        if alert.get("image"):
            r = requests.post(f"{api}/sendPhoto", json={"chat_id": chat, "photo": alert["image"], "caption": text,
                                                         "parse_mode": "HTML", "reply_markup": markup}, timeout=25)
            if r.ok:
                return
        r = requests.post(f"{api}/sendMessage", json={"chat_id": chat, "text": text, "parse_mode": "HTML",
                                                       "reply_markup": markup, "disable_web_page_preview": True},
                          timeout=20)
    else:
        r = requests.post(f"{api}/sendMessage", json={"chat_id": chat, "text": f"{subject}\n\n{msg}"}, timeout=20)
    if not r.ok:
        raise RuntimeError(f"{r.status_code} {r.text[:150]}")


def telegram_find_chat(token):
    """After the user sends any message to their bot, this finds their chat ID automatically."""
    r = requests.get(f"https://api.telegram.org/bot{token.strip()}/getUpdates", timeout=20)
    if r.status_code == 401 or r.status_code == 404:
        raise RuntimeError("That bot token is wrong - copy it again from @BotFather")
    r.raise_for_status()
    for upd in reversed(r.json().get("result", [])):
        m = upd.get("message") or upd.get("channel_post") or upd.get("my_chat_member") or {}
        chat = m.get("chat")
        if chat:
            return {"chat_id": str(chat["id"]),
                    "name": chat.get("first_name") or chat.get("title") or chat.get("username") or ""}
    raise RuntimeError("No message found - open your bot in Telegram, press Start, send 'hi', then try again")


def plain_text(subject, msg, alert, bold="*"):
    if not alert.get("product"):
        return f"{bold}{subject}{bold}\n\n{msg}"
    lines = [f"{bold}{subject}{bold}", "", f"{bold}{alert['product']}{bold}"]
    if alert.get("price"):
        lines.append(f"💰 {alert['price']}")
    lines += [f"🛒 {alert.get('site', '')}", f"🕒 {alert.get('when', '')}"]
    if alert.get("previous"):
        lines.append(f"_{alert['previous']}_" if bold == "*" else alert["previous"])
    if alert.get("url"):
        lines += ["", alert["url"]]
    return "\n".join(lines)


def send_whatsapp(c, subject, msg, alert, attach=None):
    text = plain_text(subject, msg, alert)
    provider = c.get("provider", "callmebot")
    if provider == "callmebot":
        phone = re.sub(r"[^\d+]", "", str(c["phone"]))
        r = requests.get("https://api.callmebot.com/whatsapp.php",
                         params={"phone": phone, "text": text, "apikey": str(c["apikey"]).strip()}, timeout=30)
        if not r.ok or "APIKey is invalid" in r.text or "ERROR" in r.text.upper()[:300]:
            raise RuntimeError(f"CallMeBot said: {re.sub('<[^>]+>', ' ', r.text)[:150].strip()}")
    elif provider == "twilio":
        r = requests.post(f"https://api.twilio.com/2010-04-01/Accounts/{c['account_sid']}/Messages.json",
                          data={"From": f"whatsapp:{c['from_number']}", "To": f"whatsapp:{c['to_number']}",
                                "Body": text}, auth=(c["account_sid"], c["auth_token"]), timeout=30)
        r.raise_for_status()
    else:
        raise ValueError(f"unknown whatsapp provider {provider}")


def send_google_chat(c, subject, msg, alert, attach=None):
    url = str(c["webhook_url"]).strip()
    if alert.get("product"):
        widgets = [{"decoratedText": {"topLabel": "Product", "text": f"<b>{_html(alert['product'])}</b>", "wrapText": True}}]
        if alert.get("price"):
            widgets.append({"decoratedText": {"topLabel": "Price", "text": _html(alert["price"])}})
        widgets.append({"decoratedText": {"topLabel": "When", "text": _html(alert.get("when", ""))}})
        if alert.get("previous"):
            widgets.append({"textParagraph": {"text": f"<i>{_html(alert['previous'])}</i>"}})
        if alert.get("url"):
            widgets.append({"buttonList": {"buttons": [{"text": f"Open on {alert.get('site', 'site')}",
                                                        "onClick": {"openLink": {"url": alert["url"]}}}]}})
        header = {"title": subject, "subtitle": alert.get("site", "")}
        if alert.get("image"):
            header.update({"imageUrl": alert["image"], "imageType": "SQUARE"})
        body = {"text": f"{subject}: {alert['product']}",
                "cardsV2": [{"cardId": "stock", "card": {"header": header, "sections": [{"widgets": widgets}]}}]}
        r = requests.post(url, json=body, timeout=20)
        if r.ok:
            return
    r = requests.post(url, json={"text": plain_text(subject, msg, alert)}, timeout=20)
    r.raise_for_status()


def send_discord(c, subject, msg, alert, attach=None):
    requests.post(c["webhook_url"], json={"content": plain_text(subject, msg, alert, "**")}, timeout=20).raise_for_status()


def send_slack(c, subject, msg, alert, attach=None):
    requests.post(c["webhook_url"], json={"text": plain_text(subject, msg, alert)}, timeout=20).raise_for_status()


def send_ntfy(c, subject, msg, alert, attach=None):
    server = c.get("server", "https://ntfy.sh").rstrip("/")
    headers = {"Title": subject.encode("ascii", "ignore").decode().strip() or "Stock alert", "Priority": "high"}
    if alert.get("url"):
        headers["Click"] = alert["url"]
    requests.post(f"{server}/{c['topic']}", data=plain_text(subject, msg, alert, "").encode("utf-8"),
                  headers=headers, timeout=20).raise_for_status()


def send_email(c, subject, msg, alert, attach=None):
    m = MIMEMultipart()
    to = c.get("to") or c["username"]
    m["From"], m["To"], m["Subject"] = c["username"], ", ".join(to) if isinstance(to, list) else to, subject
    m.attach(MIMEText(plain_text(subject, msg, alert, ""), "plain", "utf-8"))
    for path in (attach if c.get("attach_reports", True) else None) or []:
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(path)}"'
            m.attach(part)
    with smtplib.SMTP(c.get("smtp_server", "smtp.gmail.com"), int(c.get("smtp_port", 587)), timeout=30) as s:
        s.starttls()
        s.login(c["username"], str(c["password"]).replace(" ", ""))
        s.send_message(m)


# --------------------------------------------------------------------------- reports
TS_FMT = "%Y-%m-%d %H:%M:%S"


def fmt_dur(sec):
    sec = int(max(sec, 0))
    d, r = divmod(sec, 86400)
    h, r = divmod(r, 3600)
    m = r // 60
    if d:
        return f"{d}d {h}h"
    if h:
        return f"{h}h {m}m"
    return f"{m} min" if m else "less than a minute"


def fmt_when(ts):
    try:
        return datetime.strptime(ts, TS_FMT).strftime("%d %b %Y, %I:%M %p")
    except (TypeError, ValueError):
        return ts or ""


def short_when(ts):
    try:
        return datetime.strptime(ts, TS_FMT).strftime("%d %b, %I:%M %p")
    except (TypeError, ValueError):
        return ts or ""


def product_timelines(cfg, events=None, state=None, include_removed=False, only=None):
    """One entry per product, with its own in-stock / out-of-stock periods (newest first) and a summary.
    Products you removed are hidden unless include_removed=True (their history is kept on disk).
    `only` = list of product links to include (works for removed products too)."""
    events = read_csv(EVENTS_CSV) if events is None else events
    state = load_state() if state is None else state
    current = now(cfg).replace(tzinfo=None)
    by_url = {}
    for e in events:
        by_url.setdefault(e.get("url") or e["product"], []).append(e)
    order = [p["url"] for p in cfg["products"]]
    removed_urls = [u for u in by_url if u not in order and "B0CHX1W1XY" not in u]
    if only:
        wanted = set(only)
        order = [u for u in order + removed_urls if u in wanted]
    elif include_removed:
        order += removed_urls
    conf = {p["url"]: p for p in cfg["products"]}
    out = []
    for url in order:
        evs = sorted(by_url.get(url, []), key=lambda e: f"{e['date']} {e['time']}")
        p, s = conf.get(url, {}), state.get(url, {})
        name = p.get("name") or s.get("name") or (evs[-1]["product"] if evs else "") or s.get("title") or url
        removed = url not in conf
        periods = []
        for e in evs:
            ts = f"{e['date']} {e['time']}"
            if e["new_status"] not in (IN, OUT):
                continue
            if periods and periods[-1]["status"] == e["new_status"]:
                continue
            if periods:
                periods[-1]["end"] = ts
            periods.append({"status": e["new_status"], "start": ts, "end": None, "price": e.get("price", "")})
        end_now = (s.get("last_checked") if removed else None) or current.strftime(TS_FMT)
        secs = {IN: 0, OUT: 0}
        for per in periods:
            a = datetime.strptime(per["start"], TS_FMT)
            b = datetime.strptime(per["end"] or end_now, TS_FMT)
            per["secs"] = max(0, (b - a).total_seconds())
            per["lasted"] = fmt_dur(per["secs"])
            per["ongoing"] = per["end"] is None and not removed
            secs[per["status"]] += per["secs"]
        tracked = secs[IN] + secs[OUT]
        out.append({
            "name": name, "url": url, "site": site_name(url) if url.startswith("http") else "",
            "image": s.get("image", ""), "removed": removed, "price": s.get("price", ""),
            "status": s.get("status") or (periods[-1]["status"] if periods else "NOT CHECKED"),
            "since": s.get("since") or (periods[-1]["start"] if periods else ""),
            "last_checked": s.get("last_checked", ""),
            "came_back": sum(1 for e in evs if e["event"] == "BACK IN STOCK"),
            "went_out": sum(1 for e in evs if e["event"] == "WENT OUT OF STOCK"),
            "time_in": fmt_dur(secs[IN]), "time_out": fmt_dur(secs[OUT]),
            "pct_in": round(100 * secs[IN] / tracked) if tracked else None,
            "tracking_since": periods[0]["start"] if periods else "",
            "periods": list(reversed(periods)),             # newest first
        })
    return out


def build_reports(cfg):
    events = read_csv(EVENTS_CSV)
    checks = read_csv(CHECKS_CSV)
    items = product_timelines(cfg, events, load_state())
    build_excel(cfg, items, checks)
    build_pdf(cfg, items)
    return XLSX_FILE, PDF_FILE


def _safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n]+', " ", str(name)).strip()
    return (name[:60].strip() or "product")


def export_products(cfg, urls, fmt):
    """Report for just the chosen products. Returns (bytes, file name, content type)."""
    items = product_timelines(cfg, only=urls)
    if not items:
        raise ValueError("Pick at least one product")
    base = _safe_filename(items[0]["name"]) if len(items) == 1 else f"{len(items)} products"
    stamp = now(cfg).strftime("%Y-%m-%d")
    out_dir = os.path.join(DATA, "exports")
    os.makedirs(out_dir, exist_ok=True)
    if fmt == "csv":
        import io
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["Product", "Site", "Status", "From", "To", "Lasted", "Price", "Link"])
        for it in items:
            for per in it["periods"]:
                w.writerow([it["name"], it["site"], _status_word(per["status"], per["ongoing"]), fmt_when(per["start"]),
                            "still now" if per["ongoing"] else fmt_when(per["end"] or it["last_checked"]),
                            per["lasted"], per["price"], it["url"]])
        return ("\ufeff" + buf.getvalue()).encode("utf-8"), f"{base} - stock history {stamp}.csv", "text/csv; charset=utf-8"
    path = os.path.join(out_dir, f"export.{fmt}")
    if fmt == "xlsx":
        wanted = {it["url"] for it in items}
        build_excel(cfg, items, [c for c in read_csv(CHECKS_CSV) if c.get("url") in wanted], path)
        ctype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif fmt == "pdf":
        build_pdf(cfg, items, path)
        ctype = "application/pdf"
    else:
        raise ValueError("Unknown format")
    with open(path, "rb") as f:
        return f.read(), f"{base} - stock history {stamp}.{fmt}", ctype


def _status_word(st, ongoing=False):
    if st == IN:
        return "Available" + (" (now)" if ongoing else "")
    if st == OUT:
        return "Out of stock" + (" (now)" if ongoing else "")
    return "Not checked yet"


def build_excel(cfg, items, checks, path=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    navy, white = "1F3A5F", "FFFFFF"
    fill = lambda c: PatternFill("solid", fgColor=c)
    GREEN, RED, GREY, BAND = fill("D4EDDA"), fill("F8D7DA"), fill("EEEEEE"), fill("E8EEF7")
    thin = Side(style="thin", color="D0D5DD")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)
    status_fill = lambda st: GREEN if st == IN else RED if st == OUT else GREY

    def header(ws, row, titles):
        for i, t in enumerate(titles, 1):
            c = ws.cell(row, i, t)
            c.font, c.fill, c.border = Font(bold=True, color=white), fill(navy), box
            c.alignment = Alignment(vertical="center")

    # ---- Sheet 1: one line per product
    ws = wb.active
    ws.title = "Summary"
    cols = ["Product", "Site", "Status now", "Since", "Price", "Came back in stock", "Went out of stock",
            "Total time available", "Total time out of stock", "Available % of time", "Tracking since", "Link"]
    header(ws, 1, cols)
    for r, it in enumerate(items, 2):
        vals = [it["name"] + ("  (removed)" if it["removed"] else ""), it["site"], _status_word(it["status"]),
                fmt_when(it["since"]), it["price"], f"{it['came_back']} times", f"{it['went_out']} times",
                it["time_in"] if it["periods"] else "", it["time_out"] if it["periods"] else "",
                f"{it['pct_in']}%" if it["pct_in"] is not None else "", fmt_when(it["tracking_since"]), it["url"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = box
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 1))
        ws.cell(r, 3).fill = status_fill(it["status"])
        ws.cell(r, 3).font = Font(bold=True)
        if it["url"].startswith("http"):
            ws.cell(r, 12).hyperlink = it["url"]
            ws.cell(r, 12).value = "Open"
            ws.cell(r, 12).font = Font(color="1D6FE0", underline="single")
    for i, w in enumerate([46, 14, 16, 22, 12, 12, 12, 14, 14, 12, 22, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"

    # ---- Sheet 2: item-wise history, one block per product
    ws = wb.create_sheet("Item-wise history")
    widths = [18, 24, 24, 16, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r = 1
    ws.cell(r, 1, "Each product's stock history - latest first").font = Font(bold=True, size=14)
    r += 2
    for it in items:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, it["name"] + ("  (removed)" if it["removed"] else ""))
        c.font, c.fill = Font(bold=True, size=12, color=navy), BAND
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 32
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        summary = f"{it['site']}  |  Now: {_status_word(it['status'])}"
        if it["since"]:
            summary += f" since {fmt_when(it['since'])}"
        if it["periods"]:
            summary += (f"  |  Came back {it['came_back']}x, went out {it['went_out']}x"
                        f"  |  Available {it['pct_in']}% of the time")
        ws.cell(r, 1, summary).font = Font(italic=True, color="5B6475")
        r += 1
        header(ws, r, ["Status", "From", "To", "Lasted", "Price"])
        r += 1
        if not it["periods"]:
            ws.cell(r, 1, "No checks recorded yet").font = Font(italic=True, color="8A93A6")
            r += 1
        for per in it["periods"]:
            vals = [_status_word(per["status"], per["ongoing"]), fmt_when(per["start"]),
                    "still now" if per["ongoing"] else fmt_when(per["end"] or it["last_checked"]),
                    per["lasted"], per["price"]]
            for cidx, v in enumerate(vals, 1):
                cell = ws.cell(r, cidx, v)
                cell.border = box
            ws.cell(r, 1).fill = status_fill(per["status"])
            ws.cell(r, 1).font = Font(bold=True)
            r += 1
        r += 2

    # ---- Sheet 3: every check (raw), for anyone who wants detail
    ws = wb.create_sheet("All checks (raw)")
    header(ws, 1, ["Date", "Time", "Product", "Site", "Status", "Price"])
    for i, ch in enumerate(reversed(checks[-5000:]), 2):
        for cidx, k in enumerate(["date", "time", "product", "site", "status", "price"], 1):
            ws.cell(i, cidx, ch.get(k, ""))
        ws.cell(i, 5).fill = status_fill(ch.get("status"))
    for i, w in enumerate([12, 10, 46, 14, 16, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path or XLSX_FILE)


def build_pdf(cfg, items, path=None):
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    st = getSampleStyleSheet()
    small = st["BodyText"].clone("small", fontSize=8.5, leading=11)
    meta = st["BodyText"].clone("meta", fontSize=9, leading=12, textColor=colors.HexColor("#5B6475"))
    h3 = st["Heading3"].clone("h3", fontSize=12.5, leading=15, spaceAfter=2, textColor=colors.HexColor("#1F3A5F"))
    GREEN, RED, GREY = colors.HexColor("#D4EDDA"), colors.HexColor("#F8D7DA"), colors.HexColor("#EEEEEE")
    G2, R2 = colors.HexColor("#0F9D58"), colors.HexColor("#D93025")
    W = 186 * mm
    doc = SimpleDocTemplate(path or PDF_FILE, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm, title="Stock report - item wise")
    esc = lambda s: str(s or "").replace("₹", "Rs ").replace("&", "&amp;").replace("<", "&lt;")
    P = lambda s, style=small: Paragraph(esc(s), style)
    fillfor = lambda stt: GREEN if stt == IN else RED if stt == OUT else GREY

    def table(head, rows, widths, status_col=None, statuses=None):
        data = [[Paragraph(f'<font color="white"><b>{h}</b></font>', small) for h in head]] + \
               [[P(c) for c in r] for r in rows]
        t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
        style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                 ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C9CED8")),
                 ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
        for i, stt in enumerate(statuses or [], 1):
            style.append(("BACKGROUND", (status_col, i), (status_col, i), fillfor(stt)))
        t.setStyle(TableStyle(style))
        return t

    def timeline_bar(it):
        """Green/red strip: the product's whole tracked time, oldest on the left, now on the right."""
        pers = list(reversed(it["periods"]))
        total = sum(p["secs"] for p in pers) or 1
        d = Drawing(W, 16)
        x = 0
        for p in pers:
            w = max(1.2, W * p["secs"] / total)
            d.add(Rect(x, 4, w, 9, fillColor=G2 if p["status"] == IN else R2, strokeColor=None))
            x += w
        d.add(String(0, -5, (fmt_when(pers[0]["start"]) if pers else "") + "   (green = available, red = out of stock)",
                     fontName="Helvetica", fontSize=7, fillColor=colors.grey))
        d.add(String(W, -5, "now", fontName="Helvetica", fontSize=7, fillColor=colors.grey, textAnchor="end"))
        return d

    t = now(cfg)
    parts = [Paragraph("Stock report - item wise", st["Title"]),
             P(f"Generated {t.strftime('%d %b %Y, %I:%M %p')}.  Each product has its own section below: "
               f"when it was available, when it was out of stock, and how long each lasted (latest first).", meta),
             Spacer(1, 5 * mm), Paragraph("All products at a glance", st["Heading2"])]
    rows, sts = [], []
    for it in items:
        rows.append([it["name"] + (" (removed)" if it["removed"] else ""), it["site"], _status_word(it["status"]),
                     short_when(it["since"]), f"{it['came_back']}x", f"{it['went_out']}x",
                     f"{it['pct_in']}%" if it["pct_in"] is not None else "-"])
        sts.append(it["status"])
    parts.append(table(["Product", "Site", "Now", "Since", "Came back", "Went out", "Available"],
                       rows or [["No products yet", "", "", "", "", "", ""]], [60, 19, 24, 31, 17, 17, 18],
                       status_col=2, statuses=sts))
    parts.append(Spacer(1, 8 * mm))

    for it in items:
        block = [Paragraph(esc(it["name"]) + (" <font color='#8A93A6'>(removed)</font>" if it["removed"] else ""), h3)]
        line = f"{it['site']}  |  Now: <b>{_status_word(it['status'])}</b>"
        if it["since"]:
            line += f" since {fmt_when(it['since'])}"
        if it["price"]:
            line += f"  |  {it['price']}"
        block.append(Paragraph(esc(line).replace("&lt;b>", "<b>").replace("&lt;/b>", "</b>"), meta))
        if it["periods"]:
            block.append(P(f"Came back in stock {it['came_back']} times, went out of stock {it['went_out']} times.  "
                           f"Available {it['time_in']} in total, out of stock {it['time_out']} "
                           f"({it['pct_in']}% available).", meta))
            block += [Spacer(1, 2 * mm), timeline_bar(it), Spacer(1, 4 * mm)]
            prow = [[_status_word(p["status"], p["ongoing"]), fmt_when(p["start"]),
                     "still now" if p["ongoing"] else fmt_when(p["end"] or it["last_checked"]), p["lasted"], p["price"]]
                    for p in it["periods"][:60]]
            block.append(table(["Status", "From", "To", "Lasted", "Price"], prow, [36, 42, 42, 36, 30],
                               status_col=0, statuses=[p["status"] for p in it["periods"][:60]]))
        else:
            block.append(P("No checks recorded yet.", meta))
        block.append(Spacer(1, 9 * mm))
        parts.append(KeepTogether(block) if len(it["periods"]) <= 18 else block[0])
        if len(it["periods"]) > 18:
            parts += block[1:]
    doc.build(parts)


# --------------------------------------------------------------------------- main logic
import threading
from concurrent.futures import ThreadPoolExecutor

CHECK_LOCK = threading.Lock()
WAKE = threading.Event()
FORCE = threading.Event()
PROGRESS = {"running": False, "done": 0, "total": 0, "current": "", "next_check": None, "last_run": None}

EVENT_TITLES = {
    "FIRST CHECK - IN STOCK": "✅ AVAILABLE NOW",
    "FIRST CHECK - OUT OF STOCK": "⛔ NOT AVAILABLE RIGHT NOW",
    "WENT OUT OF STOCK": "🔴 WENT OUT OF STOCK",
    "BACK IN STOCK": "🟢 BACK IN STOCK - available again!",
}


def interval_for(cfg, status):
    """Out-of-stock products are watched more closely so you hear the moment they come back."""
    s = cfg["settings"]
    if status == IN:
        return float(s.get("check_every_minutes", 5)) * 60
    return float(s.get("fast_check_minutes", 2)) * 60


def _check_site_group(group, s):
    """Products of the SAME site run one after another (polite); different sites run in parallel."""
    out = []
    for i, p in enumerate(group):
        PROGRESS["current"] = p.get("name") or p["url"]
        try:
            res = check_product(p, s)
        except Exception as e:
            res = {"status": UNKNOWN, "price": "", "detail": f"error: {e}", "site": site_name(p["url"])}
        out.append((p, res))
        PROGRESS["done"] += 1
        if i < len(group) - 1:
            time.sleep(random.uniform(*s.get("delay_between_products_seconds", [2, 4])))
    return out


def check_all(cfg, send=True, products=None):
    if not CHECK_LOCK.acquire(blocking=False):
        print("A check is already running - skipped.")
        return False
    try:
        return _check_all(cfg, send, products)
    finally:
        PROGRESS.update(running=False, current="", last_run=now(cfg).strftime("%Y-%m-%d %H:%M:%S"))
        CHECK_LOCK.release()


def _check_all(cfg, send, products):
    s = cfg["settings"]
    if products is None:
        products = [p for p in cfg["products"] if p.get("active", True)]
    if not products:
        return True
    PROGRESS.update(running=True, done=0, total=len(products), current="")
    groups = {}
    for p in products:
        groups.setdefault(site_name(p["url"]), []).append(p)
    with ThreadPoolExecutor(max_workers=min(6, len(groups))) as ex:
        results = [r for chunk in ex.map(lambda g: _check_site_group(g, s), groups.values()) for r in chunk]

    state = load_state()
    alerts = []
    for p, res in results:
        t = now(cfg)
        date, clock = t.strftime("%Y-%m-%d"), t.strftime("%H:%M:%S")
        name = p.get("name") or res.get("title") or p["url"]
        prev = state.get(p["url"], {})
        old = prev.get("status")
        print(f"[{date} {clock}] {res['status']:<13} {res['price']:<10} {name[:50]}  ({res['detail'][:60]})")
        append_csv(CHECKS_CSV, CHECK_FIELDS, {"date": date, "time": clock, "product": name, "site": res["site"],
                                              "status": res["status"], "price": res["price"], "detail": res["detail"],
                                              "url": p["url"]})
        entry = dict(prev)
        entry.update({"last_checked": f"{date} {clock}", "last_ts": time.time(), "detail": res["detail"],
                      "title": res.get("title") or prev.get("title", "")})
        if res.get("image"):
            entry["image"] = res["image"]
        if res["price"]:
            entry["price"] = res["price"]
        when = t.strftime("%d %b %Y, %I:%M:%S %p")
        base = {"product": name, "site": res["site"], "price": res["price"] or prev.get("price", ""),
                "url": p["url"], "image": entry.get("image", ""), "when": when}
        if res["status"] == UNKNOWN:
            entry["fails"] = prev.get("fails", 0) + 1
            if entry["fails"] == int(s.get("alert_after_failures", 6)) and send:
                alerts.append(("⚠️ Can't read this product page",
                               dict(base, previous=f"The site blocked {entry['fails']} checks in a row. It keeps trying.")))
        else:
            entry["fails"] = 0
            if res["status"] != old:
                event = (f"FIRST CHECK - {res['status']}" if old is None else
                         "BACK IN STOCK" if res["status"] == IN else "WENT OUT OF STOCK")
                entry.update({"status": res["status"], "since": f"{date} {clock}"})
                append_csv(EVENTS_CSV, EVENT_FIELDS, {"date": date, "time": clock, "product": name, "site": res["site"],
                                                      "event": event, "old_status": old or "-", "new_status": res["status"],
                                                      "price": res["price"], "url": p["url"], "note": res["detail"]})
                if not (old is None and not s.get("notify_on_first_check", True)):
                    previous = ""
                    if old and prev.get("since"):
                        dur = fmt_dur(time.time() - datetime.strptime(prev["since"], "%Y-%m-%d %H:%M:%S")
                                      .replace(tzinfo=tz(cfg)).timestamp())
                        previous = f"Was {'out of stock' if old == OUT else 'available'} for {dur}"
                    alerts.append((EVENT_TITLES[event], dict(base, previous=previous)))
            target = p.get("target_price")
            pnum = to_num(res["price"])
            if target and pnum and pnum <= float(target) and res["status"] == IN and not prev.get("target_hit"):
                entry["target_hit"] = True
                alerts.append((f"💰 PRICE DROP - now {res['price']}", dict(base, previous=f"Your alert price: ₹{target:,.0f}")))
            elif target and pnum and pnum > float(target):
                entry["target_hit"] = False
        state[p["url"]] = entry
    save_state(state)

    # alerts go out FIRST (so you hear immediately), reports are rebuilt after
    if send:
        for title, alert in alerts:
            sent, errs = notify_all(cfg, title, "", alert=alert)
            print(f"   -> {title}: sent to {', '.join(sent) or 'no channel'}")
    try:
        build_reports(cfg)
    except Exception as e:
        print("Report error:", e)
    return True


def to_num(price):
    m = re.search(r"[\d,]+(?:\.\d+)?", price or "")
    try:
        return float(m.group(0).replace(",", "")) if m else None
    except ValueError:
        return None


def daily_summary(cfg):
    state = load_state()
    lines = []
    for p in cfg["products"]:
        s = state.get(p["url"], {})
        lines.append(f"{'🟢' if s.get('status') == IN else '🔴' if s.get('status') == OUT else '⚪'} "
                     f"{(p.get('name') or s.get('title') or p['url'])[:50]} - {s.get('status', 'not checked')} "
                     f"{s.get('price', '')}")
    xlsx, pdf = build_reports(cfg)
    return notify_all(cfg, "📊 Daily stock summary", "\n".join(lines) or "No products", attach=[xlsx, pdf])


def due_products(cfg, force=False):
    state, t = load_state(), time.time()
    due, next_at = [], None
    for p in cfg["products"]:
        if not p.get("active", True):
            continue
        st_ = state.get(p["url"], {})
        at = st_.get("last_ts", 0) + interval_for(cfg, st_.get("status"))
        if force or at <= t:
            due.append(p)
        else:
            next_at = at if next_at is None else min(next_at, at)
    return due, next_at


def run_forever(cfg_initial=None):
    last_summary_day = None
    while True:
        cfg = load_config()  # re-read every round so changes from the dashboard apply at once
        try:
            due, _ = due_products(cfg, force=FORCE.is_set())
            FORCE.clear()
            if due:
                check_all(cfg, products=due)
        except Exception as e:
            print("Error during check:", e)
        s = cfg["settings"]
        hhmm = s.get("daily_summary_time")
        t = now(cfg)
        if hhmm and t.strftime("%H:%M") >= str(hhmm) and last_summary_day != t.date():
            last_summary_day = t.date()
            try:
                daily_summary(cfg)
            except Exception as e:
                print("Summary error:", e)
        _, next_at = due_products(cfg)
        PROGRESS["next_check"] = next_at
        wait = max(1, min(15, (next_at - time.time()) if next_at else 15))
        WAKE.wait(wait)   # "Check now" / adding a product wakes this up early
        WAKE.clear()


# --------------------------------------------------------------------------- find / restore old data
SKIP_DIRS = {"appdata", "node_modules", ".git", "windows", "program files", "program files (x86)", "$recycle.bin",
             "site-packages", "__pycache__", ".cache", "library", "venv", ".venv", "programdata"}


def _data_info(folder):
    cfg_path, data = os.path.join(folder, "config.yaml"), os.path.join(folder, "data")
    products = 0
    try:
        with open(cfg_path, encoding="utf-8") as f:
            products = len((yaml.safe_load(f) or {}).get("products") or [])
    except Exception:
        pass
    events = max(0, len(read_csv(os.path.join(data, "events.csv"))))
    checks = 0
    try:
        with open(os.path.join(data, "checks.csv"), encoding="utf-8-sig") as f:
            checks = max(0, sum(1 for _ in f) - 1)
    except OSError:
        pass
    stamps = [os.path.getmtime(p) for p in (cfg_path, os.path.join(data, "state.json"), os.path.join(data, "events.csv"))
              if os.path.exists(p)]
    if not products and not events and not checks:
        return None
    return {"path": folder, "products": products, "events": events, "checks": checks,
            "last_used": datetime.fromtimestamp(max(stamps)).strftime("%d %b %Y, %I:%M %p") if stamps else "",
            "_ts": max(stamps) if stamps else 0}


def find_old_data(max_dirs=60000):
    """Look through the usual places (Downloads, Desktop, Documents, OneDrive, next to this folder, backups)
    for other Restock Watch / Stock Tracker folders that still hold products or history."""
    home = os.path.expanduser("~")
    roots = [os.path.dirname(BASE), os.path.join(BASE, "backup")]
    for name in ("Downloads", "Desktop", "Documents"):
        roots.append(os.path.join(home, name))
    try:
        roots += [os.path.join(home, d) for d in os.listdir(home) if d.lower().startswith("onedrive")]
    except OSError:
        pass
    roots.append(home)
    me = os.path.normcase(os.path.abspath(BASE))
    seen, found, visited = set(), [], 0
    for root in roots:
        if not os.path.isdir(root):
            continue
        base_depth = root.rstrip("\\/").count(os.sep)
        for dirpath, dirnames, filenames in os.walk(root):
            visited += 1
            if visited > max_dirs:
                break
            depth = dirpath.count(os.sep) - base_depth
            dirnames[:] = [d for d in dirnames if d.lower() not in SKIP_DIRS and not d.startswith(".")
                           and depth < (2 if root == home else 6)]
            key = os.path.normcase(os.path.abspath(dirpath))
            if key == me or key in seen:
                continue
            looks_like_app = "stock_tracker.py" in filenames or (
                "config.yaml" in filenames and os.path.isdir(os.path.join(dirpath, "data"))) or (
                os.path.basename(os.path.dirname(dirpath)).lower() == "backup" and "config.yaml" in filenames)
            if looks_like_app:
                seen.add(key)
                info = _data_info(dirpath)
                if info:
                    found.append(info)
    found.sort(key=lambda x: (x["events"] + x["checks"] > 0, x["_ts"]), reverse=True)
    return found


def _is_example(p):
    return str(p.get("name", "")).startswith("Example -") or "B0CHX1W1XY" in str(p.get("url", ""))


def _merge_csv(dst, src, fields, key):
    rows = {}
    for path in (dst, src):
        for r in read_csv(path):
            rows.setdefault(key(r), r)
    out = sorted(rows.values(), key=lambda r: f"{r.get('date', '')} {r.get('time', '')}")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in out:
            w.writerow({k: r.get(k, "") for k in fields})
    return len(out)


def restore_old_data(folder):
    """MERGE products, alert settings and full history from `folder` into this app.
    Nothing is deleted: what is here stays, the other folder's data is added, and a backup is made first."""
    import shutil
    folder = os.path.abspath(folder)
    if os.path.normcase(folder) == os.path.normcase(BASE):
        raise ValueError("That is this folder already")
    if not _data_info(folder):
        raise ValueError("No products or history found in that folder")
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    bk = os.path.join(BASE, "backup", f"{stamp}_before_merge")
    os.makedirs(bk, exist_ok=True)
    if os.path.exists(CONFIG_FILE):
        shutil.copy2(CONFIG_FILE, bk)
    if os.path.isdir(DATA):
        shutil.copytree(DATA, os.path.join(bk, "data"), dirs_exist_ok=True)
    src_cfg, src_data = os.path.join(folder, "config.yaml"), os.path.join(folder, "data")
    added = 0
    with CHECK_LOCK:                                   # don't let a check write while we merge
        cur = load_config()
        old = {}
        if os.path.exists(src_cfg):
            with open(src_cfg, encoding="utf-8") as f:
                old = yaml.safe_load(f) or {}
        have = {p["url"] for p in cur["products"]}
        cur["products"] = [p for p in cur["products"] if not _is_example(p)]
        for p in old.get("products") or []:
            if p.get("url") and p["url"] not in have and not _is_example(p):
                cur["products"].append(p)
                have.add(p["url"])
                added += 1
        for k, v in (old.get("notifications") or {}).items():      # keep any alert app that was set up there
            mine = cur["notifications"].get(k) or {}
            if isinstance(v, dict) and v.get("enabled") and not mine.get("enabled"):
                cur["notifications"][k] = v
        for k, v in (old.get("settings") or {}).items():
            cur["settings"].setdefault(k, v)
        save_config(cur)
        if os.path.isdir(src_data):
            _merge_csv(EVENTS_CSV, os.path.join(src_data, "events.csv"), EVENT_FIELDS,
                       lambda r: (r.get("date"), r.get("time"), r.get("url") or r.get("product"), r.get("new_status")))
            _merge_csv(CHECKS_CSV, os.path.join(src_data, "checks.csv"), CHECK_FIELDS,
                       lambda r: tuple(r.get(k, "") for k in CHECK_FIELDS))
            st_path = os.path.join(src_data, "state.json")
            if os.path.exists(st_path):
                with open(st_path, encoding="utf-8") as f:
                    theirs = json.load(f)
                mine = load_state()
                for url, v in theirs.items():
                    if url not in mine or str(v.get("last_checked", "")) > str(mine[url].get("last_checked", "")):
                        mine[url] = v
                save_state(mine)
    WAKE.set()
    try:
        build_reports(load_config())
    except Exception:
        pass
    info = _data_info(BASE) or {}
    return {"ok": True, "added": added, "products": info.get("products", 0), "events": info.get("events", 0),
            "checks": info.get("checks", 0), "backup": bk}


def restore_cli():
    print("\n  Looking for your old Restock Watch data (Downloads, Desktop, Documents, OneDrive)...\n")
    found = find_old_data()
    cur = _data_info(BASE) or {"products": 0, "events": 0}
    print(f"  This folder now has: {cur['products']} products, {cur['events']} stock changes\n")
    if not found:
        print("  No other Restock Watch folders with data were found.")
        print("  If you know where the old folder is, run:  python stock_tracker.py restore \"C:\\path\\to\\old\\folder\"\n")
        return
    for i, f in enumerate(found, 1):
        print(f"  {i}. {f['path']}")
        print(f"     {f['products']} products, {f['events']} stock changes, {f['checks']} checks  - last used {f['last_used']}\n")
    print("  Bringing data back MERGES it: products and history from that folder are added to this one.")
    print("  Nothing here is deleted, and a backup is made first.\n")
    choice = input("  Type the number to bring back (or press Enter to cancel): ").strip()
    if not choice.isdigit() or not 1 <= int(choice) <= len(found):
        print("  Cancelled - nothing changed.")
        return
    r = restore_old_data(found[int(choice) - 1]["path"])
    print(f"\n  Done! Added {r['added']} products. This folder now has {r['products']} products, "
          f"{r['events']} stock changes and {r['checks']} checks.")
    print(f"  (What was here before is saved in {r['backup']})\n")


# --------------------------------------------------------------------------- ports + web dashboard
def port_is_free(port, host="0.0.0.0"):
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if os.name != "nt":   # ignore ports that are only closing (TIME_WAIT); never used on Windows
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind((host, port))
            return True
        except OSError:
            return False


def find_free_port(start=8787, tries=200):
    for p in range(start, start + tries):
        if port_is_free(p):
            return p
    raise RuntimeError(f"No free port found between {start} and {start + tries}")


def show_ports(start, end):
    busy = [p for p in range(start, end + 1) if not port_is_free(p)]
    print(f"Ports {start}-{end}:")
    print(f"  BUSY ({len(busy)}): {', '.join(map(str, busy)) or 'none'}")
    free = next((p for p in range(start, end + 1) if p not in busy), None)
    print(f"  First FREE port: {free}")
    print("\nTo see every busy port and which program uses it:")
    if os.name == "nt":
        print("  netstat -ano | findstr LISTENING        (last column = PID)")
        print("  tasklist /FI \"PID eq <PID>\"             (shows the program name)")
    else:
        print("  lsof -iTCP -sTCP:LISTEN -P -n    or    ss -ltnp")


_HIST = {"mtime": None, "data": {}}
SECRET_KEYS = ("bot_token", "apikey", "password", "auth_token", "webhook_url")
CHANNEL_FIELDS = {
    "desktop": [],
    "telegram": ["bot_token", "chat_id"],
    "whatsapp": ["provider", "phone", "apikey", "account_sid", "auth_token", "from_number", "to_number"],
    "google_chat": ["webhook_url"],
    "email": ["smtp_server", "smtp_port", "username", "password", "to"],
    "ntfy": ["topic"],
    "discord": ["webhook_url"],
    "slack": ["webhook_url"],
}


def price_history():
    """Last 40 prices per product, read from checks.csv (cached until the file changes)."""
    try:
        mt = os.path.getmtime(CHECKS_CSV)
    except OSError:
        return {}
    if _HIST["mtime"] != mt:
        h = {}
        for row in read_csv(CHECKS_CSV)[-20000:]:
            v = to_num(row.get("price"))
            if v:
                h.setdefault(row["url"], []).append(v)
        _HIST.update(mtime=mt, data={k: v[-40:] for k, v in h.items()})
    return _HIST["data"]


def api_state():
    cfg, state, hist = load_config(), load_state(), price_history()
    products = []
    for p in cfg["products"]:
        s = state.get(p["url"], {})
        due_in = None
        if s.get("last_ts"):
            due_in = max(0, int(s["last_ts"] + interval_for(cfg, s.get("status")) - time.time()))
        products.append({"name": p.get("name") or s.get("title") or "", "url": p["url"], "site": site_name(p["url"]),
                         "status": s.get("status", "NOT CHECKED"), "since": s.get("since", ""),
                         "price": s.get("price", ""), "last_checked": s.get("last_checked", ""),
                         "target": p.get("target_price"), "detail": s.get("detail", ""), "image": s.get("image", ""),
                         "fails": s.get("fails", 0), "active": p.get("active", True), "history": hist.get(p["url"], []),
                         "due_in": due_in})
    n = cfg["notifications"]
    channels = [k for k in CHANNELS if (n.get(k) or ({"enabled": True} if k == "desktop" else {})).get("enabled")]
    nxt = PROGRESS.get("next_check")
    active = {p["url"] for p in cfg["products"]}
    events = [e for e in reversed(read_csv(EVENTS_CSV)[-400:]) if e.get("url") in active][:80]
    today = now(cfg).strftime("%Y-%m-%d")
    return {"app": "restock-watch", "version": APP_VERSION, "folder": BASE, "other_copies": OTHER_COPIES,
            "products": products, "events": events,
            "changes_today": sum(1 for e in events if e["date"] == today and not e["event"].startswith("FIRST")),
            "progress": {k: PROGRESS[k] for k in ("running", "done", "total", "current", "last_run")},
            "next_check_in": max(0, int(nxt - time.time())) if nxt else None,
            "every_minutes": cfg["settings"].get("check_every_minutes", 5),
            "fast_minutes": cfg["settings"].get("fast_check_minutes", 2),
            "now": now(cfg).strftime("%Y-%m-%d %H:%M:%S"), "channels": channels}


def api_settings_get():
    cfg = load_config()
    out = {}
    for ch, fields in CHANNEL_FIELDS.items():
        c = cfg["notifications"].get(ch) or {}
        item = {"enabled": bool(c.get("enabled", ch == "desktop"))}
        for f in fields:
            v = c.get(f, "")
            if f in SECRET_KEYS:
                v = str(v or "")
                item[f] = ""
                item[f + "_saved"] = ("•••• " + v[-4:]) if v and "Your" not in v and "XXXX" not in v and len(v) > 6 else ""
            else:
                item[f] = ", ".join(v) if isinstance(v, list) else ("" if v is None else str(v))
        out[ch] = item
    s = cfg["settings"]
    return {"channels": out, "check_every_minutes": s.get("check_every_minutes", 5),
            "fast_check_minutes": s.get("fast_check_minutes", 2), "daily_summary_time": s.get("daily_summary_time", ""),
            "start_with_windows": bool(s.get("start_with_windows", True)), "windows": os.name == "nt",
            "port": int(s.get("web_port", 8787))}


def api_settings_save(body):
    cfg = load_config()
    if "channel" in body:
        ch = body["channel"]
        if ch not in CHANNEL_FIELDS:
            raise ValueError("unknown channel")
        c = dict(cfg["notifications"].get(ch) or {})
        c["enabled"] = bool(body.get("enabled"))
        for f in CHANNEL_FIELDS[ch]:
            if f not in body:
                continue
            v = str(body[f]).strip()
            if f in SECRET_KEYS and not v:
                continue          # empty = keep the saved secret
            if f == "to":
                c[f] = [x.strip() for x in v.split(",") if x.strip()]
            elif f == "smtp_port":
                c[f] = int(v or 587)
            else:
                c[f] = v
        if ch == "whatsapp":
            c.setdefault("provider", "callmebot")
        if ch == "email":
            c.setdefault("smtp_server", "smtp.gmail.com")
            c.setdefault("smtp_port", 587)
            c.setdefault("attach_reports", True)
            if not c.get("to") and c.get("username"):
                c["to"] = [c["username"]]
        cfg["notifications"][ch] = c
    for k in ("check_every_minutes", "fast_check_minutes"):
        if k in body:
            cfg["settings"][k] = max(1.0, float(body[k]))
    if "start_with_windows" in body:
        cfg["settings"]["start_with_windows"] = bool(body["start_with_windows"])
        set_autostart(cfg["settings"]["start_with_windows"])
    if "daily_summary_time" in body:
        cfg["settings"]["daily_summary_time"] = body["daily_summary_time"] or None
    save_config(cfg)
    WAKE.set()


SAMPLE_ALERT = {"product": "Test product - Restock Watch is connected", "site": "Amazon", "price": "₹1,999",
                "url": "https://www.amazon.in/", "image": "", "previous": "This is a test. Real alerts look like this."}


OTHER_COPIES = []
EXIT_PORT_BUSY = 3          # 2_START.bat waits and tries again on this code


def startup_folder():
    return os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")


def set_autostart(enabled):
    """Windows: start Restock Watch (minimised, no browser pop-up) every time you sign in.
    Uses a tiny launcher in the Startup folder that runs 2_START.bat from THIS folder."""
    if os.name != "nt":
        return False
    folder = startup_folder()
    vbs = os.path.join(folder, "Restock Watch.vbs")
    old_lnk = os.path.join(folder, "Restock Watch.lnk")      # from an earlier version - avoid starting twice
    try:
        if os.path.exists(old_lnk):
            os.remove(old_lnk)
        if not enabled:
            if os.path.exists(vbs):
                os.remove(vbs)
            return False
        os.makedirs(folder, exist_ok=True)
        bat = os.path.join(BASE, "2_START.bat")
        text = ('Set sh = CreateObject("WScript.Shell")\r\n'
                f'sh.CurrentDirectory = "{BASE}"\r\n'
                f'sh.Run """{bat}"" silent", 7, False\r\n')
        current = ""
        if os.path.exists(vbs):
            try:
                with open(vbs, encoding="utf-16") as f:
                    current = f.read()
            except (UnicodeError, OSError):
                pass
        if current.replace("\r\n", "\n") != text.replace("\r\n", "\n"):
            with open(vbs, "w", encoding="utf-16", newline="") as f:   # UTF-16 so any folder name works
                f.write(text)
        return True
    except OSError as e:
        print("  Could not set up start-with-Windows:", e)
        return False


def watch_other_copies(my_port):
    """Every 2 minutes, look for another Restock Watch running from a DIFFERENT folder (it would keep its own
    separate products and history, which is confusing). The dashboard shows a warning if one is found."""
    local = requests.Session()
    local.trust_env = False
    time.sleep(5)            # let any copy that is starting at the same moment finish starting
    while True:
        found = []
        for port in range(8780, 8811):
            if port == my_port or port_is_free(port):
                continue
            try:
                d = local.get(f"http://127.0.0.1:{port}/api/state", timeout=1.5).json()
                if "products" in d and os.path.normcase(str(d.get("folder") or "?")) != os.path.normcase(BASE):
                    found.append({"port": port, "folder": d.get("folder") or "an older version (folder unknown)",
                                  "products": len(d["products"])})
            except Exception:
                pass
        OTHER_COPIES[:] = found
        if found:
            for f in found:
                print(f"  !! Another Restock Watch is running from {f['folder']} (port {f['port']}). "
                      f"Close it so all your products stay in one place.")
        time.sleep(60)


def stop_old_copy(port):
    """If another Restock Watch (any version, any folder) is using our port, ask it to close so this one takes over.
    Its products and history stay on disk - only the program is closed."""
    local = requests.Session()
    local.trust_env = False   # never route localhost through a proxy
    try:
        d = local.get(f"http://127.0.0.1:{port}/api/state", timeout=3).json()
    except Exception:
        return False
    if "products" not in d:
        return False          # something else, not Restock Watch
    try:
        local.post(f"http://127.0.0.1:{port}/api/stop", timeout=3)
    except Exception:
        pass
    for _ in range(15):
        time.sleep(0.3)
        if port_is_free(port):
            other = d.get("folder")
            print(f"  Closed the Restock Watch that was already running on port {port}"
                  + (f" (from {other})." if other and os.path.normcase(other) != os.path.normcase(BASE) else "."))
            return True
    print(f"\n  !! An older Restock Watch is running on port {port} and could not be closed automatically.\n"
          f"  !! Close its black window (or restart the PC) and start again.\n")
    return False


def who_uses_port(port):
    """Name of the program listening on a port (Windows), so the message can say what to close."""
    if os.name != "nt":
        return ""
    import subprocess
    try:
        out = subprocess.run(["netstat", "-ano", "-p", "tcp"], capture_output=True, text=True, timeout=10).stdout
        pid = next((ln.split()[-1] for ln in out.splitlines()
                    if f":{port} " in ln and "LISTENING" in ln.upper()), None)
        if not pid:
            return ""
        t = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                           capture_output=True, text=True, timeout=10).stdout.strip()
        return f"{t.split(',')[0].strip(chr(34))} (PID {pid})" if t and "," in t else f"PID {pid}"
    except Exception:
        return ""


def run_web(port, host):
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    ui_file = os.path.join(BASE, "dashboard.html")

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def send(self, body, ctype="application/json", name=None, code=200):
            if isinstance(body, (dict, list)):
                body = json.dumps(body, ensure_ascii=False)
            data = body if isinstance(body, bytes) else body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Cache-Control", "no-store")
            if name:
                self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):
            path = self.path.split("?")[0]
            files = {"/excel": (XLSX_FILE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                     "/pdf": (PDF_FILE, "application/pdf"), "/csv": (EVENTS_CSV, "text/csv; charset=utf-8")}
            if path == "/api/state":
                return self.send(api_state())
            if path == "/api/settings":
                return self.send(api_settings_get())
            if path == "/api/olddata":
                return self.send({"found": [{k: v for k, v in f.items() if k != "_ts"} for f in find_old_data()]})
            if path == "/api/report":
                return self.send({"items": product_timelines(load_config(), include_removed="removed=1" in self.path)})
            if path == "/download":
                from urllib.parse import parse_qs, urlsplit
                q = parse_qs(urlsplit(self.path).query)
                try:
                    data, name, ctype = export_products(load_config(), q.get("url", []), (q.get("fmt") or ["xlsx"])[0])
                except ValueError as e:
                    return self.send(str(e), "text/plain; charset=utf-8", code=400)
                self.send_response(200)
                self.send_header("Content-Type", ctype)
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Disposition", f"attachment; filename=\"{name.encode('ascii', 'ignore').decode()}\"; "
                                 f"filename*=UTF-8''{quote(name)}")
                self.end_headers()
                return self.wfile.write(data)
            if path in files:
                fpath, ctype = files[path]
                if path != "/csv":
                    build_reports(load_config())
                if not os.path.exists(fpath):
                    return self.send("Nothing recorded yet - add a product first.", "text/plain; charset=utf-8")
                with open(fpath, "rb") as f:
                    return self.send(f.read(), ctype, os.path.basename(fpath))
            with open(ui_file, "rb") as f:
                self.send(f.read(), "text/html; charset=utf-8")

        def do_POST(self):
            n = int(self.headers.get("Content-Length", 0) or 0)
            try:
                body = json.loads(self.rfile.read(n) or b"{}")
            except ValueError:
                body = {}
            try:
                return self.route(body)
            except Exception as e:
                return self.send({"error": str(e)}, code=400)

        def route(self, body):
            cfg = load_config()
            if self.path == "/api/add":
                url = (body.get("url") or "").strip()
                m = re.search(r"https?://\S+", url)
                if not m:
                    return self.send({"error": "Paste the full product link (it starts with https://)"}, code=400)
                url = m.group(0)
                if any(p["url"] == url for p in cfg["products"]):
                    return self.send({"error": "This product is already being watched"}, code=400)
                item = {"name": (body.get("name") or "").strip(), "url": url}
                if to_num(str(body.get("target") or "")):
                    item["target_price"] = to_num(str(body["target"]))
                cfg["products"].append(item)
                save_config(cfg)
                WAKE.set()
            elif self.path == "/api/update":
                for p in cfg["products"]:
                    if p["url"] == body.get("url"):
                        if "active" in body:
                            p["active"] = bool(body["active"])
                        if "name" in body:
                            p["name"] = str(body["name"]).strip()
                        if "target" in body:
                            t = to_num(str(body["target"] or ""))
                            if t:
                                p["target_price"] = t
                            else:
                                p.pop("target_price", None)
                save_config(cfg)
            elif self.path == "/api/remove":
                gone = [p for p in cfg["products"] if p["url"] == body.get("url")]
                cfg["products"] = [p for p in cfg["products"] if p["url"] != body.get("url")]
                save_config(cfg)
                if gone and gone[0].get("name"):          # remember your name for it, for the hidden history
                    st_ = load_state()
                    st_.setdefault(gone[0]["url"], {})["name"] = gone[0]["name"]
                    save_state(st_)
            elif self.path == "/api/check":
                if PROGRESS["running"]:
                    return self.send({"ok": True, "already": True})
                FORCE.set()
                WAKE.set()
            elif self.path == "/api/settings":
                api_settings_save(body)
            elif self.path == "/api/test":
                ch = body.get("channel")
                sent, errs = notify_all(cfg, "✅ Test alert from Restock Watch", "", alert=dict(SAMPLE_ALERT, when=now(cfg).strftime("%d %b %Y, %I:%M:%S %p")), only=ch)
                if errs:
                    return self.send({"error": errs.get(ch) or next(iter(errs.values()))}, code=400)
                if ch == "desktop":
                    try:
                        import plyer  # noqa
                    except ImportError:
                        return self.send({"error": "Desktop pop-ups need: pip install plyer"}, code=400)
                return self.send({"ok": True, "sent": sent})
            elif self.path == "/api/stop":
                self.send({"ok": True})
                print("\n  Stopped because a newer copy of Restock Watch was started.\n")
                threading.Timer(0.3, lambda: os._exit(0)).start()
                return
            elif self.path == "/api/restore":
                return self.send(restore_old_data(body.get("path", "")))
            elif self.path == "/api/telegram/find":
                token = (body.get("bot_token") or "").strip() or str(
                    (cfg["notifications"].get("telegram") or {}).get("bot_token", ""))
                return self.send(telegram_find_chat(token))
            self.send({"ok": True})

    if not port_is_free(port, host):
        stop_old_copy(port)
    if not port_is_free(port, host):
        owner = who_uses_port(port)
        print(f"\n  !! Port {port} is being used by another program{': ' + owner if owner else ''}.\n"
              f"  !! Restock Watch always runs on http://localhost:{port}, so that program must be closed first.\n"
              f"  !!   - Close that program (or restart the PC), then double-click 2_START.bat again.\n"
              f"  !!   - Or, to use a different port, change  web_port: {port}  in config.yaml.\n")
        sys.exit(EXIT_PORT_BUSY)
    srv = ThreadingHTTPServer((host, port), Handler)
    threading.Thread(target=run_forever, daemon=True).start()
    threading.Thread(target=watch_other_copies, args=(port,), daemon=True).start()
    link = f"http://localhost:{port}"
    print(f"\n  Restock Watch is running:  {link}\n  Keep this window open (minimise it). Ctrl+C to stop.\n")
    if os.environ.get("NO_BROWSER") != "1":
        try:
            import webbrowser
            threading.Timer(1.2, lambda: webbrowser.open(link)).start()
        except Exception:
            pass
    cfg = load_config()
    if os.name == "nt":
        on = set_autostart(cfg["settings"].get("start_with_windows", True))
        print("  Starts automatically with Windows: " + ("ON" if on else "OFF") + "  (change it on the Alerts page)\n")
    if os.environ.get("NO_BROWSER") == "1":
        try:
            send_desktop({}, "Restock Watch is running", f"Watching your products. Dashboard: {link}", {})
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("  Stopped.")
        os._exit(0)


def main():
    ap = argparse.ArgumentParser(description="Online shopping stock tracker")
    sub = ap.add_subparsers(dest="cmd")
    w = sub.add_parser("web", help="web dashboard + tracking on http://localhost:8787")
    w.add_argument("--port", type=int, default=None, help="port number (default from config, else 8787)")
    w.add_argument("--host", default=None, help="127.0.0.1 = only this PC, 0.0.0.0 = phone on same Wi-Fi too")
    au = sub.add_parser("autostart", help="start with Windows: on / off")
    au.add_argument("onoff", choices=["on", "off"])
    rs = sub.add_parser("restore", help="find and bring back products + history from an older folder")
    rs.add_argument("folder", nargs="?", default=None)
    pp = sub.add_parser("ports", help="show busy/free ports")
    pp.add_argument("--from", dest="start", type=int, default=8000)
    pp.add_argument("--to", dest="end", type=int, default=9000)
    a = sub.add_parser("add", help="add a product URL")
    a.add_argument("url")
    a.add_argument("--name", default="")
    a.add_argument("--target-price", type=float, default=None)
    r = sub.add_parser("remove", help="remove a product by number (see list)")
    r.add_argument("number", type=int)
    sub.add_parser("list")
    sub.add_parser("once")
    sub.add_parser("run")
    sub.add_parser("report")
    sub.add_parser("summary")
    sub.add_parser("test-notify")
    args = ap.parse_args()
    cfg = load_config()

    if args.cmd == "web":
        s = cfg["settings"]
        run_web(args.port or int(s.get("web_port", 8787)), args.host or s.get("web_host", "127.0.0.1"))
    elif args.cmd == "autostart":
        cfg["settings"]["start_with_windows"] = args.onoff == "on"
        save_config(cfg)
        ok = set_autostart(args.onoff == "on")
        print("Start with Windows is now", "ON" if ok else "OFF")
    elif args.cmd == "restore":
        if args.folder:
            r = restore_old_data(args.folder)
            print(f"Done! Added {r['added']} products. Now {r['products']} products, {r['events']} stock changes, "
                  f"{r['checks']} checks. Backup: {r['backup']}")
        else:
            restore_cli()
    elif args.cmd == "ports":
        show_ports(args.start, args.end)
    elif args.cmd == "add":
        item = {"name": args.name or "", "url": args.url.strip()}
        if args.target_price:
            item["target_price"] = args.target_price
        cfg["products"].append(item)
        save_config(cfg)
        print(f"Added: {args.name or args.url}")
    elif args.cmd == "remove":
        p = cfg["products"].pop(args.number - 1)
        save_config(cfg)
        print("Removed:", p.get("name") or p["url"])
    elif args.cmd == "list":
        state = load_state()
        for i, p in enumerate(cfg["products"], 1):
            s = state.get(p["url"], {})
            print(f"{i}. [{s.get('status', 'not checked')}] {p.get('name') or s.get('title') or ''}  {p['url']}")
    elif args.cmd == "once":
        check_all(cfg)
    elif args.cmd == "run":
        print("Tracking... (Ctrl+C to stop)")
        run_forever()
    elif args.cmd == "report":
        x, p = build_reports(cfg)
        print("Saved:", x, "and", p)
    elif args.cmd == "summary":
        daily_summary(cfg)
    elif args.cmd == "test-notify":
        sent, errs = notify_all(cfg, "✅ Test alert from Restock Watch", "",
                                alert=dict(SAMPLE_ALERT, when=now(cfg).strftime("%d %b %Y, %I:%M:%S %p")))
        print("Sent to:", ", ".join(sent) or "nothing - enable a channel on the dashboard's Alerts page")
        for k, v in errs.items():
            print(f"  {k}: {v}")
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
