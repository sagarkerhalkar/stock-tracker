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
        a, b = text.index("products:"), text.index("# ----", text.index("products:"))
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(text[:a] + "products: []\n\n" + text[b:])
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
            "site": site, "title": title[:150]}


# --------------------------------------------------------------------------- notifications
def notify_all(cfg, subject, message, attach=None):
    n = cfg["notifications"]
    sent = []
    for name, fn in [("telegram", send_telegram), ("whatsapp", send_whatsapp), ("google_chat", send_google_chat),
                     ("email", send_email), ("discord", send_discord), ("slack", send_slack), ("ntfy", send_ntfy),
                     ("desktop", send_desktop)]:
        c = n.get(name) or ({"enabled": True} if name == "desktop" else {})
        if not c.get("enabled"):
            continue
        try:
            if name == "email":
                fn(c, subject, message, attach)
            else:
                fn(c, subject, message)
            sent.append(name)
        except Exception as e:
            print(f"   ! {name} failed: {e}")
    return sent


def send_desktop(c, subject, msg):
    """Pop-up notification on this computer (Windows/Mac/Linux). Needs: pip install plyer"""
    try:
        from plyer import notification
    except ImportError:
        return
    notification.notify(title=subject[:63], message=msg[:250], app_name="Stock Tracker", timeout=15)


def send_telegram(c, subject, msg):
    r = requests.post(f"https://api.telegram.org/bot{c['bot_token']}/sendMessage",
                      json={"chat_id": c["chat_id"], "text": f"{subject}\n\n{msg}", "disable_web_page_preview": False},
                      timeout=20)
    r.raise_for_status()


def send_whatsapp(c, subject, msg):
    text = f"*{subject}*\n\n{msg}"
    provider = c.get("provider", "callmebot")
    if provider == "callmebot":
        r = requests.get(f"https://api.callmebot.com/whatsapp.php?phone={quote(str(c['phone']))}"
                         f"&text={quote(text)}&apikey={quote(str(c['apikey']))}", timeout=30)
        r.raise_for_status()
    elif provider == "twilio":
        r = requests.post(f"https://api.twilio.com/2010-04-01/Accounts/{c['account_sid']}/Messages.json",
                          data={"From": f"whatsapp:{c['from_number']}", "To": f"whatsapp:{c['to_number']}", "Body": text},
                          auth=(c["account_sid"], c["auth_token"]), timeout=30)
        r.raise_for_status()
    else:
        raise ValueError(f"unknown whatsapp provider {provider}")


def send_google_chat(c, subject, msg):
    requests.post(c["webhook_url"], json={"text": f"*{subject}*\n{msg}"}, timeout=20).raise_for_status()


def send_discord(c, subject, msg):
    requests.post(c["webhook_url"], json={"content": f"**{subject}**\n{msg}"}, timeout=20).raise_for_status()


def send_slack(c, subject, msg):
    requests.post(c["webhook_url"], json={"text": f"*{subject}*\n{msg}"}, timeout=20).raise_for_status()


def send_ntfy(c, subject, msg):
    server = c.get("server", "https://ntfy.sh").rstrip("/")
    requests.post(f"{server}/{c['topic']}", data=msg.encode("utf-8"),
                  headers={"Title": subject.encode("ascii", "ignore").decode() or "Stock alert", "Priority": "high"},
                  timeout=20).raise_for_status()


def send_email(c, subject, msg, attach=None):
    m = MIMEMultipart()
    m["From"], m["To"], m["Subject"] = c["username"], ", ".join(c["to"]) if isinstance(c["to"], list) else c["to"], subject
    m.attach(MIMEText(msg, "plain", "utf-8"))
    for path in attach or []:
        if path and os.path.exists(path):
            with open(path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(path))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(path)}"'
            m.attach(part)
    with smtplib.SMTP(c.get("smtp_server", "smtp.gmail.com"), int(c.get("smtp_port", 587)), timeout=30) as s:
        s.starttls()
        s.login(c["username"], c["password"])
        s.send_message(m)


# --------------------------------------------------------------------------- reports
def availability_periods(events, cfg):
    """Turn events into rows: product | became available | went out of stock | how long it stayed."""
    periods, open_ = [], {}
    fmt = "%Y-%m-%d %H:%M:%S"
    for e in events:
        key = e["product"]
        ts = f"{e['date']} {e['time']}"
        if e["new_status"] == IN and key not in open_:
            open_[key] = (ts, e.get("price", ""), e["site"])
        elif e["new_status"] == OUT and key in open_:
            start, price, site = open_.pop(key)
            dur = datetime.strptime(ts, fmt) - datetime.strptime(start, fmt)
            periods.append([key, site, start, ts, fmt_dur(dur.total_seconds()), price])
    current = now(cfg).replace(tzinfo=None)
    for key, (start, price, site) in open_.items():
        dur = current - datetime.strptime(start, fmt)
        periods.append([key, site, start, "still in stock", fmt_dur(dur.total_seconds()), price])
    return periods


def fmt_dur(sec):
    sec = int(max(sec, 0))
    d, r = divmod(sec, 86400)
    h, r = divmod(r, 3600)
    m = r // 60
    return (f"{d}d " if d else "") + f"{h}h {m}m"


def build_reports(cfg):
    events = read_csv(EVENTS_CSV)
    checks = read_csv(CHECKS_CSV)
    state = load_state()
    periods = availability_periods(events, cfg)
    build_excel(cfg, events, checks, state, periods)
    build_pdf(cfg, events, state, periods)
    return XLSX_FILE, PDF_FILE


def build_excel(cfg, events, checks, state, periods):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="1F3A5F")
    green = PatternFill("solid", fgColor="D4EDDA")
    red = PatternFill("solid", fgColor="F8D7DA")
    grey = PatternFill("solid", fgColor="EEEEEE")

    def sheet(ws, headers, rows, status_col=None, widths=None):
        ws.append(headers)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), head_fill
            c.alignment = Alignment(vertical="center")
        for row in rows:
            ws.append(row)
            if status_col is not None:
                v = str(row[status_col])
                fill = green if (IN in v or "BACK" in v) else red if (OUT in v or "OUT" in v) else grey
                ws.cell(ws.max_row, status_col + 1).fill = fill
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, h in enumerate(headers, 1):
            w = (widths or {}).get(h) or min(max([len(str(h))] + [len(str(r[i - 1])) for r in rows[:500]]) + 2, 60)
            ws.column_dimensions[get_column_letter(i)].width = w

    ws = wb.active
    ws.title = "Current Status"
    rows = []
    for p in cfg["products"]:
        s = state.get(p["url"], {})
        rows.append([p.get("name") or s.get("title", ""), site_name(p["url"]), s.get("status", "not checked"),
                     s.get("since", ""), s.get("price", ""), s.get("last_checked", ""), p["url"]])
    sheet(ws, ["Product", "Site", "Status", "Since", "Price", "Last checked", "URL"], rows, status_col=2,
          widths={"URL": 60, "Product": 40})

    sheet(wb.create_sheet("Stock Events"), [f.replace("_", " ").title() for f in EVENT_FIELDS],
          [[e.get(f, "") for f in EVENT_FIELDS] for e in reversed(events)], status_col=4,
          widths={"Url": 50, "Product": 40})
    sheet(wb.create_sheet("Availability Periods"),
          ["Product", "Site", "Came in stock", "Went out of stock", "Stayed available", "Price"], periods,
          widths={"Product": 40})
    sheet(wb.create_sheet("All Checks"), [f.replace("_", " ").title() for f in CHECK_FIELDS],
          [[c.get(f, "") for f in CHECK_FIELDS] for c in reversed(checks[-5000:])], status_col=4,
          widths={"Url": 50, "Product": 40, "Detail": 40})
    wb.save(XLSX_FILE)


def build_pdf(cfg, events, state, periods):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    st = getSampleStyleSheet()
    small = st["BodyText"].clone("small", fontSize=8, leading=10)
    doc = SimpleDocTemplate(PDF_FILE, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm, title="Stock Availability Report")
    rupee = lambda s: str(s).replace("₹", "Rs ")
    P = lambda s: Paragraph(rupee(s).replace("&", "&amp;").replace("<", "&lt;"), small)

    def table(headers, rows, widths, status_col=None):
        data = [[Paragraph(f'<font color="white"><b>{h}</b></font>', small) for h in headers]] + \
               [[P(c) for c in r] for r in rows]
        t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
        style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                 ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                 ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
                 ("VALIGN", (0, 0), (-1, -1), "TOP")]
        for i, r in enumerate(rows, 1):
            if status_col is not None:
                v = str(r[status_col])
                c = "#D4EDDA" if (IN in v or "BACK" in v) else "#F8D7DA" if (OUT in v or "OUT" in v) else "#EEEEEE"
                style.append(("BACKGROUND", (status_col, i), (status_col, i), colors.HexColor(c)))
        t.setStyle(TableStyle(style))
        return t

    parts = [Paragraph("Stock Availability Report", st["Title"]),
             Paragraph(f"Generated {now(cfg).strftime('%d %b %Y, %I:%M %p')} ({cfg['settings'].get('timezone', 'Asia/Kolkata')})",
                       st["Normal"]), Spacer(1, 6 * mm), Paragraph("Current status", st["Heading2"])]
    rows = []
    for p in cfg["products"]:
        s = state.get(p["url"], {})
        rows.append([p.get("name") or s.get("title", ""), site_name(p["url"]), s.get("status", "not checked"),
                     s.get("since", ""), s.get("price", ""), s.get("last_checked", "")])
    parts.append(table(["Product", "Site", "Status", "Since", "Price", "Last checked"], rows or [["-"] * 6],
                       [90, 30, 32, 40, 30, 40], status_col=2))
    parts += [Spacer(1, 6 * mm), Paragraph("Availability periods (when it was in stock and for how long)", st["Heading2"])]
    parts.append(table(["Product", "Site", "Came in stock", "Went out of stock", "Stayed", "Price"],
                       periods or [["No periods yet", "", "", "", "", ""]], [90, 28, 40, 40, 32, 30]))
    parts += [Spacer(1, 6 * mm), Paragraph("Stock change history (newest first)", st["Heading2"])]
    erows = [[e["date"], e["time"], e["product"], e["site"], e["event"], e.get("price", "")] for e in reversed(events[-300:])]
    parts.append(table(["Date", "Time", "Product", "Site", "Event", "Price"], erows or [["No changes yet"] + [""] * 5],
                       [26, 22, 110, 28, 45, 30], status_col=4))
    doc.build(parts)


# --------------------------------------------------------------------------- main logic
import threading
from concurrent.futures import ThreadPoolExecutor

CHECK_LOCK = threading.Lock()
WAKE = threading.Event()
PROGRESS = {"running": False, "done": 0, "total": 0, "current": "", "next_check": None, "last_run": None}


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
            time.sleep(random.uniform(*s.get("delay_between_products_seconds", [3, 6])))
    return out


def check_all(cfg, send=True):
    if not CHECK_LOCK.acquire(blocking=False):
        print("A check is already running - skipped.")
        return False
    try:
        return _check_all(cfg, send)
    finally:
        PROGRESS.update(running=False, current="", last_run=now(cfg).strftime("%Y-%m-%d %H:%M:%S"))
        CHECK_LOCK.release()


def _check_all(cfg, send):
    s = cfg["settings"]
    products = [p for p in cfg["products"] if p.get("active", True)]
    if not products:
        print("No products yet. Add one:  python stock_tracker.py add <url> --name \"Product\"")
        return True
    PROGRESS.update(running=True, done=0, total=len(products), current="")
    groups = {}
    for p in products:
        groups.setdefault(site_name(p["url"]), []).append(p)
    with ThreadPoolExecutor(max_workers=min(6, len(groups))) as ex:
        results = [r for chunk in ex.map(lambda g: _check_site_group(g, s), groups.values()) for r in chunk]

    state = load_state()
    changes = []
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
        entry.update({"last_checked": f"{date} {clock}", "title": res.get("title") or prev.get("title", ""),
                      "detail": res["detail"]})
        if res["price"]:
            entry["price"] = res["price"]
        if res["status"] == UNKNOWN:
            entry["fails"] = prev.get("fails", 0) + 1
            if entry["fails"] == int(s.get("alert_after_failures", 6)) and send:
                notify_all(cfg, "⚠️ Stock tracker can't read a page",
                           f"{name}\nCould not read the page {entry['fails']} times in a row ({res['detail']}).\n{p['url']}")
        else:
            entry["fails"] = 0
            if res["status"] != old:
                if old is None:
                    event = f"FIRST CHECK - {res['status']}"
                elif res["status"] == IN:
                    event = "BACK IN STOCK"
                else:
                    event = "WENT OUT OF STOCK"
                entry.update({"status": res["status"], "since": f"{date} {clock}"})
                append_csv(EVENTS_CSV, EVENT_FIELDS, {"date": date, "time": clock, "product": name, "site": res["site"],
                                                      "event": event, "old_status": old or "-", "new_status": res["status"],
                                                      "price": res["price"], "url": p["url"], "note": res["detail"]})
                changes.append((name, event, res, p, date, clock, prev))
            target = p.get("target_price")
            pnum = to_num(res["price"])
            if target and pnum and pnum <= float(target) and res["status"] == IN and not prev.get("target_hit"):
                entry["target_hit"] = True
                changes.append((name, f"PRICE DROP to {res['price']} (target ₹{target})", res, p, date, clock, prev))
            elif target and pnum and pnum > float(target):
                entry["target_hit"] = False
        state[p["url"]] = entry
    save_state(state)

    xlsx, pdf = build_reports(cfg)
    for name, event, res, p, date, clock, prev in changes:
        first = event.startswith("FIRST")
        if first and not s.get("notify_on_first_check", True):
            continue
        icon = "🟢" if IN in event or "BACK" in event else "🔴" if "OUT" in event else "💰" if "PRICE" in event else "ℹ️"
        when = datetime.strptime(f"{date} {clock}", "%Y-%m-%d %H:%M:%S").strftime("%d %b %Y, %I:%M:%S %p")
        lines = [name, f"Site: {res['site']}", f"Status: {res['status']}", f"Time: {when}"]
        if res["price"]:
            lines.append(f"Price: {res['price']}")
        if prev.get("since") and not first and "PRICE" not in event:
            lines.append(f"Previous status: {prev.get('status')} since {prev['since']}")
        lines.append(p["url"])
        if send:
            title = {"FIRST CHECK - IN STOCK": "✅ AVAILABLE NOW", "FIRST CHECK - OUT OF STOCK": "⛔ NOT AVAILABLE (out of stock)",
                     "WENT OUT OF STOCK": "🔴 WENT OUT OF STOCK", "BACK IN STOCK": "🟢 BACK IN STOCK - available again!"}.get(
                event, f"{icon} {event}")
            sent = notify_all(cfg, title, "\n".join(lines),
                              attach=[xlsx, pdf] if cfg["notifications"].get("email", {}).get("attach_reports") else None)
            print(f"   -> notified: {', '.join(sent) or 'no channel enabled'}")
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
                     f"{p.get('name') or s.get('title', p['url'])[:50]} - {s.get('status', 'not checked')} "
                     f"{s.get('price', '')} (since {s.get('since', '-')})")
    xlsx, pdf = build_reports(cfg)
    notify_all(cfg, "📊 Daily stock summary", "\n".join(lines) or "No products", attach=[xlsx, pdf])


def run_forever(cfg_initial):
    last_summary_day = None
    while True:
        cfg = load_config()  # reload so new products/settings are picked up without restart
        try:
            check_all(cfg)
        except Exception as e:
            print("Error during check:", e)
        s = cfg["settings"]
        hhmm = s.get("daily_summary_time")
        t = now(cfg)
        if hhmm and t.strftime("%H:%M") >= hhmm and last_summary_day != t.date():
            last_summary_day = t.date()
            try:
                daily_summary(cfg)
            except Exception as e:
                print("Summary error:", e)
        mins = float(s.get("check_every_minutes", 15))
        wait = mins * 60 * random.uniform(0.85, 1.15)
        PROGRESS["next_check"] = time.time() + wait
        print(f"Next check in {wait / 60:.1f} min  (Ctrl+C to stop)\n")
        WAKE.wait(wait)   # the dashboard's "Check now" button wakes this up early
        WAKE.clear()


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


def api_state():
    cfg, state = load_config(), load_state()
    products = []
    for p in cfg["products"]:
        s = state.get(p["url"], {})
        products.append({"name": p.get("name") or s.get("title") or "", "url": p["url"], "site": site_name(p["url"]),
                         "status": s.get("status", "NOT CHECKED"), "since": s.get("since", ""),
                         "price": s.get("price", ""), "last_checked": s.get("last_checked", ""),
                         "target": p.get("target_price"), "detail": s.get("detail", ""),
                         "fails": s.get("fails", 0), "active": p.get("active", True)})
    n = cfg["notifications"]
    channels = [k.replace("_", " ").title() for k, v in n.items() if isinstance(v, dict) and v.get("enabled")]
    if "desktop" not in n:
        channels.insert(0, "Desktop")
    nxt = PROGRESS.get("next_check")
    return {"products": products, "events": list(reversed(read_csv(EVENTS_CSV)[-60:])),
            "progress": {k: PROGRESS[k] for k in ("running", "done", "total", "current", "last_run")},
            "next_check_in": max(0, int(nxt - time.time())) if nxt else None,
            "every_minutes": cfg["settings"].get("check_every_minutes", 15),
            "now": now(cfg).strftime("%Y-%m-%d %H:%M:%S"), "channels": channels}


def run_web(port, host):
    import threading
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
            if path in files:
                fpath, ctype = files[path]
                if path != "/csv":
                    build_reports(load_config())
                if not os.path.exists(fpath):
                    return self.send("Nothing recorded yet - add a product and run a check first.", "text/plain")
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
            cfg = load_config()
            if self.path == "/api/add":
                url = (body.get("url") or "").strip()
                if not url.startswith("http"):
                    return self.send({"error": "Paste the full product link, starting with https://"}, code=400)
                if any(p["url"] == url for p in cfg["products"]):
                    return self.send({"error": "This product is already on your list"}, code=400)
                item = {"name": (body.get("name") or "").strip(), "url": url}
                if to_num(str(body.get("target") or "")):
                    item["target_price"] = to_num(str(body["target"]))
                cfg["products"].append(item)
                save_config(cfg)
                WAKE.set()   # check the new product right away
            elif self.path == "/api/remove":
                cfg["products"] = [p for p in cfg["products"] if p["url"] != body.get("url")]
                save_config(cfg)
            elif self.path == "/api/check":
                if PROGRESS["running"]:
                    return self.send({"ok": True, "already": True})
                WAKE.set()
            elif self.path == "/api/test-notify":
                sent = notify_all(cfg, "✅ Stock tracker test", "Notifications are working!\n" +
                                  now(cfg).strftime("%d %b %Y, %I:%M %p"))
                return self.send({"ok": True, "sent": sent})
            self.send({"ok": True})

    if not port_is_free(port, host):
        new = find_free_port(port + 1)
        print(f"Port {port} is busy -> using free port {new} instead")
        port = new
    srv = ThreadingHTTPServer((host, port), Handler)
    threading.Thread(target=run_forever, args=(load_config(),), daemon=True).start()
    link = f"http://localhost:{port}"
    print(f"\nDashboard: {link}   (tracking runs in background, Ctrl+C to stop)\n")
    if os.environ.get("NO_BROWSER") != "1":
        try:
            import webbrowser
            threading.Timer(1.2, lambda: webbrowser.open(link)).start()
        except Exception:
            pass
    srv.serve_forever()


def main():
    ap = argparse.ArgumentParser(description="Online shopping stock tracker")
    sub = ap.add_subparsers(dest="cmd")
    w = sub.add_parser("web", help="web dashboard + tracking (choose your port)")
    w.add_argument("--port", type=int, default=None, help="port number (default from config, else 8787)")
    w.add_argument("--host", default=None, help="127.0.0.1 = only this PC, 0.0.0.0 = phone on same Wi-Fi too")
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
        run_forever(cfg)
    elif args.cmd == "report":
        x, p = build_reports(cfg)
        print("Saved:", x, "and", p)
    elif args.cmd == "summary":
        daily_summary(cfg)
    elif args.cmd == "test-notify":
        sent = notify_all(cfg, "✅ Stock tracker test", "If you can read this, notifications work!\n" +
                          now(cfg).strftime("%d %b %Y, %I:%M %p"))
        print("Sent to:", ", ".join(sent) or "nothing - enable a channel in config.yaml")
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
