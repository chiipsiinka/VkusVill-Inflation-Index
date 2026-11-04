import argparse
import csv
import html
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from bs4 import BeautifulSoup


BASE = "https://vkusvill.ru"
MSK = timezone(timedelta(hours=3))

STORE = {
    "shop_no": "7064",
    "addr": "Санкт-Петербург пр-т Медиков, д. 10 к. 1",
    "lat": "59.974710000000000",
    "lon": "30.316340000000000",
}
STORE_CHECK = "Медиков"

SLOT_SINCE = "11:00"
SLOT_UNTIL = "13:00"
SLOT_EXPIRES = "10:30"

SPREADSHEET_ID = "18ypB7rF_yeQK7BgJ_AfFQIYpjkam8N-q4o7nq0gZ4JM"
BASKET_SHEET = "Basket"
PRICES_SHEET = "Daily Prices"

HEADER = [
    "Date", "Item (EN)", "Regular price (RUB)", "Promo price (RUB)", "Item #",
    "Product ID", "Product name on site", "Pack size", "Price unit",
    "In stock", "Qty left", "Product URL", "Status",
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36")

DELAY = (2.0, 4.0)

DEBUG_DIR = "debug"


def new_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    })
    return s


def get(s, url, **kw):
    last = None
    for attempt in range(3):
        try:
            r = s.get(url, timeout=30, **kw)
            if r.status_code == 200:
                return r
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            last = str(e)
        time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"GET {url} failed: {last}")


def save_debug(name, text):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    with open(os.path.join(DEBUG_DIR, name), "w", encoding="utf-8") as f:
        f.write(text)


def describe(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    title = soup.title.get_text(strip=True) if soup.title else ""
    return f"{len(page_html)} chars, title={title[:80]!r}"


def ajax_post(s, path, data):
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": BASE,
        "Referer": BASE + "/",
        "X-Requested-With": "XMLHttpRequest",
    }
    r = s.post(BASE + path, data=data, headers=headers, timeout=30)
    short = re.sub(r"\s+", " ", r.text)[:300]
    print(f"POST {path}: HTTP {r.status_code}, {len(r.text)} chars: {short}", flush=True)
    save_debug(path.strip("/").replace("/", "_") + f"_{int(time.time() * 1000)}.txt", r.text)
    if r.status_code != 200:
        raise RuntimeError(f"POST {path} failed: HTTP {r.status_code}")
    return r.text


def read_sessid(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    tag = soup.find("input", {"id": "sessid"}) or soup.find("input", {"name": "sessid"})
    if not tag or not tag.get("value"):
        raise RuntimeError("sessid not found on the main page (blocked by anti-bot protection?)")
    return tag["value"]


def header_address(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    tag = soup.select_one(".js-header-address-text")
    return tag.get_text(" ", strip=True) if tag else ""


def select_store_and_slot(s, today):
    main_page = get(s, BASE + "/").text
    save_debug("main_page.html", main_page)
    print(f"main page: {describe(main_page)}, header address: {header_address(main_page)!r}", flush=True)
    sessid = read_sessid(main_page)

    ajax_post(s, "/ajax/address_way/service_apply.php", {
        "addr": STORE["addr"], "addr_id": "false",
        "lat": STORE["lat"], "lon": STORE["lon"],
        "shop_no": STORE["shop_no"], "type": "shops", "type_service": "3",
        "route_comment": "", "new_service": "", "new_slot": "",
        "sessid": sessid,
    })

    ajax_post(s, "/ajax/address_way/modals/delivery_slots/ds_loader.php", {
        "loadAssets": "N", "is_aw_open": "Y", "date_start": "",
        "time_start": "", "reset_preorder": "N", "mode": "",
    })

    slot_date = (today + timedelta(days=1)).strftime("%Y-%m-%d")
    ajax_post(s, "/ajax/address_way/service_apply.php", {
        "from_slots": "Y", "show_terms": "N", "setDefault": "N", "getServices": "N",
        "new_service[type_service]": "3", "new_service[id_service]": "3",
        "new_service[id_online_service]": "3", "new_service[chosen]": "1",
        "new_service[other_parameters]": "", "new_service[price]": "0",
        "new_service[min_sum]": "0", "new_service[delivery_var]": "0",
        "new_slot[date_expiration]": f"{slot_date}T{SLOT_EXPIRES}:00+03:00",
        "new_slot[chosen]": "0",
        "new_slot[slot_since]": f"{SLOT_SINCE}:00.0000000",
        "new_slot[slot_until]": f"{SLOT_UNTIL}:00.0000000",
        "new_slot[slot_during]": "", "new_slot[during_title]": "",
        "new_slot[slot_title]": f"Завтра с {SLOT_SINCE} — {SLOT_UNTIL}",
        "new_slot[date_service]": slot_date,
        "new_slot[priority_slots]": "20",
        "new_slot[ShopNo]": STORE["shop_no"],
        "new_slot[gettype]": "0",
        "sessid": sessid,
    })
    return slot_date


def to_number(text):
    if text is None:
        return None
    t = text.replace(" ", "").replace("\xa0", "").replace("\u2009", "").replace(",", ".")
    m = re.search(r"\d+(?:\.\d+)?", t)
    if not m:
        return None
    v = float(m.group())
    return int(v) if v.is_integer() else v


def parse_product(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    out = {"name": None, "pack": None, "unit": None, "price": None,
           "old_price": None, "qty": None, "address": header_address(page_html)}

    for tag in soup.find_all("script", {"type": "application/ld+json"}):
        try:
            data = json.loads(tag.string or "")
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("@type") == "Product":
            out["name"] = html.unescape(data.get("name", "")).replace("\xa0", " ").strip()
            break

    block = soup.select_one(".js-product-headprice") or soup
    price = block.select_one(".js-datalayer-catalog-list-price")
    old = block.select_one(".js-datalayer-catalog-list-price-old")
    out["price"] = to_number(price.get_text()) if price else None
    out["old_price"] = to_number(old.get_text()) if old else None

    head = soup.select_one(".Product__headWeightQ")
    head_text = head.get_text(" ", strip=True) if head else ""
    shown = block.select_one(".js-product-price-item .Price")
    shown_text = shown.get_text(" ", strip=True) if shown else ""
    if re.search(r"руб\s*/\s*кг", head_text.replace("\xa0", " ")):
        out["unit"] = "кг"
    else:
        m = re.search(r"/\s*([а-яa-z]+)", shown_text)
        out["unit"] = m.group(1) if m else None

    weight = soup.select_one(".js-product-weight-for-log")
    if weight:
        out["pack"] = weight.get_text(" ", strip=True)
    elif out["unit"] == "кг":
        out["pack"] = "by weight"

    rest = soup.select_one("#product-quantity-block")
    if rest is not None and rest.get("data-quantity") not in (None, ""):
        out["qty"] = to_number(rest["data-quantity"])
    return out


UNITS_EN = {
    "мл": "ml", "л": "l",
    "кг": "kg", "гр": "g", "г": "g",
    "шт": "pcs", "уп": "pack", "упак": "pack",
}


def unit_to_en(text):
    if not text:
        return text
    return re.sub(r"[а-яё]+\.?",
                  lambda m: UNITS_EN.get(m.group().rstrip(".").lower(), m.group()),
                  text, flags=re.IGNORECASE)


def build_row(date_str, item, parsed, url):
    price, old = parsed["price"], parsed["old_price"]
    qty = parsed["qty"]
    in_stock = "" if qty is None else ("yes" if qty > 0 else "no")
    if price is None:
        regular, promo, status = "", "", "error: price not found"
    else:
        on_promo = old is not None and old > price
        regular = old if on_promo else price
        promo = price if on_promo else ""
        status = "ok"
    return [date_str, item["item"], regular, promo, item["no"], item["id"],
            parsed["name"] or "", unit_to_en(parsed["pack"]) or "", unit_to_en(parsed["unit"]) or "",
            in_stock, "" if qty is None else qty, url, status]


KEY_FILE = "service_account.json"


def open_sheet():
    import gspread
    from google.oauth2.service_account import Credentials
    if os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
        info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    else:
        with open(KEY_FILE, encoding="utf-8") as f:
            info = json.load(f)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)


def item_name(en, ru):
    return (en or "").strip() or (ru or "").strip()


def basket_from_sheet(book):
    rows = book.worksheet(BASKET_SHEET).get_all_values()
    head = rows[0]
    no_col = head.index("#") if "#" in head else 0
    col = {name: head.index(name) for name in ("Item (EN)", "Product ID", "Product URL")}
    ru_col = head.index("Item (RU)") if "Item (RU)" in head else None
    items = []
    for r in rows[1:]:
        if len(r) > col["Product URL"] and r[col["Product URL"]].startswith("http"):
            ru = r[ru_col] if ru_col is not None and len(r) > ru_col else ""
            items.append({"no": r[no_col], "item": item_name(r[col["Item (EN)"]], ru),
                          "id": r[col["Product ID"]], "url": r[col["Product URL"]]})
    return items


def basket_from_csv(path):
    with open(path, encoding="utf-8") as f:
        return [{"no": r["#"], "item": item_name(r.get("Item (EN)"), r["Item (RU)"]),
                 "id": r["Product ID"], "url": r["Product URL"]} for r in csv.DictReader(f)]


class Tee:
    def __init__(self, stream, log_file):
        self.stream, self.log_file = stream, log_file

    def write(self, text):
        self.stream.write(text)
        self.log_file.write(text)
        self.flush()

    def flush(self):
        self.stream.flush()
        self.log_file.flush()


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    ap = argparse.ArgumentParser()
    ap.add_argument("--log", metavar="FILE",
                    help="also append all output (including errors) to this file")
    ap.add_argument("--dry-run", action="store_true",
                    help="read basket.csv and write results to a local CSV instead of Google Sheets")
    ap.add_argument("--force", action="store_true",
                    help="collect again even if today's rows are already in the sheet")
    args = ap.parse_args()

    if args.log:
        log_file = open(args.log, "a", encoding="utf-8")
        sys.stdout = Tee(sys.stdout, log_file)
        sys.stderr = Tee(sys.stderr, log_file)

    now = datetime.now(MSK)
    date_str = now.strftime("%Y-%m-%d")
    print(f"=== run started {now:%Y-%m-%d %H:%M} MSK ===", flush=True)

    book = None if args.dry_run else open_sheet()

    if book is not None and not args.force:
        dates = book.worksheet(PRICES_SHEET).col_values(1)
        if date_str in dates or now.strftime("%d.%m.%Y") in dates:
            print(f"prices for {date_str} are already in '{PRICES_SHEET}', nothing to do "
                  f"(use --force to collect again)", flush=True)
            return
    items = basket_from_csv("basket.csv") if args.dry_run else basket_from_sheet(book)
    print(f"{date_str}: {len(items)} products in the basket")

    s = new_session()
    slot_date = select_store_and_slot(s, now)
    print(f"store {STORE['shop_no']} and slot {slot_date} {SLOT_SINCE}-{SLOT_UNTIL} selected")

    rows, errors = [], 0
    for i, item in enumerate(items):
        if i:
            time.sleep(random.uniform(*DELAY))
        try:
            page = get(s, item["url"]).text
            parsed = parse_product(page)
            if i == 0:
                save_debug("first_product.html", page)
                print(f"first product page: {describe(page)}", flush=True)
            if i == 0 and STORE_CHECK not in parsed["address"]:
                raise RuntimeError(f"store not selected, page header shows: {parsed['address']!r}")
            row = build_row(date_str, item, parsed, item["url"])
        except RuntimeError as e:
            if "store not selected" in str(e):
                raise
            row = [date_str, item["item"], "", "", item["no"], item["id"], "", "", "", "", "",
                   item["url"], f"error: {e}"]
        if row[-1] != "ok":
            errors += 1
        rows.append(row)
        print(" | ".join(str(x) for x in row[1:11]), "|", row[-1], flush=True)

    if args.dry_run:
        out = f"prices_{date_str}.csv"
        with open(out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(HEADER)
            w.writerows(rows)
        print(f"saved {out}")
    else:
        import gspread
        ws = book.worksheet(PRICES_SHEET)
        start = len(ws.col_values(1)) + 1
        end = start + len(rows) - 1
        if end > ws.row_count:
            ws.add_rows(end - ws.row_count)
        last_col = gspread.utils.rowcol_to_a1(1, len(HEADER)).rstrip("1")
        ws.update(values=rows, range_name=f"A{start}:{last_col}{end}",
                  value_input_option="USER_ENTERED")
        print(f"wrote {len(rows)} rows to '{PRICES_SHEET}' (rows {start}-{end})")

    if errors:
        print(f"{errors} products failed", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
