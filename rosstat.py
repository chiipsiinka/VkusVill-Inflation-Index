import argparse
import os
import re
import sys
import time
from datetime import date, timedelta
from urllib.parse import urljoin, urlparse

import openpyxl
import requests

SHEET = "Rosstat Weekly"
LOCAL_FILE = "nedel_Ipc.xlsx"
RELEASE_URLS = ["https://rosstat.gov.ru/storage/mediabank/Nedel_ipc.xlsx",
                "https://rosstat.gov.ru/storage/mediabank/nedel_Ipc.xlsx"]
SEARCH_PAGES = ["https://rosstat.gov.ru/statistics/price",
                "https://rosstat.gov.ru/free_doc/new_site/prices/ipc_data.htm"]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36")
RELEASE_DELAY = 2
FIRST_ROW = 3
WEEK_COL = 5
N_WEEKS = 7

ROSSTAT_NAMES = {
    1: "Молоко питьевое цельное пастеризованное 2,5-3,2% жирности",
    2: "Молоко питьевое цельное ультрапастеризованное 2,5-3,2% жирности",
    3: "Сметана",
    4: "Творог",
    5: "Кефир",
    6: "Сыры твердые, полутвердые и мягкие",
    7: "Масло сливочное",
    8: "Крупа гречневая-ядрица",
    9: "Рис шлифованный",
    10: "Пшено",
    11: "Макаронные изделия из пшеничной муки высшего сорта",
    12: "Вермишель",
    13: "Мука пшеничная",
    14: "Хлеб из ржаной муки и из смеси муки ржаной и пшеничной",
    15: "Хлеб и булочные изделия из пшеничной муки различных сортов",
    16: "Сахар-песок",
    17: "Колбаса вареная",
    18: "Колбаса полукопченая и варено-копченая",
    19: "Куры охлажденные и мороженые",
    20: "Рыба мороженая неразделанная",
    21: "Сосиски, сардельки",
    22: "Мыло туалетное",
    23: "Паста зубная",
    24: "Щетка зубная",
    25: "Шампунь",
    26: "Порошок стиральный",
    27: "Бумага туалетная",
}

MONTHS = {"января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6, "июля": 7,
          "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12}


def norm(text):
    return re.sub(r"[^а-я0-9]", "", str(text).lower().replace("ё", "е"))


def header_date(label, year):
    m = re.search(r"(\d{1,2})\s+([а-я]+)", str(label or "").lower())
    if m and m.group(2) in MONTHS:
        return date(year, MONTHS[m.group(2)], int(m.group(1)))
    return None


def read_release(path, year):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if str(year) not in wb.sheetnames:
            sys.exit(f"sheet '{year}' not found in {path}; sheets: {wb.sheetnames}")
        rows = list(wb[str(year)].iter_rows(values_only=True))
    finally:
        wb.close()

    head_i = next(i for i, r in enumerate(rows) if r and str(r[0]).strip() == "Наименование")
    dates = {j: header_date(c, year) for j, c in enumerate(rows[head_i]) if j}
    dates = {j: d for j, d in dates.items() if d}

    by_name = {norm(r[0]): r for r in rows[head_i + 1:] if r and r[0]}
    out, missing = {}, []
    for no, name in ROSSTAT_NAMES.items():
        key = norm(name)
        row = next((r for k, r in by_name.items() if k.startswith(key)), None)
        if row is None:
            missing.append(f"{no}: {name}")
            continue
        out[no] = {d: row[j] for j, d in dates.items()
                   if j < len(row) and isinstance(row[j], (int, float))}
    if missing:
        sys.exit("not found in the release (did Rosstat rename them?):\n  " + "\n  ".join(missing))
    odd = [f"{no}: {v}" for no, weeks in out.items() for v in weeks.values() if not 50 <= v <= 200]
    if odd:
        sys.exit("implausible values in the release (expected about 100):\n  " + "\n  ".join(odd[:10]))
    return out, max(dates.values())


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


def download(dest, year):
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9"})
    try:
        import ssl
        import truststore
        from requests.adapters import HTTPAdapter

        class SystemTrustAdapter(HTTPAdapter):
            def init_poolmanager(self, *a, **kw):
                kw["ssl_context"] = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
                return super().init_poolmanager(*a, **kw)

        s.mount("https://", SystemTrustAdapter())
    except ImportError:
        pass

    insecure = False

    def plain_get(url):
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        with requests.Session() as p:
            p.headers.update(s.headers)
            return p.get(url, timeout=60, verify=False)

    def get(url):
        nonlocal insecure
        rosstat_host = urlparse(url).hostname in ("rosstat.gov.ru", "www.rosstat.gov.ru")
        try:
            if insecure and rosstat_host:
                return plain_get(url)
            return s.get(url, timeout=60)
        except requests.exceptions.SSLError:
            if not rosstat_host or insecure:
                raise
            print("  certificate of rosstat.gov.ru is not trusted on this computer; "
                  "downloading without certificate check (the file is verified below)")
            insecure = True
            return plain_get(url)

    def fetch(url):
        try:
            r = get(url)
        except requests.RequestException as e:
            print(f"  {url}: {e}")
            return None
        if r.status_code != 200 or not r.content.startswith(b"PK"):
            print(f"  {url}: HTTP {r.status_code}, not an .xlsx file")
            return None
        return r.content

    candidates = []
    for page in SEARCH_PAGES:
        try:
            html = get(page).text
        except requests.RequestException as e:
            print(f"  {page}: {e}")
            continue
        for href in re.findall(r'href="([^"]*nedel_ipc[^"]*\.xlsx)"', html, flags=re.I):
            candidates.append(urljoin(page, href))
    candidates += RELEASE_URLS
    candidates = list(dict.fromkeys(candidates))

    tmp = "download_check.xlsx"
    best, best_last, best_url = None, None, None
    for url in candidates:
        content = fetch(url)
        if not content:
            continue
        with open(tmp, "wb") as f:
            f.write(content)
        try:
            _, last = read_release(tmp, year)
        except (SystemExit, Exception) as e:
            print(f"  {url}: file rejected: {e}")
            continue
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
        if best_last is None or last > best_last:
            best, best_last, best_url = content, last, url
    if best is None:
        return None

    print(f"  using {best_url} (latest week ends {best_last:%d.%m.%Y})")
    try:
        with open(dest, "wb") as f:
            f.write(best)
    except PermissionError:
        dest = os.path.splitext(dest)[0] + "_new.xlsx"
        with open(dest, "wb") as f:
            f.write(best)
        print(f"  the old file is open in another program, so the new one is saved as {dest}")
    return dest


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    ap = argparse.ArgumentParser()
    ap.add_argument("file", nargs="?", help="use this Rosstat release (.xlsx) instead of downloading")
    ap.add_argument("--force", action="store_true", help="download and fill even if no week is due")
    ap.add_argument("--dry-run", action="store_true", help="print the values, do not write to the sheet")
    ap.add_argument("--log", metavar="FILE", help="also append all output (including errors) to this file")
    args = ap.parse_args()

    from scraper import open_sheet
    if args.log:
        log_file = open(args.log, "a", encoding="utf-8")
        sys.stdout = Tee(sys.stdout, log_file)
        sys.stderr = Tee(sys.stderr, log_file)
    ws = open_sheet().worksheet(SHEET)

    labels = ws.row_values(2)[WEEK_COL - 1:WEEK_COL - 1 + N_WEEKS]
    weeks = [datetime_from_ru(s) for s in labels]
    year = next(d.year for d in weeks if d)

    n = len(ROSSTAT_NAMES)
    last_col = chr(ord("A") + WEEK_COL - 2 + N_WEEKS)
    grid = ws.get(f"{chr(ord('A') + WEEK_COL - 1)}{FIRST_ROW}:{last_col}{FIRST_ROW + n - 1}")
    filled = [any(len(r) > k and str(r[k]).strip() for r in grid) for k in range(N_WEEKS)]

    today = date.today()
    due = [w for w, f in zip(weeks, filled) if w and not f and today >= w + timedelta(days=RELEASE_DELAY)]
    if not due and not args.force and not args.file:
        upcoming = [w for w, f in zip(weeks, filled) if w and not f]
        if upcoming:
            nxt = upcoming[0] + timedelta(days=RELEASE_DELAY)
            print(f"Rosstat: sheet is up to date; next release (week ending {upcoming[0]:%d.%m}) "
                  f"expected on {nxt:%d.%m} at 19:00")
        else:
            print("Rosstat: all 7 weeks are already in the sheet")
        return

    path = args.file
    if not path:
        print("Rosstat: downloading the weekly release...")
        path = download(LOCAL_FILE, year)
        if path:
            print(f"  saved {LOCAL_FILE}")
        elif os.path.exists(LOCAL_FILE):
            age = (time.time() - os.path.getmtime(LOCAL_FILE)) / 86400
            print(f"  download failed; using the {LOCAL_FILE} already in the folder "
                  f"(saved {age:.0f} days ago). If it is old, download it by hand.")
            path = LOCAL_FILE
        else:
            sys.exit(f"  download failed and there is no {LOCAL_FILE} in the folder; "
                     f"download it from rosstat.gov.ru by hand")

    data, last = read_release(path, year)
    print(f"{path}: latest Rosstat week ends {last:%d.%m.%Y}")

    numbers = ws.col_values(1)[FIRST_ROW - 1:FIRST_ROW - 1 + n]
    written = []
    for k, week_end in enumerate(weeks):
        if week_end is None or week_end > last:
            continue
        col = [[data[int(no)].get(week_end, "")] for no in numbers]
        letter = chr(ord("A") + WEEK_COL - 1 + k)
        rng = f"{letter}{FIRST_ROW}:{letter}{FIRST_ROW + len(col) - 1}"
        written.append(f"Week {k + 1} ({week_end:%d.%m})")
        if args.dry_run:
            print(rng, [c[0] for c in col])
        else:
            ws.update(values=col, range_name=rng, value_input_option="RAW")

    if written:
        print(("would write: " if args.dry_run else "written: ") + ", ".join(written))
    missing = [w for w in due if w > last]
    if missing:
        print(f"Rosstat has not published the week ending {missing[0]:%d.%m} yet "
              f"(latest in the file: {last:%d.%m}); will try again on the next run")
    elif not written:
        print(f"none of our weeks is in this release yet (Week 1 ends {weeks[0]:%d.%m.%Y})")


def datetime_from_ru(text):
    m = re.fullmatch(r"\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s*", str(text or ""))
    return date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


if __name__ == "__main__":
    main()
