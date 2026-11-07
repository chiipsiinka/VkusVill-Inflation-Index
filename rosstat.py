import argparse
import io
import json
import os
import re
import sys
import time
from datetime import date, timedelta
from urllib.parse import quote, urljoin, urlparse

import openpyxl
import requests

SHEET = "Rosstat Weekly"
COMPARISON_SHEET = "Rosstat Comparison"
LOCAL_FILE = "nedel_Ipc.xlsx"
STATE_FILE = "rosstat_state.json"
RELEASE_URLS = ["https://rosstat.gov.ru/storage/mediabank/Nedel_ipc.xlsx",
                "https://rosstat.gov.ru/storage/mediabank/nedel_Ipc.xlsx"]
SEARCH_PAGES = ["https://rosstat.gov.ru/statistics/price/",
                "https://rosstat.gov.ru/statistics/price",
                "https://rosstat.gov.ru/free_doc/new_site/prices/ipc_data.htm"]
MEDIABANK = "https://rosstat.gov.ru/storage/mediabank/"
KNOWN_PDF_RELEASE = (168, date(2026, 10, 28))
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
PDF_NAMES = {**ROSSTAT_NAMES,
             14: "Хлеб из ржаной муки и из смеси ржаной и пшеничной"}

MONTHS = {"января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6, "июля": 7,
          "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12}


def norm(text):
    return re.sub(r"[^а-я0-9]", "", str(text).lower().replace("ё", "е"))


def header_date(label, year):
    m = re.search(r"(\d{1,2})\s+([а-я]+)", str(label or "").lower())
    if m and m.group(2) in MONTHS:
        return date(year, MONTHS[m.group(2)], int(m.group(1)))
    return None


def check_plausible(data):
    odd = [f"{no}: {v}" for no, weeks in data.items() for v in weeks.values() if not 50 <= v <= 200]
    if odd:
        sys.exit("implausible values in the release (expected about 100):\n  " + "\n  ".join(odd[:10]))


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
    check_plausible(out)
    return out, max(dates.values())


def pdf_text(source):
    from pypdf import PdfReader
    reader = PdfReader(io.BytesIO(source) if isinstance(source, bytes) else source)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def loose(name):
    return r"\s*".join(re.escape(ch) for ch in name if not ch.isspace())


def first_number(text):
    m = re.match(r"(\d{2,3},\d{2})", re.sub(r"\s+", "", text[:40]))
    return float(m.group(1).replace(",", ".")) if m else None


def read_pdf_release(source):
    text = pdf_text(source)
    flat = re.sub(r"\s+", " ", text)

    m = re.search(r"ОБ ОЦЕНКЕ ИНДЕКСА ПОТРЕБИТЕЛЬСКИХ ЦЕН (?:С|СО) .*?ПО (\d{1,2}) ([А-Я]+) (\d{4})",
                  flat)
    if not m or m.group(2).lower() not in MONTHS:
        sys.exit("this PDF is not Rosstat's weekly release 'Об оценке индекса потребительских цен'")
    week_end = date(int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1)))

    m = re.search(r"\(\s*оценка\s*\)", text)
    overall = first_number(text[m.end():]) if m else None

    start = text.find("в процентах")
    table = text[start:] if start >= 0 else text
    items, missing = {}, []
    for no, name in PDF_NAMES.items():
        m = re.search(loose(name), table)
        value = first_number(table[m.end():]) if m else None
        if value is None:
            missing.append(f"{no}: {name}")
        else:
            items[no] = value
    if missing:
        sys.exit("not found in the PDF release (did Rosstat rename them?):\n  " + "\n  ".join(missing))
    check_plausible({no: {week_end: v} for no, v in items.items()})
    if overall is not None and not 50 <= overall <= 200:
        overall = None
    return week_end, items, overall


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


def make_getter():
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
    plain = None

    def plain_get(url):
        nonlocal plain
        if plain is None:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            plain = requests.Session()
            plain.headers.update(s.headers)
        return plain.get(url, timeout=60, verify=False)

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

    return get


def fetch(get, url, magic, quiet=False):
    try:
        r = get(url)
    except requests.RequestException as e:
        if not quiet:
            print(f"  {url}: {e}")
        return None
    if r.status_code != 200 or not r.content.startswith(magic):
        if not quiet:
            print(f"  {url}: HTTP {r.status_code}, not the expected file")
        return None
    return r.content


def page_links(get, page, pattern):
    try:
        r = get(page)
    except requests.RequestException as e:
        print(f"  {page}: {e}")
        return []
    if r.status_code != 200:
        return []
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', r.text)
    return [urljoin(page, h) for h in hrefs if re.search(pattern, h, flags=re.I)]


def download(get, dest, year):
    candidates = []
    for page in SEARCH_PAGES:
        candidates += page_links(get, page, r"nedel_ipc[^\"']*\.xlsx$")
    candidates += RELEASE_URLS
    candidates = list(dict.fromkeys(candidates))

    tmp = "download_check.xlsx"
    best, best_last, best_url = None, None, None
    for url in candidates:
        content = fetch(get, url, b"PK")
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


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            s = json.load(f)
        return int(s["number"]), date.fromisoformat(s["date"])
    except (OSError, ValueError, KeyError):
        return KNOWN_PDF_RELEASE


def save_state(number, published):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump({"number": number, "date": published.isoformat()}, f)


def find_pdf_release(get, week_end):
    def try_url(url, quiet=True):
        content = fetch(get, url, b"%PDF", quiet=quiet)
        if not content:
            return None
        try:
            got = read_pdf_release(content)
        except (SystemExit, Exception) as e:
            print(f"  {url}: file rejected: {e}")
            return None
        if got[0] != week_end:
            return None
        print(f"  using {url}")
        return got

    title = (f"Об оценке индекса потребительских цен по {week_end.day} "
             f"{[k for k, v in MONTHS.items() if v == week_end.month][0]} {week_end.year}")
    for url in page_links(get, "https://rosstat.gov.ru/search?q=" + quote(title),
                          r"/storage/mediabank/[^\"']+\.pdf$"):
        got = try_url(url)
        if got:
            return got

    got = try_url(f"{MEDIABANK}ipc{week_end:%m}_{week_end:%d}.pdf")
    if got:
        return got

    last_no, last_date = load_state()
    weeks = max(1, (week_end + timedelta(days=RELEASE_DELAY) - last_date).days // 7 + 1)
    numbers = range(last_no + 1, last_no + 8 * weeks + 5)
    today = date.today()
    days = [week_end + timedelta(days=d) for d in (2, 3, 4, 1, 5, 6, 7, 8)]
    for published in [d for d in days if d <= today]:
        for no in numbers:
            got = try_url(f"{MEDIABANK}{no}_{published:%d-%m-%Y}.pdf")
            if got:
                save_state(no, published)
                return got
    return None


def set_overall(book, week_end, overall, dry_run):
    if overall is None:
        return
    change = round(overall - 100, 2)
    ws = book.worksheet(COMPARISON_SHEET)
    col_a = ws.col_values(1)
    row = next((i + 1 for i, v in enumerate(col_a) if datetime_from_ru(v) == week_end), None)
    if row is None:
        return
    if (ws.acell(f"B{row}").value or "").strip():
        return
    if dry_run:
        print(f"would write overall CPI change {change}% to {COMPARISON_SHEET}!B{row}")
        return
    ws.update(values=[[change]], range_name=f"B{row}", value_input_option="RAW")
    ws.format(f"B{row}", {"numberFormat": {"type": "NUMBER", "pattern": '0.00"%"'}})
    print(f"written: overall CPI change {change}% for the week ending {week_end:%d.%m}")


def write_week(ws, k, week_end, values_by_no, numbers, dry_run):
    col = [[values_by_no.get(int(no), "")] for no in numbers]
    letter = chr(ord("A") + WEEK_COL - 1 + k)
    rng = f"{letter}{FIRST_ROW}:{letter}{FIRST_ROW + len(col) - 1}"
    if dry_run:
        print(rng, [c[0] for c in col])
    else:
        ws.update(values=col, range_name=rng, value_input_option="RAW")
    return f"Week {k + 1} ({week_end:%d.%m})"


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    ap = argparse.ArgumentParser()
    ap.add_argument("file", nargs="?",
                    help="use this Rosstat .xlsx table or .pdf release instead of downloading")
    ap.add_argument("--force", action="store_true", help="download and fill even if no week is due")
    ap.add_argument("--dry-run", action="store_true", help="print the values, do not write to the sheet")
    ap.add_argument("--log", metavar="FILE", help="also append all output (including errors) to this file")
    args = ap.parse_args()

    from scraper import open_sheet
    if args.log:
        log_file = open(args.log, "a", encoding="utf-8")
        sys.stdout = Tee(sys.stdout, log_file)
        sys.stderr = Tee(sys.stderr, log_file)
    book = open_sheet()
    ws = book.worksheet(SHEET)

    labels = ws.row_values(2)[WEEK_COL - 1:WEEK_COL - 1 + N_WEEKS]
    weeks = [datetime_from_ru(s) for s in labels]
    year = next(d.year for d in weeks if d)
    n = len(ROSSTAT_NAMES)
    numbers = ws.col_values(1)[FIRST_ROW - 1:FIRST_ROW - 1 + n]

    if args.file and args.file.lower().endswith(".pdf"):
        week_end, items, overall = read_pdf_release(args.file)
        print(f"{args.file}: Rosstat release for the week ending {week_end:%d.%m.%Y}")
        if week_end not in weeks:
            sys.exit(f"the week ending {week_end:%d.%m.%Y} is not one of the sheet's weeks")
        done = write_week(ws, weeks.index(week_end), week_end, items, numbers, args.dry_run)
        print(("would write: " if args.dry_run else "written: ") + done)
        set_overall(book, week_end, overall, args.dry_run)
        return

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

    get = make_getter()
    path = args.file
    if not path:
        print("Rosstat: downloading the weekly table (.xlsx)...")
        path = download(get, LOCAL_FILE, year)
        if path:
            print(f"  saved {path}")
        elif os.path.exists(LOCAL_FILE):
            age = (time.time() - os.path.getmtime(LOCAL_FILE)) / 86400
            print(f"  download failed; using the {LOCAL_FILE} already in the folder "
                  f"(saved {age:.0f} days ago)")
            path = LOCAL_FILE

    written, have = [], set()
    if path:
        data, last = read_release(path, year)
        print(f"{path}: latest Rosstat week ends {last:%d.%m.%Y}")
        for k, week_end in enumerate(weeks):
            if week_end is None or week_end > last:
                continue
            written.append(write_week(ws, k, week_end,
                                      {no: v.get(week_end, "") for no, v in data.items()},
                                      numbers, args.dry_run))
            have.add(week_end)

    still_due = [w for w in due if w not in have]
    not_found = []
    if still_due and not args.file:
        print("Rosstat: the table has no data for "
              + ", ".join(f"{w:%d.%m}" for w in still_due) + "; looking for the PDF releases...")
        for week_end in still_due:
            got = find_pdf_release(get, week_end)
            if not got:
                not_found.append(week_end)
                continue
            _, items, overall = got
            written.append(write_week(ws, weeks.index(week_end), week_end, items, numbers,
                                      args.dry_run))
            set_overall(book, week_end, overall, args.dry_run)
    else:
        not_found = still_due

    if written:
        print(("would write: " if args.dry_run else "written: ") + ", ".join(written))
    if not_found:
        print("Rosstat release not found for the week ending "
              + ", ".join(f"{w:%d.%m}" for w in not_found)
              + "; will try again on the next run. If it is already out, download the PDF "
                "'Об оценке индекса потребительских цен' from rosstat.gov.ru and drag it "
                "onto run_rosstat.bat")
    elif not written:
        print(f"none of our weeks is in this release yet (Week 1 ends {weeks[0]:%d.%m.%Y})")


def datetime_from_ru(text):
    m = re.fullmatch(r"\s*(\d{1,2})\.(\d{1,2})\.(\d{4})\s*", str(text or ""))
    return date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


if __name__ == "__main__":
    main()
