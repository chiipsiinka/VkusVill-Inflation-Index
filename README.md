# VkusVill daily price scraper

Student project "Retail: build your own inflation index" (HSE, 6_AcE).

Once a day the scraper collects prices for a fixed basket of 27 products from vkusvill.ru and adds them
to the "Daily Prices" sheet of our Google Sheet. We use these prices to build a Jevons index and compare it
with Rosstat's weekly price data.

## Where it runs

VkusVill does not open for visitors from outside Russia, so the scraper cannot run on GitHub's servers.
It runs on our computer in St. Petersburg. Windows Task Scheduler starts `run_daily.bat` once a day.

## How it works

1. The script opens vkusvill.ru.
2. It picks our St. Petersburg store (shop 7064, pr-t Medikov 10 k1) and the first time slot of the next
   day (11:00 to 13:00). The site forgets this choice, so the script repeats it on every run. If the store
   address does not appear on the page, the script stops.
3. It reads the basket (product links and IDs) from the "Basket" sheet.
4. It opens each product page and reads the regular price, the promo price, the price unit (per pack or
   per kg), the pack size and how many are left in the store.
5. It adds one row per product to "Daily Prices".
6. `rosstat.py` then checks whether Rosstat has published a new week. If so, it downloads the table
   nedel_Ipc.xlsx from rosstat.gov.ru, finds our 27 products in it and writes the weekly price
   changes to the "Rosstat Weekly" sheet.
7. If the table does not have the new week yet (Rosstat sometimes moves or renames it), the script
   looks for the weekly PDF release "Об оценке индекса потребительских цен" instead. Besides our
   27 products, the release gives the overall CPI for the week, which goes to column B of the
   "Rosstat Comparison" sheet if that cell is still empty. Every file is checked (right week, all 27
   products, values close to 100) before anything is written.

## Files

| File | What it is |
|---|---|
| `scraper.py` | collects prices from vkusvill.ru |
| `rosstat.py` | loads Rosstat's weekly data for our products |
| `run_daily.bat` | daily launcher (prices, then Rosstat); writes a log to `log.txt` |
| `run_rosstat.bat` | loads Rosstat data by hand; you can drag a downloaded Rosstat file (.xlsx table or .pdf release) onto it |
| `basket.csv` | copy of the basket, used for test runs |
| `requirements.txt` | Python packages |

The script also needs `service_account.json`, the key that lets it write to the Google Sheet.
It stays on our computer and is not in this repository.

## Test run without Google Sheets

```
pip install -r requirements.txt
python scraper.py --dry-run
```

This reads `basket.csv` and saves the results to `prices_YYYY-MM-DD.csv`.
