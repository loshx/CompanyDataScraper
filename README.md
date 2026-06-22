# Company Data Scraper

Desktop scraper built with Python, Selenium, and CustomTkinter for collecting company information from `data2b.md` and exporting the results to Excel.

## What it does

- Search companies directly on `data2b.md`
- Use Google-assisted search when the direct lookup is not enough
- Extract company identity, contact, address, activity, and financial fields
- Export clean spreadsheets with formatted headers and filters
- Extract company IDNO values from paginated result URLs

## Tech stack

- Python 3
- Selenium
- CustomTkinter / Tkinter
- pandas
- openpyxl
- XlsxWriter

## Project files

- `main.py` - desktop UI and scraping workflows
- `requirements.txt` - Python dependencies

## Quick start

```bash
git clone https://github.com/loshx/CompanyDataScraper.git
cd CompanyDataScraper
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS / Linux:

```bash
source .venv/bin/activate
```

Install dependencies and run the app:

```bash
pip install -r requirements.txt
python main.py
```

## Usage

### Company search

1. Enter one or more company names, one per line.
2. Run the normal search or the Google-assisted search.
3. Exported results are saved to `company_data.xlsx`.

### IDNO extraction by URL

1. Open the second screen in the app.
2. Paste a paginated `data2b.md` URL.
3. Choose how many pages to scan.
4. Extracted IDs are saved to `company_ids.xlsx`.

## Notes

- The scraper depends on the current structure of `data2b.md`, so selectors may need updates if the website changes.
- Chrome is installed automatically through `webdriver-manager`.

