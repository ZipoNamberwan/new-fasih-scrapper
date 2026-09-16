# Fasih Scraper

Logs into [fasih-sm.bps.go.id](https://fasih-sm.bps.go.id/app), lets you pick a survey/period, downloads respondent data region-by-region, and exports it to an Excel file.

## Requirements

- Python 3.9+
- Windows, macOS, or Linux (a real browser window opens by default — see [Configuration](#configuration))

## Setup

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Download the Camoufox browser binary (one-time, required):
   ```
   python -m camoufox fetch
   ```

## Running

```
python scrape.py
```

A browser window opens first, then the terminal asks for:

1. **Username / Email** and **Password**
2. **OTP** (leave blank if your account doesn't use one)
3. **Action mode**:
   - `1` — Start a new download (deletes any existing `data.xlsx`/JSON results for the selected period first)
   - `2` — Resume download (skips regions already completed and reuses cached JSON detail files)
   - `3` — Extract only (rebuilds `data.xlsx` from JSON files already downloaded, no network scraping)
4. **Survey** to scrape, from the list fetched after login
5. **Survey period(s)** — a single number, comma-separated (`1,2`), or `all`
6. **Region level** (skipped if scraping metadata only has one level) — how deep into the region hierarchy to drill down (e.g. province → regency → district)

The script then paginates through respondents for every target region, fetches each respondent's detail, and writes progress as it goes — so it's safe to stop and resume later with mode `2`.

At the end it prompts `Press Enter to close browser...` so you can inspect the page before it closes.

## Output

Results are written under:
```
result/{survey_id}/{period_id}/
  data.xlsx              # final exported spreadsheet
  region_progress.json   # tracks which regions have been completed (for resume)
  json/{assignment_id}.json   # cached raw detail response per respondent
  template/                    # survey template JSON definitions (used to build roster sheets)
```

`data.xlsx` includes a `Data` sheet plus one sheet per repeating-group ("roster") field defined in the survey template, and an `Items_Detail` fallback sheet for anything the template doesn't explain.

## Configuration

Settings live at the top of [scrape.py](scrape.py):

| Setting | Purpose |
|---|---|
| `HEADLESS` | Set `True` to run the browser without a visible window (keep `False` while testing login/OTP) |
| `ROOT_REGION_NAME` / `ROOT_REGION_CODE` | Starting province for the region hierarchy (defaults to Jawa Timur) |
| `ENABLE_USER_PROMPT_REGION_LEVEL` | If `False`, always uses the deepest region level instead of asking |
| `DEFAULT_DELAY` / `DETAIL_DELAY` | Seconds to wait between paginated/detail requests (avoid rate limiting) |
| `RESULT_DIR` / `OUTPUT_FILENAME` | Where output is written and the Excel file name |

The browser profile (cookies/login session) persists in `camoufox_profile/`. Delete that folder to force a clean login.

## Extracting Excel without re-scraping

If you already have JSON files downloaded and just want to regenerate the Excel file, run `scrape.py` and choose action mode `3`, or call it directly:
```
python -c "from extract_excel import extract_json_to_excel; extract_json_to_excel(SURVEY_ID, PERIOD_ID)"
```
