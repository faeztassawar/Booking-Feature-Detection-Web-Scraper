# Website Booking Feature Scraper

Scrapes company websites to detect online appointment/booking features and writes results into a local Excel (`.xlsx`) file.

## What It Does

1. Reads website URLs from the `Website` column (matched by header name)
2. Visits each site (homepage + a few related internal pages)
3. Detects online booking via CTAs, URL patterns, third-party schedulers, embeds, and appointment forms
4. Writes `YES`, `NO`, `SKIP`, or `ERROR` into the `Booking Feature` column

## Prerequisites

- Python 3.9 or higher
- An `.xlsx` workbook with header row containing **`Website`** and **`Booking Feature`**

## Installation

```bash
cd /home/dev/Projects/Scrapper
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

`playwright` / SeleniumBase UC are used selectively for JS-heavy or bot-blocked pages (not on every site).

## Spreadsheet Format

Row 1 must include these headers (column order can vary):

| Column | Purpose |
|--------|---------|
| `Website` | Domain or URL to inspect |
| `Booking Feature` | Output: `YES` / `NO` / `SKIP` / `ERROR` |

Other columns (name, company, email, etc.) are left unchanged.

## Running the Scraper

Pass the workbook path as a required CLI argument:

```bash
python scraper.py "Booking Data - Health US SL L2 Yes (1).xlsx"
```

### Expected Output

```
[15:30:00] INFO - ============================================================
[15:30:00] INFO - Healthcare Website Booking Feature Scraper
[15:30:00] INFO - ============================================================
[15:30:00] INFO - Loading local workbook: Booking Data - Health US SL L2 Yes (1).xlsx
[15:30:01] INFO - Columns resolved: Website=col 5, Booking Feature=col 15
[15:30:01] INFO - Rows: 7000 websites | already done: 3200 | remaining (blank): 3800
[15:30:01] INFO - ------------------------------------------------------------
[15:30:01] INFO - [3201/7000] Checking: exampleclinic.com
[15:30:07] INFO -     Booking Feature: YES (Found booking-related link)
...
```

**Tip:** Do not open the same `.xlsx` in Excel while the scraper is running (file lock / overwrite risk).

## Cell Values

| Value | Meaning |
|-------|---------|
| `YES` | Online booking/appointment feature found |
| `NO` | Inspected successfully; no online booking found |
| `SKIP` | Empty or invalid URL |
| `ERROR` | Could not inspect (timeout, DNS, captcha/bot wall, SSL, etc.) |

## What Counts as "YES"

- Booking CTAs: “Book Appointment”, “Schedule Online”, “Virtual Visit”, “Schedule a Tour”, etc.
- Booking URL paths: `/appointment`, `/book`, `/schedule`, `/booking`, …
- Third-party platforms (Zocdoc, Calendly, Acuity, MyChart, Jane, Mindbody, Everseat, Cal.com, …)
- Embedded booking iframes / scheduling widgets
- Appointment request forms (date/time/provider-style fields)
- Booking found on a linked internal page (appointment, providers, …)

## What Counts as "NO"

- Generic contact forms (name / email / message only) — including “Make an appointment” links that only go to `/contact`
- Phone-only instructions (“Call us to schedule”)
- Mere mention of “appointment” with no online action
- Patient portal / login alone (not treated as online booking)
- Bill-pay portals (e.g. HealthiPass pay links)

## Resume Behavior

- Rows with `YES`, `NO`, `SKIP`, or `ERROR` are **skipped** (not overwritten)
- Only blank `Booking Feature` cells are processed
- Empty website cells are written as `SKIP` so resume does not get stuck
- To retry `ERROR` rows later, change `DONE_VALUES` in `scraper.py` so `ERROR` is omitted

Results are saved to disk every 10 writes and again on exit / `Ctrl+C`.

## Rate Limiting & Resilience

- ~2.5s delay between websites
- Up to 4 pages inspected per site (homepage + follow-ups)
- Multiple fetch rounds with rotating TLS fingerprints for flaky hosts
- URL variants tried (`https`/`http`, `www`/bare)
- Soft bot walls may fall back to Playwright, then SeleniumBase UC
- Hard 403 / captcha / forbidden shells are marked `ERROR` (not false `NO`)

## Files

| File | Description |
|------|-------------|
| `scraper.py` | Main scraper |
| `requirements.txt` | Python dependencies |
| `*.xlsx` | Input/output workbooks you provide |
| `README.md` | This file |

## Troubleshooting

### "Workbook not found"
- Check the path passed to `python scraper.py "..."`
- Run from the project directory, or use an absolute path

### "Could not find 'Website' / 'Booking Feature' column"
- Confirm row 1 has headers named exactly `Website` and `Booking Feature`

### Many `ERROR` results
- Site may block datacenter IPs (Cloudflare / SiteGround captcha)
- Re-run later, or temporarily allow `ERROR` retries via `DONE_VALUES`

### File looks locked / save failed
- Close the workbook in Excel/LibreOffice before starting
- Only one scraper process should write a given file

## Stopping the Scraper

Press `Ctrl+C` to stop gracefully. The current site finishes, the workbook is saved, then the process exits. Re-run the same command to continue from blank cells.

## Limitations

- Heavy SPAs may hide booking UI until JS runs; browser fallback helps but is not perfect
- Bot-protected sites may still return `ERROR`
- Detection is heuristic — spot-check edge cases when you need audit-level accuracy
