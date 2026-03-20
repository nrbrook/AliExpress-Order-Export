# AliExpress Order Export

This project exports AliExpress buyer orders to CSV and can optionally save invoice PDFs. It is built around the AliExpress MTop endpoints observed from real order-list, order-detail, and invoice traffic.

## What it does

- Stores AliExpress auth cookies in `.auth/cookies.json`
- Supports an interactive `--setup` flow that detects installed browsers, lets you pick a browser/profile, imports cookies, and tests the live API connection
- Imports those cookies from either:
  - an installed browser profile via `--setup`
  - a HAR capture
  - a Firefox profile or `cookies.sqlite`
- Calls the order list, order detail, invoice list, and invoice file APIs directly
- Exports a CSV into `exports/` with the selected date range in the filename
- Writes a readable order CSV with key fields plus JSON payload columns for the full exposed data
- Writes a separate order-line CSV with one row per purchased item, including item titles and SKU data
- Can also write an XLSX workbook with an `Orders` sheet and an `Order Lines` sheet, each with embedded thumbnails when the images can be fetched
- Saves invoice PDFs into `exports/pdf/` and skips duplicates
- Supports offline export from HAR captures for debugging or one-off recovery

## Install

The script declares its dependencies in a [PEP 723](https://peps.python.org/pep-0723/) header, so you can run it with `uv run aliexpress_export.py` from the repo root without creating a virtual environment first.

Optional: create a project `.venv` and install the same dependencies for editor tooling (for example Ruff):

```bash
uv venv --python 3.13
uv sync
```

## Usage

Run the guided browser setup once:

```bash
uv run aliexpress_export.py --setup
```

Or, import cookies once from Firefox, then export live:

```bash
uv run aliexpress_export.py \
  --start-date 2026-03-01 \
  --end-date 2026-03-31 \
  --firefox-profile "$HOME/Library/Application Support/Firefox/Profiles/<profile>"
```

Reuse the stored cookies on later runs:

```bash
uv run aliexpress_export.py \
  --start-date 2026-03-01 \
  --end-date 2026-03-31 \
  --download-pdfs
```

Generate the CSVs plus an XLSX workbook with thumbnails:

```bash
uv run aliexpress_export.py \
  --start-date 2026-03-01 \
  --end-date 2026-03-31 \
  --download-pdfs \
  --xlsx
```

Import cookies from a HAR file instead:

```bash
uv run aliexpress_export.py \
  --start-date 2026-03-01 \
  --end-date 2026-03-31 \
  --import-har "/path/to/orders.har"
```

Export directly from HAR capture(s) without making live API calls:

```bash
uv run aliexpress_export.py \
  --start-date 2026-03-01 \
  --end-date 2026-03-31 \
  --input-har "/path/to/list.har" \
  --input-har "/path/to/detail.har" \
  --download-pdfs \
  --xlsx
```

## Notes

- The login step is expected to happen in your own browser, not through Playwright automation.
- `--setup` reads cookies from your normal browser profile. Keep that browser logged into AliExpress while running setup.
- The stored auth artifact is a cookie jar, not a single bearer token. The MTop request signing token is derived from those cookies.
- The main order CSV is row-oriented and keeps the full AliExpress payloads in JSON columns. The separate order-line CSV is the easiest place to inspect product titles.
- The XLSX workbook is optional and keeps thumbnails in a dedicated `Thumbnail` column. If image downloads fail, the workbook is still written without those embedded images.
- Some sessions may expire and require a fresh HAR or Firefox-cookie import.
