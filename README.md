# Stock Market Toolkit

A small collection of focused Python scripts for retrieving stock-market data and analyzing the **distribution of returns (DoR)** of an asset across daily, weekly, monthly and quarterly timeframes.

The first (and currently only) workflow ports an Excel DoR template into Python: fetch OHLC, compute returns, bin them around the mean ± k·σ, and emit per-asset stats plus a multi-sheet Excel report covering a configurable universe of assets.

## Layout

```
stock-market-toolkit/
├── scripts/        # Python programs (one .py per goal)
├── notebooks/      # Matching .ipynb files for exploration
├── config/         # Asset universe (assets.csv)
├── raw-data/       # As-fetched yfinance CSVs
├── data/           # Cleaned CSVs with returns columns
└── output/         # Stats JSON, bin CSVs, xlsx reports
```

File-name convention for cached data: `{TICKER}_{d|w|m|q}_{YYYY-MM-DD}.csv`.

## Setup

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Dependencies: `yfinance`, `pandas`, `numpy`, `matplotlib`, `jupyter`, `openpyxl`.

## Creating a brand-new report from scratch

End-to-end recipe for building a fresh Distribution-of-Returns workbook over your own list of assets:

1. **Define the universe.** Either overwrite `config/assets.csv` or create a new file with the same columns:

   ```csv
   ticker,yfinance_symbol,asset_class,display_name
   AAPL,AAPL,Equity - Large Cap,Apple
   BTCUSD,BTC-USD,Crypto,Bitcoin
   GOLD,GC=F,Commodity,Gold Futures
   ```

   - `ticker` — short name used in cached filenames (uppercase, no spaces)
   - `yfinance_symbol` — exact symbol yfinance recognises (`^GSPC`, `EURUSD=X`, `BTC-USD`, `GC=F`, …)
   - `asset_class` / `display_name` — labels used in the Summary sheet

2. **(Optional) Start with a clean slate.** The pipeline tags every cached file with today's date, so old runs don't collide — but if you want a tidy `output/` folder you can wipe the gitignored data dirs:

   ```bash
   rm -f raw-data/* data/* output/*
   ```

3. **Run the orchestrator.** Point it at your config (omit `--config` to use `config/assets.csv`):

   ```bash
   python scripts/dor_run.py --config config/my_universe.csv --allow-empty
   ```

   - `--allow-empty` skips assets that fail to fetch (delisted tickers, bad symbols) instead of aborting the run.
   - Add `--out path/to/report.xlsx` to override the default output path.

   For each asset this fetches d/w/m OHLC from yfinance, derives quarterly bars, and writes per-asset stats and bin CSVs into `output/`.

4. **Open the report.** It lands at `output/Distribution_of_Returns_{YYYY-MM-DD}.xlsx` (or wherever `--out` pointed). Sheet order: **Summary** (raw values) → **Grading** (A–E quintile heat-map across the universe) → one detail tab per asset and timeframe.

5. **Iterate cheaply.** If you only want to tweak the report layout or bin scheme, re-run with `--skip-fetch` to reuse today's cached CSVs instead of hitting yfinance again:

   ```bash
   python scripts/dor_run.py --config config/my_universe.csv --skip-fetch
   ```

   `--skip-fetch` requires the `data/{TICKER}_{tf}_{today}.csv` files to already exist for every asset and timeframe — i.e. it only works on the same calendar day as the original fetch.

## Scripts

### `dor.py` — single-asset fetch & clean
Interactive helper: prompts for a ticker and a timeframe (`d`/`w`/`m`), pulls full history from yfinance, drops the in-progress current bar and any corrupt OHLC rows, computes `C-C`, `H-L` and (daily only) `O-C` returns, and writes:
- `raw-data/{TICKER}_{tf}_{date}.csv` — untouched yfinance output
- `data/{TICKER}_{tf}_{date}.csv` — cleaned + returns columns
- `output/{TICKER}_{tf}_{date}_summary.json` — short metadata report

```bash
python scripts/dor.py
```

Also exposes `fetch_prices`, `clean_data`, `descriptive_stats`, `save_results` for reuse.

### `dor_quarterly.py` — derive quarterly bars from monthly
yfinance does not expose quarterly bars. This module groups monthly OHLC into 3-month windows **anchored on the latest closed monthly bar** (not calendar quarters), aggregates Open/High/Low/Close and recomputes returns. Used by the orchestrator.

### `dor_bins.py` — bin distribution & positive-return stats
Mirrors the COUNTIF/COUNTIFS bin scheme from the Excel template. Edges are `μ + k·σ` for C-C and O-C returns and `k·σ` (no mean offset) for H-L returns, with different `k` lists for daily/weekly vs. monthly/quarterly. Returns a tidy DataFrame of `(label, edge, frequency, probability, cumulative)` plus a positive-returns summary block (avg, count, frequency, frequency-adjusted return).

### `dor_run.py` — multi-asset orchestrator
Runs the full pipeline over every row in `config/assets.csv`:
1. Fetch d/w/m bars from yfinance (or read cached CSVs with `--skip-fetch`)
2. Derive quarterly bars from the cleaned monthly frame
3. For every applicable return column at each timeframe, compute descriptive stats, bins, and positive-return analysis
4. Persist per-asset `*_stats.json` and `*_bins.csv` files
5. Build the consolidated workbook via `dor_report.write_workbook`

```bash
python scripts/dor_run.py                # fetch + report (default config)
python scripts/dor_run.py --skip-fetch   # rebuild report from today's cached CSVs
python scripts/dor_run.py --allow-empty  # skip delisted/broken tickers instead of aborting
python scripts/dor_run.py --config path/to/assets.csv --out path/to/report.xlsx
```

Default report path: `output/Distribution_of_Returns_{YYYY-MM-DD}.xlsx`.

### `dor_report.py` — xlsx writer
Builds the final workbook with three layers:
- **Summary sheet** — C-C standard deviation and average H-L return per asset, across all four timeframes, formatted as percentages.
- **Grading sheet** — same eight metrics, but each cell is replaced by a quintile grade **A–E** computed across the universe in this report, with a green→red colour scale. A cell shows e.g. `B  (1.50%)`: the letter is the rank bucket (A = lowest 20%, E = highest 20%); the value is the underlying number. Ranks are recomputed per column, so an asset can be an A on daily volatility and a D on quarterly. Use this sheet for at-a-glance comparison across the universe.
- **Per-asset detail sheets** — one per asset per timeframe, with the OHLC + returns table on the left and per-return blocks (bin distribution, descriptive stats, positive-return analysis) on the right.

### `verify_against_excel.py` — Excel parity check
Loads a copy of the original `Distribution_of_Returns_Template.xlsx`, runs the Python pipeline on the same OHLC, and asserts every descriptive-stats cell matches to floating-point tolerance. Useful when changing any of the math.

```bash
python scripts/verify_against_excel.py [TEMPLATE_PATH]
```

## Configuring the asset universe

Edit `config/assets.csv`. Each row defines one asset:

| ticker | yfinance_symbol | asset_class | display_name |
|--------|-----------------|-------------|--------------|
| EURUSD | EURUSD=X | FX Major | EUR/USD |
| SP500 | ^GSPC | Equity Index | S&P 500 |
| PCAR | PCAR | Equity - Large Cap | Paccar |

`ticker` is the local short name used in filenames; `yfinance_symbol` is what gets sent to the API.

## Notebooks

`notebooks/` mirrors the scripts and is used to step through the pipeline interactively and inspect intermediate frames:
- `dor.ipynb` — single-asset walk-through
- `dor_full_pipeline.ipynb` — end-to-end orchestrator
- `dor_summary_table.ipynb` — exploration of the consolidated summary
