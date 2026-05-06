"""Multi-asset Distribution-of-Returns pipeline.

For each asset in config/assets.csv:
  1. Fetch d/w/m OHLC from yfinance (or read cached CSVs with --skip-fetch)
  2. Derive quarterly from monthly via dor_quarterly.monthly_to_quarterly
  3. For every (timeframe, return column): compute descriptive_stats, bins,
     and positive-return analysis
  4. Persist per-asset stats JSON and bins CSV

Step 5 (separate commit) wires the xlsx report on top of these artefacts.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from dor import DATA_DIR, OUTPUT_DIR, clean_data, descriptive_stats, fetch_prices, save_results
from dor_bins import compute_bins, positive_return_stats, scheme_for
from dor_quarterly import monthly_to_quarterly

CONFIG_DEFAULT = ROOT / "config" / "assets.csv"
TIMEFRAMES = ["d", "w", "m", "q"]
RETURN_COLS_BY_TF = {
    "d": ["C-C Returns", "H-L Returns", "O-C Returns"],
    "w": ["C-C Returns", "H-L Returns"],
    "m": ["C-C Returns", "H-L Returns"],
    "q": ["C-C Returns", "H-L Returns"],
}


@dataclass
class Asset:
    ticker: str
    yfinance_symbol: str
    asset_class: str
    display_name: str


def load_assets(path: Path) -> list[Asset]:
    with path.open() as f:
        reader = csv.DictReader(f)
        return [Asset(**{k: v.strip() for k, v in row.items()}) for row in reader]


def cleaned_csv_path(ticker: str, tf: str, today: str) -> Path:
    return DATA_DIR / f"{ticker}_{tf}_{today}.csv"


def stats_json_path(ticker: str, tf: str, today: str) -> Path:
    return OUTPUT_DIR / f"{ticker}_{tf}_{today}_stats.json"


def bins_csv_path(ticker: str, tf: str, today: str) -> Path:
    return OUTPUT_DIR / f"{ticker}_{tf}_{today}_bins.csv"


def read_cleaned(ticker: str, tf: str, today: str) -> pd.DataFrame:
    """Read a cleaned CSV back, restoring the Date string index."""
    df = pd.read_csv(cleaned_csv_path(ticker, tf, today), dtype={"Date": str})
    return df.set_index("Date")


def fetch_and_save(asset: Asset, tf: str, today: str) -> pd.DataFrame:
    """Fetch from yfinance and persist raw + cleaned. Returns the cleaned frame."""
    raw = fetch_prices(asset.yfinance_symbol, tf)
    save_results(asset.ticker, tf, raw)  # writes raw, cleaned, summary.json
    return read_cleaned(asset.ticker, tf, today)


def derive_and_save_quarterly(ticker: str, today: str) -> pd.DataFrame:
    """Derive quarterly from cached monthly, persist as data/{ticker}_q_{today}.csv."""
    monthly = read_cleaned(ticker, "m", today)
    quarterly = monthly_to_quarterly(monthly)
    quarterly.to_csv(cleaned_csv_path(ticker, "q", today))
    return quarterly


def compute_one(frame: pd.DataFrame, tf: str) -> dict:
    """Compute stats + bins + positives for every applicable return column at one timeframe."""
    out = {}
    for col in RETURN_COLS_BY_TF[tf]:
        if col not in frame.columns:
            continue
        series = frame[col]
        out[col] = {
            "stats": descriptive_stats(series),
            "positives": positive_return_stats(series),
            "bins": compute_bins(series, scheme_for(col, tf)),
        }
    return out


def persist_one(ticker: str, tf: str, today: str, results: dict) -> None:
    # JSON: stats + positives only (bins go to a separate CSV).
    json_payload = {
        "ticker": ticker,
        "timeframe": tf,
        "fetched_on": today,
        "returns": {
            col: {"stats": payload["stats"], "positives": payload["positives"]}
            for col, payload in results.items()
        },
    }
    stats_json_path(ticker, tf, today).write_text(json.dumps(json_payload, indent=2, default=float))

    # Bins CSV: one long-format file per (ticker, tf), with return_col as a column.
    rows = []
    for col, payload in results.items():
        bin_df = payload["bins"].copy()
        bin_df.insert(0, "return_col", col)
        rows.append(bin_df)
    if rows:
        pd.concat(rows, ignore_index=True).to_csv(bins_csv_path(ticker, tf, today), index=False)


def process_asset(asset: Asset, today: str, skip_fetch: bool, allow_empty: bool) -> dict | None:
    print(f"  [{asset.ticker}] yfinance={asset.yfinance_symbol}")
    frames: dict[str, pd.DataFrame] = {}
    try:
        for tf in ("d", "w", "m"):
            if skip_fetch:
                frames[tf] = read_cleaned(asset.ticker, tf, today)
            else:
                frames[tf] = fetch_and_save(asset, tf, today)
        frames["q"] = derive_and_save_quarterly(asset.ticker, today)
    except Exception as exc:
        if allow_empty:
            print(f"    SKIPPED: {exc}")
            return None
        raise

    asset_results: dict = {
        "ticker": asset.ticker,
        "yfinance_symbol": asset.yfinance_symbol,
        "asset_class": asset.asset_class,
        "display_name": asset.display_name,
        "frames": frames,
        "by_tf": {},
    }
    for tf in TIMEFRAMES:
        results = compute_one(frames[tf], tf)
        persist_one(asset.ticker, tf, today, results)
        asset_results["by_tf"][tf] = results
    return asset_results


def run(config_path: Path, skip_fetch: bool, allow_empty: bool) -> list[dict]:
    today = date.today().isoformat()
    assets = load_assets(config_path)
    print(f"Loaded {len(assets)} assets from {config_path}")
    out = []
    for asset in assets:
        res = process_asset(asset, today, skip_fetch=skip_fetch, allow_empty=allow_empty)
        if res is not None:
            out.append(res)
    print(f"\nProcessed {len(out)} of {len(assets)} assets.")
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=CONFIG_DEFAULT)
    parser.add_argument("--skip-fetch", action="store_true",
                        help="Read cached cleaned CSVs from data/ instead of hitting yfinance.")
    parser.add_argument("--allow-empty", action="store_true",
                        help="Skip assets that fail to fetch (e.g. delisted) instead of aborting.")
    args = parser.parse_args()
    run(args.config, args.skip_fetch, args.allow_empty)


if __name__ == "__main__":
    main()
