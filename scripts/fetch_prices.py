"""Fetch all available price data for a ticker at a chosen timeframe."""

from datetime import date
from pathlib import Path
import json

import pandas as pd
import yfinance as yf

INTERVAL_MAP = {
    "d": "1d",
    "w": "1wk",
    "m": "1mo",
}

ROOT = Path(__file__).resolve().parent.parent
RAW_DATA_DIR = ROOT / "raw-data"
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

KEEP_COLUMNS = ["Open", "High", "Low", "Close", "Adj Close"]


def fetch_prices(ticker: str, timeframe: str):
    timeframe = timeframe.lower().strip()
    if timeframe not in INTERVAL_MAP:
        raise ValueError(
            f"Unknown timeframe '{timeframe}'. Choose: {list(INTERVAL_MAP)}"
        )

    interval = INTERVAL_MAP[timeframe]
    df = yf.Ticker(ticker).history(period="max", interval=interval, auto_adjust=False)
    if df.empty:
        raise ValueError(f"No data returned for ticker '{ticker}'.")
    return df


def clean_data(df):
    cleaned = df.copy()

    cleaned.index = cleaned.index.strftime("%d-%m-%y")
    cleaned.index.name = "Date"

    cleaned = cleaned[KEEP_COLUMNS]

    cleaned = cleaned.sort_index(
        ascending=False,
        key=lambda idx: pd.to_datetime(idx, format="%d-%m-%y"),
    )

    prev_adj_close = cleaned["Adj Close"].shift(-1)
    cleaned["C-C Returns"] = (cleaned["Adj Close"] - prev_adj_close) / prev_adj_close * 100
    cleaned["H-L Returns"] = (cleaned["High"] - cleaned["Low"]) / cleaned["Low"] * 100
    cleaned["O-C Returns"] = (cleaned["Close"] - cleaned["Open"]) / cleaned["Open"] * 100

    return cleaned


def save_results(ticker: str, timeframe: str, df):
    RAW_DATA_DIR.mkdir(exist_ok=True)
    DATA_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    ticker_safe = ticker.upper()
    today = date.today().isoformat()
    base_name = f"{ticker_safe}_{timeframe}_{today}"

    raw_csv_path = RAW_DATA_DIR / f"{base_name}.csv"
    df.to_csv(raw_csv_path)

    cleaned = clean_data(df)
    clean_csv_path = DATA_DIR / f"{base_name}.csv"
    cleaned.to_csv(clean_csv_path)

    summary = {
        "ticker": ticker_safe,
        "timeframe": timeframe,
        "interval": INTERVAL_MAP[timeframe],
        "fetched_on": today,
        "rows": len(df),
        "start": df.index.min().isoformat(),
        "end": df.index.max().isoformat(),
        "raw_csv_path": str(raw_csv_path.relative_to(ROOT)),
        "clean_csv_path": str(clean_csv_path.relative_to(ROOT)),
    }
    report_path = OUTPUT_DIR / f"{base_name}_summary.json"
    report_path.write_text(json.dumps(summary, indent=2))
    return raw_csv_path, clean_csv_path, report_path, summary


def prompt_inputs():
    ticker = input("Ticker (e.g. AAPL): ").strip()
    timeframe = input("Timeframe [d/w/m]: ").strip().lower()
    return ticker, timeframe


def main():
    ticker, timeframe = prompt_inputs()
    df = fetch_prices(ticker, timeframe)
    raw_csv_path, clean_csv_path, report_path, summary = save_results(
        ticker, timeframe, df
    )
    print(f"Fetched {summary['rows']} rows ({summary['start']} -> {summary['end']})")
    print(f"Raw:     {raw_csv_path}")
    print(f"Clean:   {clean_csv_path}")
    print(f"Report:  {report_path}")


if __name__ == "__main__":
    main()
