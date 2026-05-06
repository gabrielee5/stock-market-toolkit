"""Derive quarterly OHLC frames from monthly OHLC.

yfinance does not expose quarterly bars, so the project rule is:
- anchor on the latest closed monthly bar
- group every 3 months going BACKWARD from the anchor (NOT calendar quarters)
- per group: Open = first month's open, Close/Adj Close = last month's,
  High = max of all 3 months, Low = min of all 3 months
- drop the oldest group if it has fewer than 3 months
"""

import pandas as pd

OHLC_COLS = ["Open", "High", "Low", "Close", "Adj Close"]


def monthly_to_quarterly(monthly_clean: pd.DataFrame) -> pd.DataFrame:
    """Aggregate a cleaned monthly frame into a cleaned quarterly frame.

    Input expectations (matching `dor.clean_data` output):
      - DataFrame indexed by Date strings in 'DD-MM-YY' format
      - sorted newest-first
      - includes columns Open, High, Low, Close, Adj Close (returns columns
        are ignored if present and recomputed at quarterly resolution)

    Output: same conventions, with two return columns ('C-C Returns',
    'H-L Returns') recomputed for the quarterly bars. The oldest row is
    dropped (its C-C return is undefined).
    """
    if not all(c in monthly_clean.columns for c in OHLC_COLS):
        missing = [c for c in OHLC_COLS if c not in monthly_clean.columns]
        raise ValueError(f"Monthly frame missing columns: {missing}")

    # Re-sort ascending so groupby-by-3-from-the-right is intuitive.
    asc = monthly_clean[OHLC_COLS].sort_index(
        ascending=True,
        key=lambda idx: pd.to_datetime(idx, format="%d-%m-%y"),
    )
    n = len(asc)
    if n < 3:
        raise ValueError(f"Need at least 3 monthly rows to form a quarter, got {n}.")

    # Right-anchored grouping: group_id = (n - 1 - i) // 3 with i = ascending row index.
    # The latest month gets group_id 0, walking backward.
    pos = pd.RangeIndex(n)
    group_id = (n - 1 - pos) // 3
    asc = asc.assign(_g=group_id.values)

    # Drop incomplete oldest group (only the smallest group_id can be partial after right-anchoring).
    sizes = asc.groupby("_g").size()
    full_groups = sizes[sizes == 3].index
    asc = asc[asc["_g"].isin(full_groups)]

    aggregated = asc.groupby("_g").agg({
        "Open":      "first",
        "High":      "max",
        "Low":       "min",
        "Close":     "last",
        "Adj Close": "last",
    })
    # Date label = the LATEST month in the group (i.e. the last row of the ascending slice).
    last_dates = asc.reset_index().groupby("_g")["Date"].last()
    aggregated["Date"] = last_dates
    aggregated = aggregated.set_index("Date")

    # Re-sort newest-first to match project convention.
    quarterly = aggregated.sort_index(
        ascending=False,
        key=lambda idx: pd.to_datetime(idx, format="%d-%m-%y"),
    )

    # Recompute returns on the newest-first quarterly frame (no O-C at quarterly).
    # Oldest row's C-C is NaN (no prior quarter). Keep the row so H-L count matches
    # Excel; downstream per-column dropna() in stat/bin helpers handles the NaN.
    prev_adj = quarterly["Adj Close"].shift(-1)
    quarterly["C-C Returns"] = (quarterly["Adj Close"] - prev_adj) / prev_adj
    quarterly["H-L Returns"] = (quarterly["High"] - quarterly["Low"]) / quarterly["Low"]

    return quarterly
