"""Distribution-of-returns bin computation and positive-return analysis.

Mirrors the COUNTIF / COUNTIFS scheme used by the Excel template:
    row 4:    COUNTIF(returns, "<=" & edge[0])               -> left tail
    rows 5+:  COUNTIFS(returns, "<=" & edge[i], ">" & edge[i-1])  -> middle bands
    row N:    COUNTIF(returns, ">" & edge[-1])               -> right tail

C-C / O-C bins use edges of the form (mu + k*sigma).
H-L bins use edges of the form (k*sigma) -- no mean offset, since H-L is
always non-negative. The first H-L edge is 0 (literal in the template).
"""

import pandas as pd

# k-multipliers per scheme. Each list defines the bin edges; the function
# adds one left-tail row and one right-tail row, so total frequency rows = len(ks) + 1.
BIN_K = {
    "cc_daily_weekly":      [-3.0, -2.25, -1.5, -0.75, 0.0, 0.75, 1.5, 2.25, 3.0],
    "cc_monthly_quarterly": [-3.0, -2.4, -1.8, -1.2, -0.6, 0.0, 0.6, 1.2, 1.8, 2.4, 3.0],
    "hl_daily_weekly":      [0.0, 0.4, 0.8, 1.2, 1.6, 2.0, 2.4, 2.8, 3.2],
    "hl_monthly_quarterly": [0.0, 0.4, 0.8, 1.2, 1.6, 2.0, 2.4, 2.8, 3.2, 3.6, 4.0],
    "oc_daily":             [-3.0, -2.25, -1.5, -0.75, 0.0, 0.75, 1.5, 2.25, 3.0],
}


def scheme_for(return_col: str, timeframe: str) -> str:
    """Pick the bin scheme matching a (return column, timeframe) pair."""
    tf = timeframe.lower()
    if return_col == "C-C Returns":
        return "cc_daily_weekly" if tf in ("d", "w") else "cc_monthly_quarterly"
    if return_col == "H-L Returns":
        return "hl_daily_weekly" if tf in ("d", "w") else "hl_monthly_quarterly"
    if return_col == "O-C Returns":
        return "oc_daily"
    raise ValueError(f"No bin scheme for column '{return_col}'.")


def compute_bins(returns: pd.Series, scheme: str) -> pd.DataFrame:
    """Bin a returns series with the named scheme.

    Returns a DataFrame with one row per bin (in template order: left tail,
    middle bands ascending, right tail). Columns:
      - k_label: human-readable bin label (e.g. "<= -3.00σ", "(0.4σ, 0.8σ]", "> 3.20σ")
      - edge:    upper edge of the bin (None for the right-tail row)
      - frequency: count of returns in the bin
      - probability: frequency / total count
      - cumulative: running sum of probability
    """
    if scheme not in BIN_K:
        raise ValueError(f"Unknown scheme '{scheme}'. Choose: {list(BIN_K)}")

    r = returns.dropna().to_numpy()
    n = len(r)
    if n == 0:
        raise ValueError("Cannot compute bins on an empty series.")

    mu = r.mean()
    sigma = r.std(ddof=1)
    ks = BIN_K[scheme]
    is_hl = scheme.startswith("hl_")
    edges = [(k * sigma) if is_hl else (mu + k * sigma) for k in ks]

    rows = []
    # Left tail: <= edges[0]
    freq = int((r <= edges[0]).sum())
    rows.append({
        "k_label": f"<= {ks[0]:g}σ" if not is_hl or ks[0] != 0 else "<= 0",
        "edge": edges[0],
        "frequency": freq,
    })
    # Middle: (edges[i-1], edges[i]]
    for i in range(1, len(edges)):
        freq = int(((r > edges[i - 1]) & (r <= edges[i])).sum())
        rows.append({
            "k_label": f"({ks[i-1]:g}σ, {ks[i]:g}σ]",
            "edge": edges[i],
            "frequency": freq,
        })
    # Right tail: > edges[-1]
    freq = int((r > edges[-1]).sum())
    rows.append({
        "k_label": f"> {ks[-1]:g}σ",
        "edge": None,
        "frequency": freq,
    })

    df = pd.DataFrame(rows)
    df["probability"] = df["frequency"] / n
    df["cumulative"] = df["probability"].cumsum()
    return df


def positive_return_stats(returns: pd.Series) -> dict:
    """Mirror the Excel template's row-32/34 positive-return block."""
    r = returns.dropna()
    pos = r[r > 0]
    count = int(r.count())
    count_pos = int(pos.count())
    avg_pos = float(pos.mean()) if count_pos else 0.0
    freq_pct = (count_pos / count) if count else 0.0
    return {
        "average_positive": avg_pos,
        "count_positive": count_pos,
        "freq_pct_positive": freq_pct,
        "freq_adjusted_return": freq_pct * avg_pos,
    }
