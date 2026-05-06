# Stock Market Toolkit

A collection of small, focused Python programs to retrieve and analyze stock market data.

## Project Goals

Build a set of vital files for stock market analysis. Each program targets a specific, well-defined goal.

## Guiding Principles

- **Simple over clever**: each program should be small, readable, and do one thing.
- **Plain Python scripts** for the actual work (data retrieval, computation, transformation).
- **Companion Jupyter notebook** alongside each script — used to run the script step by step and visualize intermediate results when useful.
- **Structured output**: each program produces a result file (JSON, CSV, or a short report) so results can be inspected, shared, or fed into the next step.
- **No premature abstraction**: don't build frameworks or generic pipelines until concrete needs justify them.

## Suggested Layout

```
stock-market-toolkit/
├── scripts/        # Python programs (one .py per goal)
├── notebooks/      # Matching .ipynb files for exploration / visualization
├── raw-data/       # Raw data pulled from APIs, saved as-is
│                   # File name convention: {TICKER}_{timeframe}_{YYYY-MM-DD}.csv
├── data/           # Cleaned/manipulated data ready for analysis
│                   # Same naming convention as raw-data
├── output/         # Generated reports, JSON results, charts
└── CLAUDE.md
```

## Workflow per Task

1. Define the specific goal in plain language.
2. Write the minimal Python script that achieves it.
3. Create a matching notebook that imports the script (or its functions) and visualizes the steps.
4. Save the result to `output/` as JSON or a short report.
