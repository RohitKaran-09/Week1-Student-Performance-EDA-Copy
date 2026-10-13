# Task 1 — Automotive Market Overview & Trend Analysis

A research-driven review of India's listed automotive sector, combining publicly available historical equity prices with official industry sales and policy information. The code downloads five years of daily price data, calculates comparable trend and risk measures, produces charts, and generates a Word report and submission-description document.

## Scope
- Benchmark: Nifty 50 (`^NSEI`). The code also checks the official Nifty Auto Yahoo symbol (`^CNXAUTO`); if its historical series is unavailable, it uses a clearly labeled equal-weighted proxy from the included auto stocks instead.
- Representative listed companies: Maruti Suzuki, Mahindra & Mahindra, Bajaj Auto, Eicher Motors, Hero MotoCorp, TVS Motor, Ashok Leyland, Bharat Forge, Bosch, and Exide Industries.
- Historical data: five-year daily series where available via Yahoo Finance's public market-data interface. Adjusted close is used when present; index closing levels are used as supplied. In the current run Yahoo supplied only one usable `^CNXAUTO` observation, so the historical comparison uses an equal-weighted daily-return basket of the downloaded auto equities, not a synthetic or official index series.
- Industry context: SIAM FY2025–26 and August 2026 releases, official GST policy notices, PM E-DRIVE dashboard, corporate FY26 results and current business news.

## Run
```bash
python -m pip install -r requirements.txt
python automotive_market_analysis.py
```

The script requires live market data and stops if core benchmark data are unavailable; it does not use synthetic prices. Outputs include the DOCX report, separate Word submission description, historical-price CSV, return/risk metrics, and charts.

## Interpretation
Historical returns and risk measures are descriptive, not forecasts. The report is educational market research, not individualized investment advice. Equity adjusted-price series and index price-level series are not perfectly identical return concepts; this limitation is disclosed in the report.
