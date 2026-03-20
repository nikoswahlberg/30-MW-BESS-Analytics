# BESS Revenue & Investment Analysis

A complete data pipeline modelling a **30 MW / 36 MWh** battery energy
storage system operating in the Finnish reserve markets.

The pipeline fetches real market prices, runs a physically constrained
State-of-Charge simulation across all relevant Finnish reserve markets, and
produces an investment model with IRR, NPV, and scenario analysis.

---

## Project Structure

```
bess_project/
│
├── data_retrieval_fingrid.py   # Step 1 — fetch prices, run SoC simulation, store to DB
├── data_analytics.py           # Step 2 — charts + 5-sheet investment Excel workbook
├── generate_dashboard.py       # Step 3 — one-page executive Excel dashboard
│
├── BESS_Analysis.xlsx          # Full analysis workbook (5 sheets, technical)
├── BESS_Dashboard.xlsx         # Executive summary dashboard (non-technical)
│
├── BESS_Revenue_Comparison.png
├── Monthly_Revenue_2025.png
├── Cumulative_Revenue_2025.png
├── BESS_Strategy_Trend_2025.png
├── BESS_SoC_Profile_2025.png
│
├── docker-compose.yml          # PostgreSQL container
├── requirements.txt
└── .env                        # API keys and DB credentials (not committed to git)
```

---

## Setup

### 1. Clone and create virtual environment

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure `.env`

```env
# Fingrid Open Data API — reserve market prices (FCR-N, FCR-D, mFRR)
# Register at: https://data.fingrid.fi
FINGRID_API_KEY=your_fingrid_key_here

# ENTSO-E Transparency Platform — Finnish day-ahead spot prices
# 1. Register at https://transparency.entsoe.eu
# 2. Email transparency@entsoe.eu  subject: "Restful API access"
# 3. Key appears in your account under "Web API Security Token"
ENTSOE_API_KEY=your_entsoe_key_here

# Database (postgres or sqlite)
DB_TYPE=postgres

# PostgreSQL credentials (only needed when DB_TYPE=postgres)
POSTGRES_DB=bess_db
POSTGRES_USER=your_user
POSTGRES_PASSWORD=your_password
POSTGRES_HOST=localhost
POSTGRES_PORT=5433
```

### 3. Start PostgreSQL (optional — skip if using SQLite)

```bash
docker-compose up -d
```

---

## Running the Pipeline

Run the three scripts in order:

```bash
# Step 1: Fetch all market data and run SoC simulation (~5–10 min due to API rate limits)
python data_retrieval_fingrid.py

# Step 2: Generate charts and full investment workbook
python data_analytics.py

# Step 3: Generate the executive dashboard
python generate_dashboard.py
```

---

## Data Sources

| Data | Source | Dataset |
|---|---|---|
| Day-ahead spot price (FI) | ENTSO-E Transparency Platform | `query_day_ahead_prices('FI')` |
| FCR-N hourly price | Fingrid Open Data | Dataset 317 |
| FCR-D up hourly price | Fingrid Open Data | Dataset 318 |
| FCR-D down hourly price | Fingrid Open Data | Dataset 283 |
| mFRR up-regulation price | Fingrid Open Data | Dataset 244 |
| mFRR down-regulation price | Fingrid Open Data | Dataset 106 |

> **Note:** Fingrid dataset 245 ("Wind power generation forecast") is sometimes
> incorrectly cited as a spot price source. It is **not** a price dataset.
> Finnish day-ahead prices must come from ENTSO-E.

---

## Market Model

### FCR Markets (Capacity — earned every hour within SoC window)

| Market | SoC Requirement | Revenue type |
|---|---|---|
| FCR-N | 35–65% SoC | €/MW/h capacity fee |
| FCR-D Up | ≥ 10 MWh (28%) SoC | €/MW/h capacity fee |
| FCR-D Down | ≤ 75% SoC | €/MW/h capacity fee |

### Energy Markets (Activation — dispatched on demand)

| Market | Activation probability | Revenue type |
|---|---|---|
| Spot arbitrage | Discharge when price > €80/MWh | €/MWh energy payment |
| mFRR Up | ~12% of hours (Fingrid decides) | €/MWh energy payment |
| mFRR Down | ~6% of hours (Fingrid decides) | €/MWh for absorption |

### State-of-Charge Constraints

- Operating range: **10–90%** of capacity (3.6–32.4 MWh)
- Spot and mFRR discharge only draw from headroom **above the FCR-N floor**
  (12.6 MWh), preserving FCR capability for the following hour
- Forced grid recharge when SoC falls below the FCR-D up minimum (10 MWh)

---

## Outputs

### `BESS_Dashboard.xlsx` — Executive Summary (non-technical)

One-sheet overview intended for stakeholders who will not read the terminal output:
- KPI cards: operational revenue, IRR, NPV, payback
- Market participation breakdown table with hour counts and revenue shares
- Monthly revenue comparison across all three strategies
- 20-year condensed cash flow table (operational estimate)
- All assumptions and data sources in the footer

### `BESS_Analysis.xlsx` — Full Technical Workbook (5 sheets)

| Sheet | Contents |
|---|---|
| Raw Price Data | All 6 market price columns + legacy revenue columns |
| SoC Simulation | Hour-by-hour market selection, revenue, state of charge |
| Market Breakdown | Revenue and hours per market; monthly breakdown |
| Investment Model | 20-year DCF, IRR/NPV/payback (model max + operational estimate) |
| Scenarios | Bull / Base / Bear sensitivity (CAPEX ±33%, revenue ±25%, WACC 7–10%) |

### Charts

| File | Description |
|---|---|
| `BESS_Revenue_Comparison.png` | Hourly revenue: naive strategies vs SoC-optimised |
| `Monthly_Revenue_2025.png` | Monthly totals, all three strategies side by side |
| `Cumulative_Revenue_2025.png` | Running total across the year |
| `BESS_Strategy_Trend_2025.png` | 7-day rolling average: volatility vs stability |
| `BESS_SoC_Profile_2025.png` | Battery SoC over the year + daily dominant market |

---

## Financial Assumptions

| Parameter | Value | Source |
|---|---|---|
| CAPEX | €12 M (€400/kW) | Ember "How cheap is battery storage" Jan 2026 |
| OPEX | €180k/yr fixed + 0.5% of revenue | Standard BESS O&M range |
| Revenue growth | 3% p.a. | Fingrid: reserve demand +134% forecast over 5 yrs |
| Degradation | 2% p.a. capacity loss | LFP chemistry, real-world data |
| Cell replacement | €3.5 M at year 12 | Cells only; inverters and BOS retained |
| WACC | 8% | Infrastructure equity fund target |
| Corporation tax | 20% | Finland |
| Depreciation | Straight-line, 10 years | |
| Project life | 20 years | |
| Operational discount | −20% on model revenue | Bidding friction, ~92% availability, min bid volumes |

### IRR benchmarks

- **3–7% unlevered** — Western EU merchant BESS (Capstone DC, Nov 2025)
- **€40,700/MW/month** — Exilion Finland, H2 2023 (exceptional market conditions)
- Finland is structurally above the Western EU average due to higher price volatility
  and growing reserve market demand

---

## Known Model Limitations

1. **Perfect foresight** — the SoC optimiser knows all prices in advance.
   Real dispatch performance is typically 10–25% lower. The operational
   estimate (−20% haircut) is applied to all investment metrics.

2. **mFRR mutual exclusivity** — the model treats FCR and mFRR as mutually
   exclusive per hour. Real operators split capacity between markets
   simultaneously. This slightly understates total mFRR revenue.

3. **FCR activation drift** — FCR-D probabilistic activations use fixed
   fractions (2.5% up, 1.0% down). In high-volatility years this could
   reach 4–6%, increasing recharge costs modestly.

4. **24 NaN spot hours** — DST clock-change gaps from ENTSO-E.
   Treated as €0/MWh. Negligible annual impact.

5. **Capacity doubling** — The project has an option to double capacity to
   60 MW. This is not modelled but would benefit from significantly lower
   €/kW BOS costs on the existing site.

---

## Background

The modelled asset is a 30 MW / 36 MWh battery energy storage system
in Finland, fully operational since Q4 2024. It participates in Fingrid's
reserve markets and supports the integration of renewable energy into the
Finnish grid. The project is classified as EU Taxonomy-aligned (EU/2020/852).

The EU supports this activity through the InvestEU fund.