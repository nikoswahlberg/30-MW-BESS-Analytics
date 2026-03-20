"""
BESS Revenue Pipeline — data_retrieval_fingrid.py
==================================================
Spot prices: ENTSO-E Transparency Platform (entsoe-py)
Reserve prices: Fingrid Open Data API

Why ENTSO-E for spot prices?
  Fingrid's FAQ explicitly states they cannot publish hourly spot prices as
  the data is not owned by them. Dataset 245 on Fingrid is the wind power
  generation FORECAST in MW (15-min resolution), not €/MWh prices — using
  it as a price produced fictional revenues of €57 M vs a realistic ~€1–3 M.

Setup (one-time):
  1. Register at https://transparency.entsoe.eu
  2. Email transparency@entsoe.eu — subject: "Restful API access"
     Body: include your registered email address
  3. API key appears in your account under "Web API Security Token"
  4. Add to .env:  ENTSOE_API_KEY=your-key-here
  5. pip install entsoe-py  (or add to requirements.txt)

Finnish day-ahead spot prices (FI bidding zone) for 2025:
  Typical range: −10 to ~300 €/MWh on normal days; spikes to ~1,000 €/MWh
  during extreme cold-snap events. Median typically €30–80 €/MWh.
  If you see a median of >500, you are using the wrong data source.

Physical model: 30 MW / 36 MWh BESS  (Finland)
"""

import os
import time
import requests
import numpy as np
import pandas as pd
from sqlalchemy import create_engine
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.dates as mdates
from dotenv import load_dotenv

# entsoe-py — install with: pip install entsoe-py
try:
    from entsoe import EntsoePandasClient
    ENTSOE_AVAILABLE = True
except ImportError:
    ENTSOE_AVAILABLE = False
    print("⚠  entsoe-py not installed. Run: pip install entsoe-py")

# ---------------------------------------------------------------------------
# 1.  Configuration
# ---------------------------------------------------------------------------
load_dotenv()
FINGRID_API_KEY  = os.getenv("FINGRID_API_KEY")
ENTSOE_API_KEY   = os.getenv("ENTSOE_API_KEY")
DB_TYPE          = os.getenv("DB_TYPE", "sqlite").lower()

# --- Battery Physical Parameters ---
BESS_MW  = 30
BESS_MWH = 36

MIN_SOC  = BESS_MWH * 0.10    #  3.6 MWh
MAX_SOC  = BESS_MWH * 0.90    # 32.4 MWh
CHG_EFF  = 0.95
DCH_EFF  = 0.95

# FCR SoC feasibility windows
FCR_D_UP_MIN_SOC   = BESS_MW * (20 / 60)    # 10.0 MWh
FCR_N_MIN_SOC      = BESS_MWH * 0.35        # 12.6 MWh
FCR_N_MAX_SOC      = BESS_MWH * 0.65        # 23.4 MWh
FCR_D_DOWN_MAX_SOC = BESS_MWH * 0.75        # 27.0 MWh

# Arbitrage thresholds
SPOT_DISCHARGE_MIN = 80     # €/MWh — discharge only above this
SPOT_CHARGE_MAX    = 30     # €/MWh — charge only below this
MFRR_MIN_PREMIUM   = 1.10   # mFRR up must beat 110 % of spot

# FCR probabilistic activation (capacity markets — SoC drift only)
FCR_D_UP_ACT_FRAC   = 0.025
FCR_D_UP_ACT_MIN    = 15
FCR_D_DOWN_ACT_FRAC = 0.010
FCR_D_DOWN_ACT_MIN  = 15

# mFRR activation probabilities (energy activation market — Fingrid dispatches)
# mFRR UP: ~12 % of hours Fingrid activates up-regulation in Finland.
#   Source: Fingrid balancing statistics; typical Nordic range 8–18 %.
#   When activated, average duration ≈ 20 min → energy = BESS_MW × (20/60) MWh.
# mFRR DOWN: rarer (~6 % of hours); Fingrid absorbs excess generation.
# These fractions scale the theoretical revenue to a realistic dispatch volume.
# Without them the model overstates mFRR revenue by 5–10×.
MFRR_UP_ACT_PROB    = 0.12    # 12 % of hours actually activated
MFRR_UP_ACT_MIN     = 20      # average activation duration (minutes)
MFRR_DOWN_ACT_PROB  = 0.06    # 6 % of hours actually activated
MFRR_DOWN_ACT_MIN   = 20      # average activation duration (minutes)

# Output files
DB_NAME_SQLITE = "bess_model.db"
EXCEL_FILE     = "BESS_Analysis.xlsx"
CHART_FILE     = "BESS_Revenue_Comparison.png"

# Fingrid reserve market datasets (all confirmed correct)
FINGRID_DATASETS = {
    "FCR_N":      {"id": "317", "unit": "€/MW/h", "resample": None},
    "FCR_D_UP":   {"id": "318", "unit": "€/MW/h", "resample": None},
    "FCR_D_DOWN": {"id": "283", "unit": "€/MW/h", "resample": None},
    "mFRR_UP":    {"id": "244", "unit": "€/MWh",  "resample": "mean"},
    "mFRR_DOWN":  {"id": "106", "unit": "€/MWh",  "resample": "mean"},
}

# ---------------------------------------------------------------------------
# 2.  Database engine
# ---------------------------------------------------------------------------
def get_db_engine():
    if DB_TYPE == "postgres":
        user = os.getenv("POSTGRES_USER")
        pw   = os.getenv("POSTGRES_PASSWORD")
        host = os.getenv("POSTGRES_HOST")
        port = os.getenv("POSTGRES_PORT")
        db   = os.getenv("POSTGRES_DB")
        return create_engine(f"postgresql://{user}:{pw}@{host}:{port}/{db}")
    return create_engine(f"sqlite:///{DB_NAME_SQLITE}")


# ---------------------------------------------------------------------------
# 3.  Spot price fetch — ENTSO-E
# ---------------------------------------------------------------------------
def fetch_spot_prices(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Fetches Finnish day-ahead spot prices from ENTSO-E Transparency Platform.
    Returns DataFrame with columns ['Timestamp', 'Price_Spot'] in €/MWh.

    Requires: pip install entsoe-py
              ENTSOE_API_KEY in .env
    """
    if not ENTSOE_AVAILABLE:
        print("   [!] entsoe-py not installed — run: pip install entsoe-py")
        return pd.DataFrame()
    if not ENTSOE_API_KEY:
        print("   [!] ENTSOE_API_KEY missing in .env")
        print("       Register at https://transparency.entsoe.eu")
        print("       Email transparency@entsoe.eu — subject: 'Restful API access'")
        return pd.DataFrame()

    print("   Fetching FI day-ahead spot prices from ENTSO-E...", end=" ", flush=True)
    try:
        client = EntsoePandasClient(api_key=ENTSOE_API_KEY)
        start  = pd.Timestamp(start_date[:10], tz="Europe/Helsinki")
        end    = pd.Timestamp(end_date[:10],   tz="Europe/Helsinki") + pd.Timedelta(days=1)

        series = client.query_day_ahead_prices("FI", start=start, end=end)
        # Convert to UTC-naive hourly index matching Fingrid data
        series = series.tz_convert("UTC").tz_localize(None)
        series = series.resample("1h").mean()  # already hourly; safety resample

        df = series.reset_index()
        df.columns = ["Timestamp", "Price_Spot"]
        # Trim to requested year
        df = df[(df["Timestamp"] >= start_date[:10]) &
                (df["Timestamp"] <  end_date[:10])]
        print(f"Done. {len(df)} hourly rows.")
        return df

    except Exception as e:
        print(f"\n   [!] ENTSO-E fetch failed: {e}")
        return pd.DataFrame()


def verify_spot_prices(df: pd.DataFrame):
    """Sanity-checks spot prices and warns if values look like MW forecasts."""
    if "Price_Spot" not in df.columns or df.empty:
        return
    s = df["Price_Spot"].dropna()
    print(f"\n  ── Spot price sanity check (ENTSO-E FI) ────────────")
    print(f"     Min : {s.min():>10.2f} €/MWh")
    print(f"     P25 : {s.quantile(0.25):>10.2f} €/MWh")
    print(f"     Med : {s.median():>10.2f} €/MWh")
    print(f"     P75 : {s.quantile(0.75):>10.2f} €/MWh")
    print(f"     Max : {s.max():>10.2f} €/MWh")
    print(f"     Mean: {s.mean():>10.2f} €/MWh")
    if s.median() > 500:
        print("  ⚠  Median > 500 €/MWh — values seem too high for normal spot prices.")
    elif s.median() < 0:
        print("  ⚠  Negative median — check timezone or data alignment.")
    else:
        print("  ✓  Values look plausible for Finnish day-ahead spot prices.")
    print(f"  ────────────────────────────────────────────────────\n")


# ---------------------------------------------------------------------------
# 4.  Reserve price fetch — Fingrid
# ---------------------------------------------------------------------------
def fetch_fingrid_data(dataset_id: str, start_time: str, end_time: str,
                       label: str) -> pd.DataFrame:
    """
    Fetches all pages for a Fingrid dataset.
    Returns DataFrame ['Timestamp', 'Price_<label>'], resampled to hourly.
    """
    if not FINGRID_API_KEY:
        print(f"   [!] FINGRID_API_KEY missing — skipping {label}")
        return pd.DataFrame()

    url     = f"https://data.fingrid.fi/api/datasets/{dataset_id}/data"
    headers = {"x-api-key": FINGRID_API_KEY, "Accept": "application/json"}
    params  = {
        "pageSize": 20000, "startTime": start_time, "endTime": end_time,
        "sortBy": "startTime", "sortOrder": "asc",
    }

    all_data, page = [], 1
    while True:
        params["page"] = page
        print(f"   [{label}] Page {page}...", end=" ", flush=True)
        try:
            r = requests.get(url, headers=headers, params=params, timeout=30)
            if r.status_code == 429:
                print("Rate-limited — cooling 12s...")
                time.sleep(12)
                continue
            r.raise_for_status()
            batch = r.json().get("data", [])
            if not batch:
                print("Done.")
                break
            all_data.extend(batch)
            print(f"{len(all_data)} rows")
            page += 1
            time.sleep(1)
        except Exception as e:
            print(f"\n   [!] Error: {e}")
            break

    if not all_data:
        return pd.DataFrame()

    df = pd.DataFrame(all_data)
    df["Timestamp"] = pd.to_datetime(df["startTime"])
    df = df.rename(columns={"value": f"Price_{label}"})
    df = df[["Timestamp", f"Price_{label}"]].dropna()

    if FINGRID_DATASETS[label]["resample"] == "mean":
        df = df.set_index("Timestamp").resample("1h").mean().reset_index()

    return df


# ---------------------------------------------------------------------------
# 5.  SoC simulation
# ---------------------------------------------------------------------------
def run_soc_simulation(df: pd.DataFrame) -> pd.DataFrame:
    """
    Hour-by-hour SoC-constrained market optimiser.

    FORCED RECHARGE: when SoC < FCR_D_UP_MIN_SOC the battery buys energy
    at spot price to restore to FCR_N_MIN_SOC, then continues normally next
    hour. This prevents the infinite-IDLE bug from v1.

    SPOT DISCHARGE draws only from headroom above FCR_N_MIN_SOC, preserving
    the ability to continue FCR service the following hour.

    mFRR ACTIVATION MODEL:
    FCR is a capacity market — earned every hour within the SoC window.
    mFRR is an energy activation market — Fingrid decides when to dispatch.
    Revenue is scaled by MFRR_UP_ACT_PROB / MFRR_DOWN_ACT_PROB (the fraction
    of hours Fingrid actually calls the asset). Without this, the model would
    overstate mFRR revenue by 5–10× by assuming full dispatch every eligible
    hour. The probabilities (12% up, 6% down) are calibrated to Fingrid
    balancing statistics for the Finnish bidding zone.
    """
    records = []
    soc = BESS_MWH * 0.50    # Start at 50 % SoC

    for _, row in df.iterrows():
        def p(col):
            v = row.get(col, 0)
            return float(v) if pd.notna(v) else 0.0

        spot  = p("Price_Spot")
        fcrd  = p("Price_FCR_D_UP")
        fcrn  = p("Price_FCR_N")
        fcrdd = p("Price_FCR_D_DOWN")
        mup   = p("Price_mFRR_UP")
        mdn   = p("Price_mFRR_DOWN")

        # ── FORCED RECHARGE ──────────────────────────────────────────────────
        if soc < FCR_D_UP_MIN_SOC:
            mwh_needed    = FCR_N_MIN_SOC - soc
            mwh_from_grid = mwh_needed / CHG_EFF
            cost          = spot * mwh_from_grid
            new_soc       = min(MAX_SOC, soc + mwh_needed)
            records.append({
                "Market_Selected": "CHARGING",
                "Revenue_EUR":     -cost,
                "SoC_MWh":         new_soc,
            })
            soc = new_soc
            continue

        # ── BUILD CANDIDATES ─────────────────────────────────────────────────
        candidates = []

        if soc >= FCR_D_UP_MIN_SOC and fcrd > 0:
            avg_dch = BESS_MW * FCR_D_UP_ACT_FRAC * (FCR_D_UP_ACT_MIN / 60)
            candidates.append(("FCR_D_UP", fcrd * BESS_MW,
                                max(MIN_SOC, soc - avg_dch)))

        if FCR_N_MIN_SOC <= soc <= FCR_N_MAX_SOC and fcrn > 0:
            candidates.append(("FCR_N", fcrn * BESS_MW, soc))

        if soc <= FCR_D_DOWN_MAX_SOC and fcrdd > 0:
            avg_chg = BESS_MW * FCR_D_DOWN_ACT_FRAC * (FCR_D_DOWN_ACT_MIN / 60) * CHG_EFF
            candidates.append(("FCR_D_DOWN", fcrdd * BESS_MW,
                                min(MAX_SOC, soc + avg_chg)))

        if spot >= SPOT_DISCHARGE_MIN:
            headroom = soc - FCR_N_MIN_SOC
            if headroom > 0.5:
                mwh_out    = min(headroom * DCH_EFF, BESS_MW)
                soc_drop   = mwh_out / DCH_EFF
                recharge_c = SPOT_CHARGE_MAX * mwh_out / CHG_EFF
                net_rev    = spot * mwh_out - recharge_c
                if net_rev > 0:
                    candidates.append(("SPOT_DISC", net_rev,
                                       max(FCR_N_MIN_SOC, soc - soc_drop)))

        # mFRR UP: energy activation market — Fingrid decides when to dispatch.
        # Expected revenue = price × MWh × P(activation). SoC drifts only by the
        # probabilistic activation MWh, not a full dispatch every hour.
        if mup > 0 and mup >= spot * MFRR_MIN_PREMIUM:
            headroom = soc - FCR_N_MIN_SOC
            if headroom > 0.5:
                # Expected MWh dispatched this hour = full_capacity × probability
                exp_mwh  = BESS_MW * (MFRR_UP_ACT_MIN / 60) * MFRR_UP_ACT_PROB
                exp_mwh  = min(exp_mwh, headroom * DCH_EFF)
                soc_drop = exp_mwh / DCH_EFF
                candidates.append(("mFRR_UP", mup * exp_mwh,
                                   max(FCR_N_MIN_SOC, soc - soc_drop)))

        # mFRR DOWN: energy absorption market — Fingrid decides when to dispatch.
        if mdn > 0 and soc < MAX_SOC - 1.0:
            space     = MAX_SOC - soc
            exp_mwh   = BESS_MW * (MFRR_DOWN_ACT_MIN / 60) * MFRR_DOWN_ACT_PROB
            exp_mwh   = min(exp_mwh * CHG_EFF, space)
            new_soc   = min(MAX_SOC, soc + exp_mwh)
            candidates.append(("mFRR_DOWN", mdn * (exp_mwh / CHG_EFF), new_soc))

        if spot <= SPOT_CHARGE_MAX and soc < FCR_N_MAX_SOC:
            space   = MAX_SOC - soc
            mwh_in  = min(space / CHG_EFF, BESS_MW)
            new_soc = min(MAX_SOC, soc + mwh_in * CHG_EFF)
            candidates.append(("CHARGING", -(spot * mwh_in / CHG_EFF), new_soc))

        candidates.append(("IDLE", 0.0, soc))
        best = max(candidates, key=lambda x: x[1])

        records.append({
            "Market_Selected": best[0],
            "Revenue_EUR":     best[1],
            "SoC_MWh":         best[2],
        })
        soc = best[2]

    return pd.DataFrame(records, index=df.index)


# ---------------------------------------------------------------------------
# 6.  Helpers
# ---------------------------------------------------------------------------
def cleanup_outputs():
    for f in [EXCEL_FILE, CHART_FILE, DB_NAME_SQLITE]:
        if os.path.exists(f):
            os.remove(f)


# ---------------------------------------------------------------------------
# 7.  Main pipeline
# ---------------------------------------------------------------------------
def main():
    print(f"{'='*60}")
    print(f"  BESS REVENUE PIPELINE  —  DB: {DB_TYPE.upper()}")
    print(f"  Battery: {BESS_MW} MW / {BESS_MWH} MWh  (Finland)")
    print(f"{'='*60}\n")

    START = "2025-01-01T00:00:00Z"
    END   = "2025-12-31T23:59:59Z"

    # ------------------------------------------------------------------
    # PHASE 1: Data ingestion
    # ------------------------------------------------------------------
    print("── PHASE 1: Data Ingestion ──────────────────────────────")

    # Spot from ENTSO-E
    print("\n  Fetching Spot prices (ENTSO-E, FI bidding zone, €/MWh):")
    df_spot = fetch_spot_prices(START, END)
    if df_spot.empty:
        print("  [!] CRITICAL: No spot prices. See setup instructions above.")
        return
    df_spot["Timestamp"] = df_spot["Timestamp"].dt.tz_localize(None)
    verify_spot_prices(df_spot)

    # Reserve markets from Fingrid
    frames = {"Spot": df_spot}
    for label, meta in FINGRID_DATASETS.items():
        print(f"\n  Fetching {label} (dataset {meta['id']}, {meta['unit']}):")
        df_raw = fetch_fingrid_data(meta["id"], START, END, label)
        if not df_raw.empty:
            df_raw["Timestamp"] = df_raw["Timestamp"].dt.tz_localize(None)
            frames[label] = df_raw
        else:
            print(f"   [!] No data for {label}")

    # ------------------------------------------------------------------
    # PHASE 2: Merge
    # ------------------------------------------------------------------
    print("\n── PHASE 2: Merging datasets ────────────────────────────")
    df_merged = None
    for label, df_raw in frames.items():
        df_merged = df_raw if df_merged is None else \
                    pd.merge(df_merged, df_raw, on="Timestamp", how="outer")

    df_merged = df_merged.sort_values("Timestamp").reset_index(drop=True)
    print(f"  Merged shape: {df_merged.shape}")
    print(f"  Date range:   {df_merged['Timestamp'].min()} → "
          f"{df_merged['Timestamp'].max()}")
    nan_c = df_merged.isna().sum()
    nan_c = nan_c[nan_c > 0]
    if not nan_c.empty:
        print(f"  NaN counts:\n  " +
              nan_c.to_string().replace("\n", "\n  "))

    # ------------------------------------------------------------------
    # PHASE 3: Legacy revenue columns
    # ------------------------------------------------------------------
    print("\n── PHASE 3: Legacy Revenue Columns ─────────────────────")
    df_merged["Revenue_Spot_EUR"] = df_merged["Price_Spot"].apply(
        lambda p: p * BESS_MW if pd.notna(p) and p > 50 else 0
    )
    df_merged["Revenue_FCR_EUR"] = df_merged["Price_FCR_D_UP"].apply(
        lambda p: p * BESS_MW if pd.notna(p) else 0
    )
    print("  [OK] Legacy Revenue_Spot_EUR / Revenue_FCR_EUR computed")

    # ------------------------------------------------------------------
    # PHASE 4: SoC simulation
    # ------------------------------------------------------------------
    print("\n── PHASE 4: SoC Simulation ──────────────────────────────")
    df_soc = run_soc_simulation(df_merged)
    df_soc.index = df_merged.index
    df_full = pd.concat([df_merged, df_soc], axis=1)

    summary = df_full.groupby("Market_Selected")["Revenue_EUR"].agg(
        Hours="count",
        Total_Revenue_EUR="sum",
        Avg_Rev_per_Hour="mean",
    ).sort_values("Total_Revenue_EUR", ascending=False)
    print(f"\n  Market selection breakdown:")
    print(summary.to_string())
    total = df_full["Revenue_EUR"].sum()
    print(f"\n  Total SoC-optimised revenue: €{total:,.0f}")

    # ------------------------------------------------------------------
    # PHASE 5: Storage
    # ------------------------------------------------------------------
    print("\n── PHASE 5: Storage ─────────────────────────────────────")
    cleanup_outputs()

    all_price_cols = ["Timestamp", "Price_Spot"] + \
                     [f"Price_{k}" for k in FINGRID_DATASETS] + \
                     ["Revenue_Spot_EUR", "Revenue_FCR_EUR"]
    df_prices = df_full[[c for c in all_price_cols if c in df_full.columns]]

    sim_cols = ["Timestamp", "Price_Spot", "Price_FCR_N", "Price_FCR_D_UP",
                "Price_FCR_D_DOWN", "Price_mFRR_UP", "Price_mFRR_DOWN",
                "Market_Selected", "Revenue_EUR", "SoC_MWh"]
    df_sim = df_full[[c for c in sim_cols if c in df_full.columns]]

    # --- DB write: try configured backend, fall back to SQLite automatically ---
    def write_to_db(eng, label):
        df_prices.to_sql("bess_revenues", eng, if_exists="replace", index=False)
        df_sim.to_sql("bess_soc_model",   eng, if_exists="replace", index=False)
        print(f"  [OK] bess_revenues + bess_soc_model → {label}")

    db_written = False
    if DB_TYPE == "postgres":
        try:
            from sqlalchemy import text as sa_text
            pg_engine = get_db_engine()
            with pg_engine.connect() as conn:
                conn.execute(sa_text("SELECT 1"))   # connectivity check
            write_to_db(pg_engine, "postgres")
            db_written = True
        except Exception as e:
            short_err = str(e).split("\n")[0][:120]
            print(f"  [!] PostgreSQL unavailable: {short_err}")
            print("      Is Docker running?  →  docker-compose up -d")
            print("      Falling back to SQLite for this run...")

    if not db_written:
        sqlite_engine = create_engine(f"sqlite:///{DB_NAME_SQLITE}")
        write_to_db(sqlite_engine, "sqlite (fallback)")

    # --- Excel is always written regardless of DB status ---
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_prices.to_excel(writer, sheet_name="Raw Price Data", index=False)
        df_sim.to_excel(writer, sheet_name="SoC Simulation", index=False)
    print(f"  [OK] Excel: {EXCEL_FILE} (2 sheets)")

    # ------------------------------------------------------------------
    # PHASE 6: Chart
    # ------------------------------------------------------------------
    print("\n── PHASE 6: Visualisation ───────────────────────────────")
    fig, ax = plt.subplots(figsize=(14, 7))
    ax.plot(df_full["Timestamp"],
            df_full["Revenue_Spot_EUR"] / 1000,
            label="Spot Arbitrage — Naive (discharge if >€50/MWh, no SoC limit)",
            alpha=0.45, color="#4682B4", linewidth=0.8)
    ax.plot(df_full["Timestamp"],
            df_full["Revenue_FCR_EUR"] / 1000,
            label="FCR-D Reserve — Naive (full 30 MW offered every hour)",
            alpha=0.45, color="#D2691E", linewidth=0.8)
    ax.plot(df_full["Timestamp"],
            df_full["Revenue_EUR"].clip(lower=0) / 1000,
            label="SoC-Optimised Strategy (best feasible market each hour)",
            alpha=0.85, color="#2E8B57", linewidth=1.0)

    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f"€{x:,.0f} k"))
    ax.set_ylabel("Revenue per Hour", fontweight="bold", labelpad=12)
    ax.set_xlabel("Date", fontweight="bold")
    ax.set_title(
        f"BESS {BESS_MW} MW: Hourly Revenue Comparison — Full Year 2025\n"
        f"(Blue/Orange = naive benchmarks  |  Green = SoC-constrained optimum)",
        fontsize=13, fontweight="bold", pad=15)
    ax.legend(fontsize=8)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(CHART_FILE, dpi=300)
    plt.close()
    print(f"  [OK] Chart: {CHART_FILE}")

    print(f"\n{'='*60}")
    print("  Pipeline completed successfully.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()