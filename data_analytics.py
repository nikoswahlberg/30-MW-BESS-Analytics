"""
BESS Analytics & Investment Model — data_analytics.py
======================================================
Reads from the two DB tables written by data_retrieval_fingrid.py:
  • bess_revenues   — raw hourly prices + legacy naive strategies
  • bess_soc_model  — SoC-constrained market simulation results

Outputs:
  • 4 PNG charts
  • BESS_Analysis.xlsx with 5 sheets:
      1. Raw Price Data
      2. SoC Simulation
      3. Market Breakdown
      4. Investment Model  — 20-year cash flow, NPV, IRR, payback
      5. Scenarios         — Bull / Base / Bear

Financial assumptions (sources in comments — see write_investment_sheet).
"""

import os
import math
import numpy as np
import pandas as pd
import sqlalchemy
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# 1.  Configuration
# ---------------------------------------------------------------------------
load_dotenv()
DB_TYPE  = os.getenv("DB_TYPE", "sqlite").lower()

BESS_MW  = 30
BESS_MWH = 36
EXCEL_FILE = "BESS_Analysis.xlsx"

# ---------------------------------------------------------------------------
# 2.  Investment model parameters
# ---------------------------------------------------------------------------
INV = {
    "capex_eur":        12_000_000,   # €12 M total installed (€400/kW)
    "opex_fixed_eur":      180_000,   # Fixed O&M €/yr (1.5 % of CAPEX)
    "opex_var_pct":          0.005,   # Variable O&M (0.5 % of annual revenue)
    "replacement_year":         12,   # Cell replacement year
    "replacement_cost":  3_500_000,   # Cell replacement CAPEX (€)
    "project_life":             20,   # Years
    "wacc":                   0.08,   # Discount rate (8 %)
    "corp_tax":               0.20,   # Finnish corporate tax (20 %)
    "depr_years":               10,   # Straight-line depreciation period
    "revenue_growth":         0.03,   # Nominal revenue growth p.a. (3 %)
    "degradation":            0.02,   # Annual capacity degradation (2 %)
    "scenarios": {
        "Bear": {"capex_mult": 1.33, "rev_mult": 0.75, "wacc": 0.10},
        "Base": {"capex_mult": 1.00, "rev_mult": 1.00, "wacc": 0.08},
        "Bull": {"capex_mult": 0.80, "rev_mult": 1.25, "wacc": 0.07},
    },
}

# ---------------------------------------------------------------------------
# 3.  DB helpers
# ---------------------------------------------------------------------------
def get_db_engine(fallback_to_sqlite=True):
    """
    Returns a working DB engine. If DB_TYPE=postgres and the container is
    unreachable, automatically falls back to SQLite so downstream scripts
    never crash due to a stopped Docker container.
    """
    if DB_TYPE == "postgres":
        try:
            from sqlalchemy import text as sa_text
            url = (f"postgresql://{os.getenv('POSTGRES_USER')}:"
                   f"{os.getenv('POSTGRES_PASSWORD')}@"
                   f"{os.getenv('POSTGRES_HOST')}:"
                   f"{os.getenv('POSTGRES_PORT')}/"
                   f"{os.getenv('POSTGRES_DB')}")
            engine = sqlalchemy.create_engine(url)
            with engine.connect() as conn:
                conn.execute(sa_text("SELECT 1"))
            return engine
        except Exception as e:
            if fallback_to_sqlite:
                short = str(e).split("\n")[0][:100]
                print(f"  [!] PostgreSQL unavailable ({short})")
                print("      Falling back to SQLite — run: docker-compose up -d")
                return sqlalchemy.create_engine("sqlite:///bess_model.db")
            raise
    return sqlalchemy.create_engine("sqlite:///bess_model.db")


def load_tables(engine):
    df_prices, df_sim = pd.DataFrame(), pd.DataFrame()
    try:
        df_prices = pd.read_sql("SELECT * FROM bess_revenues", engine)
        df_prices["Timestamp"] = pd.to_datetime(df_prices["Timestamp"])
        df_prices = df_prices.sort_values("Timestamp").set_index("Timestamp")
    except Exception as e:
        print(f"   [!] Could not load bess_revenues: {e}")
    try:
        df_sim = pd.read_sql("SELECT * FROM bess_soc_model", engine)
        df_sim["Timestamp"] = pd.to_datetime(df_sim["Timestamp"])
        df_sim = df_sim.sort_values("Timestamp").set_index("Timestamp")
    except Exception as e:
        print(f"   [!] Could not load bess_soc_model: {e}")
    return df_prices, df_sim


# ---------------------------------------------------------------------------
# 4.  IRR / NPV  (pure Python, overflow-safe)
# ---------------------------------------------------------------------------
def npv(rate: float, cashflows: list) -> float:
    """Net present value. cashflows[0] is the t=0 investment (negative)."""
    # Guard against overflow when rate is extreme
    try:
        return sum(cf / (1.0 + rate) ** t for t, cf in enumerate(cashflows))
    except (OverflowError, ZeroDivisionError):
        return float("inf")


def npv_derivative(rate: float, cashflows: list) -> float:
    """Analytical derivative of NPV with respect to rate."""
    try:
        return sum(-t * cf / (1.0 + rate) ** (t + 1)
                   for t, cf in enumerate(cashflows))
    except (OverflowError, ZeroDivisionError):
        return 0.0


def xirr(cashflows: list,
          guess: float = 0.05,
          tol: float = 1e-7,
          max_iter: int = 500) -> float:
    """
    IRR via Newton-Raphson, overflow-safe.

    Fixes applied vs v1:
      • rate is clamped to [−0.9999, 10.0] after every step to prevent
        (1+rate)^N from overflowing when rate is very negative (which happens
        when cumulative FCF never turns positive, e.g. tiny revenue vs large
        CAPEX). In that pathological case the function returns NaN cleanly
        rather than raising OverflowError.
      • npv() and its derivative both catch OverflowError internally.
      • A sign-check guard: if there is no sign change in NPV over a wide
        rate sweep, there is no real IRR and NaN is returned immediately.
    """
    RATE_MIN, RATE_MAX = -0.9999, 10.0

    # Quick sign-change check over a coarse grid
    test_rates = np.linspace(-0.9, 2.0, 30)
    signs = [math.copysign(1, npv(r, cashflows)) for r in test_rates
             if not math.isnan(npv(r, cashflows))]
    if len(set(signs)) < 2:
        # NPV never changes sign → no real IRR exists
        return float("nan")

    rate = max(RATE_MIN, min(guess, RATE_MAX))
    for _ in range(max_iter):
        f   = npv(rate, cashflows)
        df_ = npv_derivative(rate, cashflows)
        if df_ == 0:
            return float("nan")
        rate_new = rate - f / df_
        rate_new = max(RATE_MIN, min(rate_new, RATE_MAX))
        if abs(rate_new - rate) < tol:
            return rate_new
        rate = rate_new
    return float("nan")


def payback_years(cashflows: list) -> float:
    """Undiscounted payback period in years (excludes t=0 investment)."""
    cumulative = 0.0
    for t, cf in enumerate(cashflows):
        prev = cumulative
        cumulative += cf
        if t > 0 and cumulative >= 0:
            frac = abs(prev) / cf if cf != 0 else 0
            return (t - 1) + frac
    return float("inf")


# ---------------------------------------------------------------------------
# 5.  20-year cash flow model
# ---------------------------------------------------------------------------
def build_cashflows(annual_revenue_y1: float, params: dict) -> pd.DataFrame:
    rows = []
    capex       = params["capex_eur"]
    depr_annual = capex / params["depr_years"]
    repl_cost   = params["replacement_cost"]
    repl_yr     = params["replacement_year"]

    for yr in range(1, params["project_life"] + 1):
        cap_factor  = (1 - params["degradation"]) ** (yr - 1)
        grow_factor = (1 + params["revenue_growth"]) ** (yr - 1)
        gross_rev   = annual_revenue_y1 * cap_factor * grow_factor

        opex   = params["opex_fixed_eur"] + params["opex_var_pct"] * gross_rev
        ebitda = gross_rev - opex

        # Depreciation: straight-line on initial CAPEX; replacement amortised
        # over remaining project life after replacement year
        depr = depr_annual if yr <= params["depr_years"] else 0.0
        if yr > repl_yr:
            remaining = params["project_life"] - repl_yr
            depr += repl_cost / remaining

        ebit  = ebitda - depr
        tax   = max(0.0, ebit * params["corp_tax"])
        nopat = ebit - tax
        repl  = repl_cost if yr == repl_yr else 0.0
        fcf   = nopat + depr - repl   # add back non-cash D&A, subtract capex

        rows.append({
            "Year": yr, "Gross_Revenue": gross_rev, "OPEX": opex,
            "EBITDA": ebitda, "Depreciation": depr, "EBIT": ebit,
            "Tax": tax, "NOPAT": nopat, "Replacement_CAPEX": repl,
            "Free_Cashflow": fcf,
        })

    df = pd.DataFrame(rows)
    df["Cumulative_FCF"] = df["Free_Cashflow"].cumsum() - capex
    return df


# ---------------------------------------------------------------------------
# 6.  Chart formatters
# ---------------------------------------------------------------------------
def fmt_millions(x, _):   return f"€{x:.0f} M"
def fmt_thousands(x, _):  return f"€{x * 1000:.0f} k"
def fmt_monthly(x, _):    return f"€{x:.2f} M"


# ---------------------------------------------------------------------------
# 7.  Charts
# ---------------------------------------------------------------------------
def chart_monthly(df_prices, df_sim):
    print("   Creating: Monthly_Revenue_2025.png")
    df_m = df_prices[["Revenue_Spot_EUR", "Revenue_FCR_EUR"]].copy() / 1_000_000
    has_sim = not df_sim.empty and "Revenue_EUR" in df_sim.columns
    if has_sim:
        df_m["Revenue_SoC_EUR"] = df_sim["Revenue_EUR"].clip(lower=0) / 1_000_000

    monthly = df_m.resample("ME").sum()
    monthly.index = monthly.index.strftime("%B")
    colours = ["#4682B4", "#D2691E"] + (["#2E8B57"] if has_sim else [])
    labels  = ["Spot Arbitrage (Naive)", "FCR-D Reserve (Naive)"] + \
              (["SoC-Optimised Strategy"] if has_sim else [])

    fig, ax = plt.subplots(figsize=(13, 6))
    monthly.plot(kind="bar", color=colours, width=0.75, ax=ax)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_monthly))
    ax.set_ylabel("Total Monthly Revenue\n(€ Millions)",
                  fontweight="bold", labelpad=12)
    ax.axhline(0, color="black", linewidth=0.8, linestyle="--")
    ax.set_title("BESS 30 MW: Total Monthly Revenue 2025",
                 fontsize=15, fontweight="bold", pad=20)
    ax.set_xlabel("Month", fontweight="bold")
    ax.legend(labels)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("Monthly_Revenue_2025.png", dpi=300)
    plt.close()


def chart_cumulative(df_prices, df_sim):
    print("   Creating: Cumulative_Revenue_2025.png")
    df_m  = df_prices[["Revenue_Spot_EUR", "Revenue_FCR_EUR"]].copy() / 1_000_000
    cumul = df_m.cumsum()

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(cumul.index, cumul["Revenue_Spot_EUR"],
            color="#4682B4", linewidth=2.0, label="Spot Arbitrage (Naive)")
    ax.plot(cumul.index, cumul["Revenue_FCR_EUR"],
            color="#D2691E", linewidth=2.0, label="FCR-D Reserve (Naive)")

    if not df_sim.empty and "Revenue_EUR" in df_sim.columns:
        soc_cum = (df_sim["Revenue_EUR"].clip(lower=0) / 1_000_000).cumsum()
        ax.plot(soc_cum.index, soc_cum, color="#2E8B57", linewidth=2.5,
                label="SoC-Optimised Strategy")
        final = soc_cum.iloc[-1]
        ax.annotate(f"SoC-Opt year-end:\n€{final:.0f} M",
                    xy=(soc_cum.index[-1], final),
                    xytext=(-95, 10), textcoords="offset points",
                    fontsize=9, color="#2E8B57",
                    arrowprops=dict(arrowstyle="->", color="#2E8B57", lw=1.2))

    for col, color, lbl in [("Revenue_Spot_EUR", "#4682B4", "Spot"),
                             ("Revenue_FCR_EUR",  "#D2691E", "FCR-D")]:
        v  = cumul[col].iloc[-1]
        ax.annotate(f"{lbl} year-end:\n€{v:.0f} M",
                    xy=(cumul.index[-1], v),
                    xytext=(-90, -22), textcoords="offset points",
                    fontsize=9, color=color,
                    arrowprops=dict(arrowstyle="->", color=color, lw=1.2))

    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_millions))
    ax.set_ylabel("Accumulated Revenue Since Jan 1, 2025",
                  fontweight="bold", labelpad=12)
    ax.set_title("BESS 30 MW: 2025 Cumulative Profitability",
                 fontsize=15, fontweight="bold", pad=20)
    ax.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig("Cumulative_Revenue_2025.png", dpi=300)
    plt.close()


def chart_trend(df_prices, df_sim):
    print("   Creating: BESS_Strategy_Trend_2025.png")
    df_m = df_prices[["Revenue_Spot_EUR", "Revenue_FCR_EUR"]].copy() / 1_000_000
    # min_periods=1 ensures curve starts Jan 1, not Jan 8
    spot_smooth = df_m["Revenue_Spot_EUR"].rolling(
        window=24 * 7, min_periods=1).mean()

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.plot(df_m.index, df_m["Revenue_FCR_EUR"],
            color="#D2691E", alpha=0.4,
            label="FCR-D Reserve — Stable Capacity Fee (Raw Hourly)")
    ax.plot(spot_smooth.index, spot_smooth,
            color="#4682B4", linewidth=2,
            label="Spot Arbitrage — Price-Driven Trading (7-Day Rolling Average)")

    if not df_sim.empty and "Revenue_EUR" in df_sim.columns:
        soc_smooth = (df_sim["Revenue_EUR"].clip(lower=0) / 1_000_000
                      ).rolling(window=24 * 7, min_periods=1).mean()
        ax.plot(soc_smooth.index, soc_smooth,
                color="#2E8B57", linewidth=2, linestyle="--",
                label="SoC-Optimised Strategy (7-Day Rolling Average)")

    ax.yaxis.set_major_formatter(mticker.FuncFormatter(fmt_thousands))
    ax.set_ylabel("Revenue per Hour", fontweight="bold", labelpad=12)
    ax.set_title("Performance Trend: Trading Volatility vs. Capacity Stability",
                 fontsize=15, fontweight="bold", pad=20)
    ax.legend()
    plt.tight_layout()
    plt.savefig("BESS_Strategy_Trend_2025.png", dpi=300)
    plt.close()


def chart_soc(df_sim):
    if df_sim.empty or "SoC_MWh" not in df_sim.columns:
        print("   [SKIP] SoC profile — no bess_soc_model data")
        return
    print("   Creating: BESS_SoC_Profile_2025.png")

    market_colours = {
        "FCR_D_UP":   "#1a6faf", "FCR_N":      "#7cb8dd",
        "FCR_D_DOWN": "#d96a0b", "SPOT_DISC":  "#b02020",
        "mFRR_UP":    "#8B0000", "mFRR_DOWN":  "#006400",
        "CHARGING":   "#556B2F", "IDLE":       "#cccccc",
    }

    fig, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=True,
                             gridspec_kw={"height_ratios": [3, 1]})
    ax = axes[0]
    ax.plot(df_sim.index, df_sim["SoC_MWh"],
            color="#2E8B57", linewidth=0.7, alpha=0.85)
    ax.axhline(BESS_MWH * 0.90, color="gray", linewidth=0.8, linestyle=":",
               label=f"Max SoC ({BESS_MWH*0.90:.1f} MWh / 90 %)")
    ax.axhline(BESS_MWH * 0.10, color="gray", linewidth=0.8, linestyle="--",
               label=f"Min SoC ({BESS_MWH*0.10:.1f} MWh / 10 %)")
    ax.fill_between(df_sim.index, df_sim["SoC_MWh"], alpha=0.15, color="#2E8B57")
    ax.set_ylabel("State of Charge (MWh)", fontweight="bold")
    ax.set_ylim(0, BESS_MWH * 1.05)
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f"{x:.0f} MWh"))
    ax.set_title("BESS 30 MW / 36 MWh: State of Charge & Market Selection — 2025",
                 fontsize=14, fontweight="bold", pad=15)
    ax.legend(loc="upper right", fontsize=8)

    ax2 = axes[1]
    if "Market_Selected" in df_sim.columns:
        daily_mkt = df_sim["Market_Selected"].resample("D").agg(
            lambda x: x.value_counts().idxmax())
        for dt, mkt in daily_mkt.items():
            ax2.axvspan(dt, dt + pd.Timedelta(days=1),
                        color=market_colours.get(mkt, "#cccccc"), alpha=0.85)
        patches = [mpatches.Patch(color=c, label=m)
                   for m, c in market_colours.items()
                   if m in df_sim["Market_Selected"].unique()]
        ax2.legend(handles=patches, loc="center left",
                   bbox_to_anchor=(1.0, 0.5), fontsize=7, title="Market")
    ax2.set_ylabel("Dominant\nMarket", fontweight="bold", fontsize=8)
    ax2.set_yticks([])
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
    ax2.xaxis.set_major_locator(mdates.MonthLocator())
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("BESS_SoC_Profile_2025.png", dpi=300, bbox_inches="tight")
    plt.close()


# ---------------------------------------------------------------------------
# 8.  Excel helpers
# ---------------------------------------------------------------------------
def style_header_row(ws, row: int, n_cols: int,
                     bg: str = "1F4E79", fg: str = "FFFFFF"):
    fill = PatternFill("solid", fgColor=bg)
    font = Font(bold=True, color=fg, size=10)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)


def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def autofit_columns(ws, min_w: int = 10, max_w: int = 32):
    for col_cells in ws.columns:
        length = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[
            get_column_letter(col_cells[0].column)
        ].width = max(min_w, min(length + 2, max_w))


# ---------------------------------------------------------------------------
# 9.  Investment model sheet
# ---------------------------------------------------------------------------
def write_investment_sheet(wb, annual_revenue_y1: float):
    ws = wb.create_sheet("Investment Model")
    ws.sheet_view.showGridLines = False

    p        = INV.copy()
    cf_table = build_cashflows(annual_revenue_y1, p)

    cfs         = [-p["capex_eur"]] + cf_table["Free_Cashflow"].tolist()
    project_irr = xirr(cfs)
    project_npv = npv(p["wacc"], cfs)
    payback     = payback_years(cfs)

    # ── Title ──
    ws["A1"] = "PAISTINKULMA BESS — INVESTMENT MODEL (30 MW / 36 MWh)"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.merge_cells("A1:F1")
    ws["A2"] = (
        "Sources: CAPEX — Ember 'How cheap is battery storage' Jan 2026 (~$125/kWh all-in, adj. to €150/kWh for European 2024 delivery)  |  "
        "Revenue — Fingrid open data SoC simulation  |  "
        "Rev growth — Fingrid reserve market demand +134% forecast (5 yr)  |  "
        "IRR benchmark — 3–7% unlevered Western EU merchant BESS (Capstone DC, Nov 2025)")
    ws["A2"].font = Font(italic=True, size=8, color="595959")
    ws.merge_cells("A2:J2")

    # ── KPI block ──
    GREEN, YELLOW, RED = "375623", "7F6000", "842019"

    # Format helper: return safe string for NaN IRR
    irr_str  = f"{project_irr:.2%}" if not math.isnan(project_irr) else "N/A"
    pay_str  = f"{payback:.1f} yrs" if project_irr != float("inf") else "N/A"

    # Operational discount: model assumes perfect foresight and 100% availability.
    # Real dispatch is typically 10–25% lower due to bidding friction, minimum
    # volume requirements, and telemetry / maintenance downtime.
    OPS_DISCOUNT = 0.20   # conservative 20% haircut
    rev_ops       = annual_revenue_y1 * (1 - OPS_DISCOUNT)
    cfs_ops       = [-p["capex_eur"]] + build_cashflows(rev_ops, p)["Free_Cashflow"].tolist()
    irr_ops       = xirr(cfs_ops)
    npv_ops       = npv(p["wacc"], cfs_ops)
    pay_ops       = payback_years(cfs_ops)

    kpis = [
        ("Metric",              "Value",               "Notes"),
        ("CAPEX",               p["capex_eur"],        "Total installed cost (€)"),
        ("── Model (theoretical maximum) ──", "", "Perfect foresight, 100% availability"),
        ("Annual Revenue Y1",   annual_revenue_y1,     "From SoC simulation (€)"),
        ("Revenue / MW / month",annual_revenue_y1 / BESS_MW / 12,
         "Benchmark: Exilion FI achieved €40,700/MW/month in H2 2023"),
        ("Project IRR",         project_irr if not math.isnan(project_irr) else "N/A",
         "Unlevered; Newton-Raphson. Upper bound — see operational estimate below."),
        ("Project NPV",         project_npv,           "At 8% WACC (€) — upper bound"),
        ("Simple Payback",      payback if payback != float("inf") else "N/A",
         "Undiscounted (years) — lower bound"),
        ("── Operational estimate (−20% haircut) ──", "", "Accounts for bidding friction, availability ~92%, min bid volumes"),
        ("Adj. Revenue Y1",     rev_ops,               "Model revenue × 0.80 (€)"),
        ("Adj. Revenue/MW/mo",  rev_ops / BESS_MW / 12,"Adjusted benchmark"),
        ("Adj. Project IRR",    irr_ops if not math.isnan(irr_ops) else "N/A",
         "More realistic operational IRR estimate"),
        ("Adj. Project NPV",    npv_ops,               "At 8% WACC (€)"),
        ("Adj. Simple Payback", pay_ops if pay_ops != float("inf") else "N/A",
         "Undiscounted (years)"),
        ("── Finance ──", "", ""),
        ("WACC",                p["wacc"],             "Infrastructure equity fund discount rate"),
        ("Cell Replacement Yr", p["replacement_year"], "Partial at year 12; BOS retained"),
        ("Corp Tax Rate",       p["corp_tax"],         "Finland 20%"),
        ("Annual Degradation",  p["degradation"],      "LFP chemistry, real-world data"),
    ]

    for r_idx, (label, value, note) in enumerate(kpis, start=4):
        label_cell = ws.cell(row=r_idx, column=1, value=label)
        # Section divider rows (value == "") — style as a subheader, skip formatting
        if value == "":
            label_cell.font = Font(bold=True, italic=True, color="1F4E79")
            ws.cell(row=r_idx, column=3, value=note
                    ).font = Font(italic=True, color="595959", size=8)
            ws.merge_cells(start_row=r_idx, start_column=1,
                           end_row=r_idx, end_column=2)
            continue
        label_cell.font = Font(bold=True)
        cell = ws.cell(row=r_idx, column=2, value=value)
        ws.cell(row=r_idx, column=3, value=note
                ).font = Font(italic=True, color="595959")
        if isinstance(value, str):
            # "N/A" values coloured red
            cell.font = Font(bold=True, color=RED)
        elif "CAPEX" in label or "Revenue" in label or "NPV" in label:
            cell.number_format = '#,##0 "€"'
        elif "IRR" in label or "WACC" in label or "Tax" in label or "Degrad" in label:
            if isinstance(value, float) and not math.isnan(value):
                cell.number_format = "0.00%"
            if "IRR" in label and isinstance(value, float) and not math.isnan(value):
                colour = GREEN if value > 0.10 else (YELLOW if value > 0.06 else RED)
                cell.font = Font(bold=True, color=colour)
        elif "Payback" in label and isinstance(value, float):
            cell.number_format = '0.0 "yrs"'

    # ── Cash flow table ──
    tbl_start = 4 + len(kpis) + 1
    style_header_row(ws, tbl_start, 11)
    headers = ["Year", "Gross Revenue (€)", "OPEX (€)", "EBITDA (€)",
               "Depreciation (€)", "EBIT (€)", "Tax (€)", "NOPAT (€)",
               "Replacement CAPEX (€)", "Free Cash Flow (€)",
               "Cumulative FCF (€)"]
    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=tbl_start, column=c_idx, value=h)

    ALT_FILL  = PatternFill("solid", fgColor="D9E2F3")
    REPL_FILL = PatternFill("solid", fgColor="FCE4D6")

    # Year 0 — initial investment
    yr0_row = tbl_start + 1
    ws.cell(row=yr0_row, column=1, value=0).font = Font(bold=True)
    ws.cell(row=yr0_row, column=10, value=-p["capex_eur"]
            ).number_format = '[Red]-#,##0 "€"'
    ws.cell(row=yr0_row, column=11, value=-p["capex_eur"]
            ).number_format = '[Red]-#,##0 "€"'

    for r_idx, row in cf_table.reset_index(drop=True).iterrows():
        excel_row = tbl_start + 2 + r_idx
        data = [row["Year"], row["Gross_Revenue"], row["OPEX"], row["EBITDA"],
                row["Depreciation"], row["EBIT"], row["Tax"], row["NOPAT"],
                row["Replacement_CAPEX"], row["Free_Cashflow"],
                row["Cumulative_FCF"]]
        for c_idx, val in enumerate(data, start=1):
            cell = ws.cell(row=excel_row, column=c_idx, value=val)
            cell.border = thin_border()
            if c_idx == 1:
                cell.number_format = "0"
            elif c_idx == 10:  # FCF — green/red conditional
                cell.number_format = '[Green]#,##0 "€";[Red]-#,##0 "€"'
            elif c_idx == 11:  # Cumulative FCF
                cell.number_format = '[Green]#,##0 "€";[Red]-#,##0 "€"'
            else:
                cell.number_format = '#,##0 "€"'
            if int(row["Year"]) % 2 == 0:
                cell.fill = ALT_FILL
            if int(row["Year"]) == p["replacement_year"]:
                cell.fill = REPL_FILL

    ws.freeze_panes = f"B{tbl_start + 1}"
    autofit_columns(ws)
    return project_irr, project_npv, payback


# ---------------------------------------------------------------------------
# 10.  Scenarios sheet
# ---------------------------------------------------------------------------
def write_scenarios_sheet(wb, annual_revenue_y1: float):
    ws = wb.create_sheet("Scenarios")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SCENARIO ANALYSIS — 30 MW BESS PAISTINKULMA"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.merge_cells("A1:G1")

    scen_names    = list(INV["scenarios"].keys())
    SCEN_COLOURS  = {"Bear": "C00000", "Base": "1F4E79", "Bull": "375623"}
    GREEN, YELLOW, RED = "375623", "7F6000", "842019"

    # Header row
    style_header_row(ws, 3, len(scen_names) + 1)
    ws.cell(row=3, column=1, value="KPI / Assumption")
    for c, name in enumerate(scen_names, start=2):
        cell = ws.cell(row=3, column=c, value=name)
        cell.font = Font(bold=True, color=SCEN_COLOURS[name])
        cell.alignment = Alignment(horizontal="center")

    rows_meta = [
        ("CAPEX (€)",          lambda sc: INV["capex_eur"] * sc["capex_mult"],    '#,##0 "€"',   None),
        ("Revenue Y1 (€)",     lambda sc: annual_revenue_y1 * sc["rev_mult"],     '#,##0 "€"',   None),
        ("WACC",               lambda sc: sc["wacc"],                             "0.0%",        None),
        ("Revenue multiplier", lambda sc: sc["rev_mult"],                         '0.00"×"',     None),
        ("─── Results ───",    None, None, None),
        ("Project IRR",        None, "0.00%",  "irr"),
        ("Project NPV (€)",    None, '#,##0 "€"', "npv"),
        ("Payback (yrs)",      None, '0.0 "yrs"', "pay"),
    ]

    for r_off, (label, fn, fmt, special) in enumerate(rows_meta, start=4):
        ws.cell(row=r_off, column=1, value=label).font = Font(bold=True)

        for c_idx, name in enumerate(scen_names, start=2):
            sc      = INV["scenarios"][name]
            s_capex = INV["capex_eur"] * sc["capex_mult"]
            s_rev   = annual_revenue_y1 * sc["rev_mult"]
            s_wacc  = sc["wacc"]

            if fn is not None:
                val = fn(sc)
                cell = ws.cell(row=r_off, column=c_idx, value=val)
                if fmt:
                    cell.number_format = fmt
            elif special:
                p_sc    = {**INV, "capex_eur": s_capex, "wacc": s_wacc}
                cf_sc   = build_cashflows(s_rev, p_sc)
                cfs_sc  = [-s_capex] + cf_sc["Free_Cashflow"].tolist()
                s_irr   = xirr(cfs_sc)
                s_npv   = npv(s_wacc, cfs_sc)
                s_pay   = payback_years(cfs_sc)

                if special == "irr":
                    val = s_irr if not math.isnan(s_irr) else "N/A"
                    cell = ws.cell(row=r_off, column=c_idx, value=val)
                    if isinstance(val, float):
                        cell.number_format = "0.00%"
                        colour = GREEN if val > 0.10 else (YELLOW if val > 0.06 else RED)
                        cell.font = Font(bold=True, color=colour)
                    else:
                        cell.font = Font(bold=True, color=RED)
                elif special == "npv":
                    cell = ws.cell(row=r_off, column=c_idx, value=s_npv)
                    cell.number_format = '[Blue]#,##0 "€";[Red]-#,##0 "€"'
                elif special == "pay":
                    val = s_pay if s_pay != float("inf") else "N/A"
                    cell = ws.cell(row=r_off, column=c_idx, value=val)
                    if isinstance(val, float):
                        cell.number_format = '0.0 "yrs"'

    # Footnotes
    note_start = 4 + len(rows_meta) + 3
    notes = [
        "Key assumptions & sources:",
        f"  CAPEX base: €{INV['capex_eur']/1e6:.0f} M (€400/kW; Ember Jan 2026: ~$125/kWh all-in 2025, adj. for European Q4 2024 procurement)",
        f"  Degradation: {INV['degradation']*100:.0f}% p.a. capacity loss (LFP chemistry)",
        f"  Revenue growth: {INV['revenue_growth']*100:.0f}% nominal p.a. (Fingrid reserve demand forecast +134% over 5 yrs)",
        f"  Cell replacement: €{INV['replacement_cost']/1e6:.1f} M at year {INV['replacement_year']} (cells only; inverters and BOS retained)",
        f"  Finnish corp tax {INV['corp_tax']*100:.0f}% | Straight-line depreciation {INV['depr_years']} yrs",
        f"  IRR benchmark: 3–7% unlevered Western EU merchant BESS (Capstone DC, Nov 2025); Finland structurally higher due to price volatility",
        f"  Exilion FI benchmark: €40,700/MW/month in H2 2023 (Capalo AI, Oct 2024)",
        f"  Capacity doubling to 60 MW flagged as future option by Taaleri Energia — doubling MW at same site halves €/kW BOS cost",
        "  ── Model limitations (revenue upper bound) ──",
        "  mFRR UP/DOWN are energy activation markets: Fingrid decides when to dispatch, not the operator.",
        "  Revenue is scaled by activation probability (12% UP / 6% DOWN) calibrated to Fingrid balancing statistics.",
        "  Without this scaling, mFRR revenue would be overstated 5–10×. FCR-N and FCR-D are capacity markets",
        "  and are correctly modelled as earned every hour within the feasible SoC window.",
        "  24 NaN spot price hours (DST clock-change gaps) are treated as €0 — negligible impact on annual totals.",
        "  Revenue Y1 represents a theoretical optimum under perfect foresight; real dispatch will be ~10–20% lower",
        "  due to sub-optimal bidding, minimum bid volumes, and operational constraints not modelled here.",
    ]
    for i, note in enumerate(notes):
        cell = ws.cell(row=note_start + i, column=1, value=note)
        cell.font = Font(italic=True, size=9, color="595959")
        ws.merge_cells(start_row=note_start + i, start_column=1,
                       end_row=note_start + i, end_column=7)

    autofit_columns(ws)


# ---------------------------------------------------------------------------
# 11.  Market breakdown sheet
# ---------------------------------------------------------------------------
def write_market_breakdown_sheet(wb, df_sim):
    ws = wb.create_sheet("Market Breakdown")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "MARKET SELECTION — SoC-OPTIMISED STRATEGY"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.merge_cells("A1:F1")

    if df_sim.empty or "Market_Selected" not in df_sim.columns:
        ws["A3"] = "No SoC simulation data available (run data_retrieval_fingrid.py first)."
        return

    grp = (df_sim.groupby("Market_Selected")
           .agg(Hours=("Revenue_EUR", "count"),
                Total_Revenue_EUR=("Revenue_EUR", "sum"),
                Avg_Hourly_EUR=("Revenue_EUR", "mean"),
                Min_Revenue_EUR=("Revenue_EUR", "min"),
                Max_Revenue_EUR=("Revenue_EUR", "max"))
           .reset_index()
           .sort_values("Total_Revenue_EUR", ascending=False))
    grp["Share_Hours"]   = grp["Hours"] / grp["Hours"].sum()
    grp["Share_Revenue"] = grp["Total_Revenue_EUR"] / grp["Total_Revenue_EUR"].sum()

    style_header_row(ws, 3, len(grp.columns))
    for c_idx, col in enumerate(grp.columns, start=1):
        ws.cell(row=3, column=c_idx, value=col)

    ALT_FILL = PatternFill("solid", fgColor="D9E2F3")
    for r_idx, row in grp.reset_index(drop=True).iterrows():
        excel_row = r_idx + 4
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=excel_row, column=c_idx, value=val)
            col_name = grp.columns[c_idx - 1]
            if "EUR" in col_name:
                cell.number_format = '#,##0 "€"'
            elif "Share" in col_name:
                cell.number_format = "0.0%"
            cell.border = thin_border()
            if r_idx % 2 == 1:
                cell.fill = ALT_FILL

    # Monthly breakdown by market
    mth_hdr = 3 + len(grp) + 3
    ws.cell(row=mth_hdr - 1, column=1,
            value="Monthly Revenue by Market (€)").font = Font(bold=True,
                                                                color="1F4E79")
    monthly_mkt = (df_sim.groupby(
        [df_sim.index.to_period("M"), "Market_Selected"])["Revenue_EUR"]
        .sum().unstack(fill_value=0))
    monthly_mkt.index = monthly_mkt.index.strftime("%B %Y")

    style_header_row(ws, mth_hdr, len(monthly_mkt.columns) + 1)
    ws.cell(row=mth_hdr, column=1, value="Month")
    for c_idx, col in enumerate(monthly_mkt.columns, start=2):
        ws.cell(row=mth_hdr, column=c_idx, value=col)

    ALT2 = PatternFill("solid", fgColor="E2EFDA")
    for r_idx, (month, row) in enumerate(monthly_mkt.iterrows()):
        er = mth_hdr + 1 + r_idx
        ws.cell(row=er, column=1, value=month)
        for c_idx, val in enumerate(row, start=2):
            cell = ws.cell(row=er, column=c_idx, value=val)
            cell.number_format = '#,##0 "€"'
            if r_idx % 2 == 1:
                cell.fill = ALT2

    autofit_columns(ws)


# ---------------------------------------------------------------------------
# 12.  Main
# ---------------------------------------------------------------------------
def run_analytics():
    print("=" * 60)
    print("  BESS ANALYTICS & INVESTMENT MODEL")
    print("=" * 60)

    engine = get_db_engine()
    df_prices, df_sim = load_tables(engine)

    if df_prices.empty:
        print("[!] No price data found — run data_retrieval_fingrid.py first.")
        return

    # Revenue Y1: prefer SoC simulation; fall back to FCR-D naive
    if not df_sim.empty and "Revenue_EUR" in df_sim.columns:
        revenue_y1 = float(df_sim["Revenue_EUR"].clip(lower=0).sum())
        src = "SoC-optimised"
    else:
        revenue_y1 = float(df_prices["Revenue_FCR_EUR"].sum())
        src = "FCR-D naive (fallback)"
    print(f"\n  Revenue Y1 ({src}): €{revenue_y1:,.0f}")

    if revenue_y1 < 100_000:
        print("  ⚠  Revenue is very low — investment metrics will show "
              "negative IRR / infinite payback.")
        print("     This likely means the SoC simulation has a data issue "
              "(check spot price units in Phase 1 output).")

    sns.set_theme(style="whitegrid")

    # ── Charts ──────────────────────────────────────────────────────────────
    print("\n── Charts ───────────────────────────────────────────────")
    chart_monthly(df_prices, df_sim)
    chart_cumulative(df_prices, df_sim)
    chart_trend(df_prices, df_sim)
    chart_soc(df_sim)

    # ── Excel workbook ───────────────────────────────────────────────────────
    print("\n── Excel Workbook ───────────────────────────────────────")
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_prices.reset_index().to_excel(
            writer, sheet_name="Raw Price Data", index=False)
        if not df_sim.empty:
            df_sim.reset_index().to_excel(
                writer, sheet_name="SoC Simulation", index=False)

    wb = load_workbook(EXCEL_FILE)
    write_market_breakdown_sheet(wb, df_sim)

    print("   Building Investment Model sheet...")
    base_irr, base_npv, base_payback = write_investment_sheet(wb, revenue_y1)
    irr_display = f"{base_irr:.2%}" if not math.isnan(base_irr) else "N/A"
    npv_display = f"€{base_npv:,.0f}"
    pay_display = f"{base_payback:.1f} yrs" if base_payback != float("inf") else "N/A"
    print(f"   → IRR: {irr_display}  |  NPV: {npv_display}  |  Payback: {pay_display}")

    print("   Building Scenarios sheet...")
    write_scenarios_sheet(wb, revenue_y1)

    # Reorder sheets
    order = ["Raw Price Data", "SoC Simulation", "Market Breakdown",
             "Investment Model", "Scenarios"]
    wb._sheets.sort(key=lambda s: order.index(s.title)
                    if s.title in order else 99)
    wb.save(EXCEL_FILE)
    print(f"   [OK] {EXCEL_FILE} saved (5 sheets)")

    print(f"\n{'='*60}")
    print("  All outputs generated successfully.")
    print(f"{'='*60}")


if __name__ == "__main__":
    run_analytics()