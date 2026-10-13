"""Task 1: Automotive Market Overview & Trend Analysis (India).

Downloads five years of public daily market data, computes return/risk metrics,
makes charts, and creates Word report + submission description. Synthetic price
data is never substituted for missing market data. Research context is dated and
its public-source URLs are listed in the generated report.
"""
from datetime import datetime, timezone
from pathlib import Path
import json
import warnings

import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs"
CHARTS = ROOT / "visualizations"
REPORT_PATH = ROOT / "Automotive_Market_Overview_and_Trend_Analysis_Report.docx"
SUBMISSION_DOCX = ROOT / "Submission_Description.docx"
SUBMISSION_TXT = ROOT / "submission_description.txt"
START_PERIOD = "5y"
TRADING_DAYS = 252

TICKERS = {
    "^CNXAUTO": ("Nifty Auto", "Sector index"),
    "^NSEI": ("Nifty 50", "Broad-market index"),
    "MARUTI.NS": ("Maruti Suzuki", "Passenger vehicles"),
    "M&M.NS": ("Mahindra & Mahindra", "SUVs, LCVs and tractors"),
    "BAJAJ-AUTO.NS": ("Bajaj Auto", "Two/three-wheelers and exports"),
    "EICHERMOT.NS": ("Eicher Motors", "Motorcycles and commercial vehicles"),
    "HEROMOTOCO.NS": ("Hero MotoCorp", "Two-wheelers"),
    "TVSMOTOR.NS": ("TVS Motor", "Two/three-wheelers"),
    "ASHOKLEY.NS": ("Ashok Leyland", "Commercial vehicles"),
    "BHARATFORG.NS": ("Bharat Forge", "Auto components and industrial forging"),
    "BOSCHLTD.NS": ("Bosch", "Automotive technology and components"),
    "EXIDEIND.NS": ("Exide Industries", "Batteries"),
}

SOURCES = [
    ("S1", "Society of Indian Automobile Manufacturers (SIAM), Auto Industry Performance of August 2026, 15 September 2026",
     "https://www.siam.in/pressrelease-details.aspx?mpgid=53&pgidtrail=50&pid=620",
     "SIAM reports August 2026 production of 3,157,822 units across covered categories; domestic sales included 439,309 passenger vehicles, 93,764 three-wheelers and 2,034,698 two-wheelers. It reports year-on-year growth of 36.5%, 22.8% and 10.5%, respectively, and cautions that the prior-year base was lower."),
    ("S2", "Press Information Bureau / Ministry of Heavy Industries, GST rates for automobiles, 4 September 2025",
     "https://www.pib.gov.in/PressReleasePage.aspx?PRID=2163763&lang=2&reg=48",
     "The government describes GST reductions across vehicle categories and auto parts, aimed at supporting affordability and downstream component demand."),
    ("S3", "Press Information Bureau / Ministry of Road Transport and Highways, Road Transport and Auto Sector Get a Big Push Through GST Rationalisation, 12 September 2025",
     "https://www.pib.gov.in/PressReleasePage.aspx?PRID=2165916&lang=2&reg=48",
     "Official notice reports measures for two-wheelers, cars, tractors, buses, commercial vehicles and auto components."),
    ("S4", "PM E-DRIVE, Scheme Notifications",
     "https://pmedrive.heavyindustries.gov.in/policy_document",
     "Official notifications include August 2026 enhancement/extension for e-two-wheelers, September 2026 amendments for e-trucks and e-buses, and other scheme updates."),
    ("S5", "Mahindra & Mahindra, Results Q4 F26 and FY26, 5 May 2026",
     "https://www.mahindra.com/news-room/press-release/en/m-and-m-results-q4-f26-and-fy26",
     "M&M reported FY26 consolidated revenue of Rs 198,639 crore (+25%) and consolidated PAT of Rs 17,099 crore; its release reports FY26 auto volumes up 19% and UV volumes up 20%. Figures are company-reported and definitions in the release should be consulted."),
    ("S6", "Reuters, Hyundai Motor Group bets $15.5 billion on India as global auto hub, 13 October 2026",
     "https://www.reuters.com/world/asia-pacific/hyundai-motor-group-invest-155-billion-india-through-2032-2026-10-13/",
     "Reuters reports Hyundai Motor Group's announced $15.5 billion India investment through 2032, including R&D and capacity-related investment, underscoring competition and localization."),
    ("S7", "Reuters, Indian firms seen posting faster profit growth despite global headwinds, 8 October 2026",
     "https://www.reuters.com/world/india/indian-firms-seen-posting-faster-profit-growth-despite-global-headwinds-2026-10-08/",
     "Reuters summarizes brokerage expectations of stronger September-quarter earnings growth but also notes demand uncertainties and possible commodity-cost pressure. This is a forecast summary, not a confirmed earnings result."),
    ("S8", "Reuters, India steel ministry unlikely to seek new curbs on cheap imports, 7 October 2026",
     "https://www.reuters.com/world/china/india-steel-ministry-unlikely-seek-new-curbs-on-cheap-imports-2026-10-07/",
     "Reuters discusses steel import and pricing conditions as a potential input-cost and margin factor for steel consumers, including automotive manufacturers."),
    ("S9", "Yahoo Finance public historical market data interface (accessed at run time)",
     "https://finance.yahoo.com/",
     "Historical price data are retrieved programmatically via yfinance. Availability, corporate actions, index symbol conventions and adjustments depend on the source. The script records the actual download window and latest observation."),
    ("S10", "Nifty Indices, Nifty Auto index overview and methodology",
     "https://www.niftyindices.com/indices/equity/sectoral-indices/nifty-auto",
     "Nifty Auto is an Indian automotive-sector index calculated using free-float market capitalization. This report does not calculate historical official-index returns because Yahoo returned only one usable observation for its ^CNXAUTO symbol in this run; instead it uses a clearly labeled equally weighted basket of the included stock sample."),
]

def download_market_data():
    """Download a bundle, then individual tickers if the bundle request is sparse."""
    raw = yf.download(list(TICKERS), period=START_PERIOD, interval="1d",
                      auto_adjust=False, group_by="ticker", progress=False,
                      threads=True, timeout=45)
    print("Raw Yahoo download shape:", getattr(raw, "shape", None))
    print("Raw Yahoo column sample:", repr(list(raw.columns)[:12]) if hasattr(raw, "columns") else "no columns")
    frames = {}
    diagnostics = {}
    for ticker in TICKERS:
        df = pd.DataFrame()
        try:
            if isinstance(raw.columns, pd.MultiIndex):
                if ticker in raw.columns.get_level_values(0):
                    df = raw[ticker].copy()
                elif ticker in raw.columns.get_level_values(1):
                    df = raw.xs(ticker, axis=1, level=1).copy()
            else:
                df = raw.copy()
        except Exception as exc:
            warnings.warn(f"Could not read bundled columns for {ticker}: {exc}")
        if df.empty or "Close" not in df.columns or df["Close"].dropna().empty:
            try:
                ind = yf.Ticker(ticker).history(period=START_PERIOD, interval="1d", auto_adjust=False, actions=False)
                if ind.empty:
                    ind = yf.download(ticker, period=START_PERIOD, interval="1d",
                                      auto_adjust=False, progress=False, threads=False, timeout=45)
                if isinstance(ind.columns, pd.MultiIndex):
                    if ticker in ind.columns.get_level_values(0):
                        ind = ind[ticker].copy()
                    elif ticker in ind.columns.get_level_values(1):
                        ind = ind.xs(ticker, axis=1, level=1).copy()
                    elif "Close" in ind.columns.get_level_values(0):
                        ind.columns = ind.columns.get_level_values(0)
                    elif "Close" in ind.columns.get_level_values(1):
                        ind.columns = ind.columns.get_level_values(1)
                df = ind.copy()
            except Exception as exc:
                warnings.warn(f"Could not retrieve {ticker}: {exc}")
        diagnostics[ticker] = {"shape": list(df.shape), "columns": [str(c) for c in list(df.columns)[:10]], "has_close": "Close" in df.columns}
        if not df.empty:
            df.index = pd.to_datetime(df.index).tz_localize(None)
            df = df[~df.index.duplicated(keep="last")].sort_index()
            if "Close" in df:
                close = pd.to_numeric(df["Close"], errors="coerce")
                adj = pd.to_numeric(df["Adj Close"], errors="coerce") if "Adj Close" in df else close.copy()
                # Adjusted close is used for equities to incorporate distribution/split adjustment.
                chosen = close if ticker.startswith("^") else adj.fillna(close)
                chosen = chosen.dropna()
                diagnostics[ticker]["n_price_values_after_cleaning"] = int(len(chosen))
                diagnostics[ticker]["price_window_start"] = chosen.index.min().strftime("%Y-%m-%d") if len(chosen) else None
                diagnostics[ticker]["price_window_end"] = chosen.index.max().strftime("%Y-%m-%d") if len(chosen) else None
                if len(chosen) >= 100:
                    frames[ticker] = chosen.rename(TICKERS[ticker][0])
    if "^NSEI" not in frames:
        raise RuntimeError("Nifty 50 benchmark data could not be downloaded. Available series: " + json.dumps({"frames": list(frames), "diagnostics": diagnostics}, default=str) + ". No synthetic prices were used.")
    stock_cols = [TICKERS[t][0] for t in TICKERS if t not in ("^CNXAUTO", "^NSEI") and TICKERS[t][0] in [s.name for s in frames.values()]]
    if len(stock_cols) < 6:
        raise RuntimeError(f"Only {len(stock_cols)} individual automotive stocks were available; not enough for a meaningful sample basket.")
    prices = pd.concat(frames.values(), axis=1).sort_index()
    prices = prices[~prices.index.duplicated(keep="last")]
    prices = prices.dropna(how="all")
    # Yahoo returned only a single usable observation for ^CNXAUTO during this run.
    # Build a transparent equal-weighted proxy from daily returns of the actual downloaded automotive stocks.
    stock_returns = prices[stock_cols].pct_change(fill_method=None)
    basket_returns = stock_returns.mean(axis=1, skipna=True)
    prices["Auto sample basket"] = (1 + basket_returns.fillna(0)).cumprod() * 100
    return prices, frames

def build_metrics(prices):
    daily = prices.pct_change(fill_method=None)
    now = prices.index.max()
    periods = {"1Y": now - pd.DateOffset(years=1),
               "3Y": now - pd.DateOffset(years=3),
               "5Y": now - pd.DateOffset(years=5),
               "YTD": pd.Timestamp(year=now.year, month=1, day=1)}
    period_returns = {}
    for label, threshold in periods.items():
        chunk = prices.loc[prices.index >= threshold]
        row = {}
        for col in prices.columns:
            s = chunk[col].dropna()
            row[col] = float(s.iloc[-1] / s.iloc[0] - 1) if len(s) >= 2 else np.nan
        period_returns[label] = row
    annual_vol = daily.std(ddof=1) * np.sqrt(TRADING_DAYS)
    total_returns = prices.iloc[-1] / prices.iloc[0] - 1
    drawdowns = prices / prices.cummax() - 1
    max_dd = drawdowns.min()
    metrics = pd.DataFrame({
        "Name": [next((v[0] for k,v in TICKERS.items() if v[0]==c),c) for c in prices.columns],
        "Ticker": ["Derived, not tradable" if c=="Auto sample basket" else next((k for k,v in TICKERS.items() if v[0]==c),c) for c in prices.columns],
        "Category": ["Equal-weighted sample proxy" if c=="Auto sample basket" else next((v[1] for k,v in TICKERS.items() if v[0]==c),"Unknown") for c in prices.columns],
        "Latest close": prices.iloc[-1].values,
        "Latest observation": [prices[c].last_valid_index().strftime("%Y-%m-%d") if prices[c].last_valid_index() is not None else "n/a" for c in prices.columns],
        "5Y available-period total return": total_returns.values,
        "Annualized daily volatility": annual_vol.reindex(prices.columns).values,
        "Maximum drawdown": max_dd.reindex(prices.columns).values,
    }).set_index("Name")
    for period, row in period_returns.items():
        metrics[f"{period} return"] = pd.Series(row)
    # Add annualized return over the observed window; avoids pretending a short history is a full 5 years.
    days = max((prices.index[-1] - prices.index[0]).days, 1)
    metrics["Annualized return (observed window)"] = (1 + metrics["5Y available-period total return"]) ** (365.25 / days) - 1
    return daily, drawdowns, metrics, period_returns

def make_charts(prices, daily, drawdowns, metrics, period_returns):
    CHARTS.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid", context="notebook")
    paths = []
    # 1 benchmarks normalized
    fig, ax = plt.subplots(figsize=(9,5.4))
    benchmark_cols = [c for c in ["Auto sample basket","Nifty 50"] if c in prices]
    norm = prices[benchmark_cols].dropna(how="all")
    if len(norm):
        norm = norm.div(norm.bfill().iloc[0]).mul(100)
        for c in norm: ax.plot(norm.index, norm[c], label=c, linewidth=2)
    ax.set(title="1. Five-Year Indexed Trend: Auto Sample Basket vs Broad Market", ylabel="Indexed level (first common value = 100)", xlabel="Date")
    ax.legend()
    fig.text(.01,.005,"Caption: Index levels are rebased to 100 to compare relative price movement, not absolute index values. The final series end on the latest available observations.",fontsize=8)
    fig.tight_layout(rect=[0,.06,1,1]); p=CHARTS/"01_sector_vs_broad_market.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    # 2 company returns at 1Y / 5Y
    subset=metrics.loc[~metrics.index.isin(["Nifty Auto","Nifty 50","Auto sample basket"])].copy()
    cols=[c for c in ["1Y return","5Y return"] if c in subset]
    fig,ax=plt.subplots(figsize=(10,6.5))
    if cols:
        sub=subset[cols].sort_values(cols[0])
        sub.plot(kind="barh",ax=ax)
        ax.set(title="2. Individual Auto Stocks: One-Year vs Available Five-Year Return",xlabel="Simple price return",ylabel="")
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_:f"{x:.0%}"))
        ax.legend(["1Y return","5Y return"],loc="best")
    fig.text(.01,.005,"Caption: Returns are based on the downloaded adjusted-close history for equities and close levels for indices. '5Y' is the available period up to five years. Historical returns do not predict future returns.",fontsize=8)
    fig.tight_layout(rect=[0,.06,1,1]);p=CHARTS/"02_company_relative_returns.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    # 3 risk-return scatter
    fig,ax=plt.subplots(figsize=(9,5.8))
    x=metrics["Annualized daily volatility"];y=metrics["Annualized return (observed window)"]
    ax.scatter(x,y,s=65,alpha=.8)
    for name in metrics.index:
        if pd.notna(x[name]) and pd.notna(y[name]):
            ax.annotate(name,(x[name],y[name]),xytext=(4,4),textcoords="offset points",fontsize=7)
    ax.axhline(0,color="gray",linewidth=1,linestyle="--")
    ax.set(title="3. Observed Return vs Annualized Volatility",xlabel="Annualized daily volatility",ylabel="Annualized return over observed price window")
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_:f"{x:.0%}"));ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_:f"{x:.0%}"))
    fig.text(.01,.005,"Caption: This descriptive risk-return map is not a forward-looking efficient frontier. Volatility uses daily percentage changes annualized by sqrt(252); return annualization depends on the observed window.",fontsize=8)
    fig.tight_layout(rect=[0,.06,1,1]);p=CHARTS/"03_risk_return_map.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    # 4 drawdown indexes
    fig,ax=plt.subplots(figsize=(9,5.4))
    for c in ["Auto sample basket","Nifty 50"]:
        if c in drawdowns: ax.plot(drawdowns.index,drawdowns[c],label=c)
    ax.axhline(0,color="gray",linewidth=1)
    ax.set(title="4. Peak-to-Trough Drawdowns: Auto Sample Basket vs Nifty 50",ylabel="Drawdown from prior running peak",xlabel="Date")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_:f"{x:.0%}"));ax.legend()
    fig.text(.01,.005,"Caption: Drawdown is the percentage decline from the previous running maximum. It shows historical downside paths, not maximum possible future losses.",fontsize=8)
    fig.tight_layout(rect=[0,.06,1,1]);p=CHARTS/"04_benchmark_drawdowns.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    # 5 correlation heatmap
    cols=[c for c in daily.columns if c != "Auto sample basket" and daily[c].notna().sum()>=100]
    corr=daily[cols].corr(min_periods=80)
    fig,ax=plt.subplots(figsize=(10,8))
    sns.heatmap(corr,annot=True,fmt=".2f",vmin=-1,vmax=1,center=0,cmap="vlag",square=True,linewidths=.3,ax=ax)
    ax.set(title="5. Daily-Return Correlation Across Auto Stocks and Nifty 50")
    fig.text(.01,.005,"Caption: Pairwise correlations use overlapping daily returns when at least 80 observations are available. The computed equal-weighted sample basket is excluded here to avoid mechanical correlations with its own constituents. Correlation describes co-movement and does not establish a common cause.",fontsize=8)
    fig.tight_layout(rect=[0,.04,1,1]);p=CHARTS/"05_daily_return_correlation.png";fig.savefig(p,dpi=160);plt.close(fig);paths.append(p)
    # 6 monthly returns of Nifty auto
    sector=prices.get("Auto sample basket")
    if sector is not None:
        monthly=sector.resample("ME").last().pct_change().dropna()
        by_year=monthly.to_frame("Monthly return");by_year["Year"]=by_year.index.year;by_year["Month"]=by_year.index.strftime("%b")
        pivot=by_year.pivot(index="Year",columns="Month",values="Monthly return")
        month_order=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        pivot=pivot.reindex(columns=month_order)
        fig,ax=plt.subplots(figsize=(10,5.4))
        sns.heatmap(pivot,annot=True,fmt=".0%",center=0,cmap="vlag",mask=pivot.isna(),ax=ax)
        ax.set(title="6. Auto Sample Basket Monthly Returns by Calendar Month",xlabel="Month",ylabel="Year")
        fig.text(.01,.005,"Caption: Monthly returns are calculated from month-end index closes. Empty cells indicate no observation or no monthly return in the partial period.",fontsize=8)
        fig.tight_layout(rect=[0,.06,1,1]);p=CHARTS/"06_sector_monthly_return_heatmap.png";fig.savefig(p,dpi=160);plt.close(fig);paths.append(p)
    # Save a direct recent trend chart
    fig,ax=plt.subplots(figsize=(9,5.4))
    cols=[c for c in ["Auto sample basket","Nifty 50"] if c in prices]
    normalized=prices[cols].div(prices[cols].dropna().iloc[0]).mul(100)
    recent=normalized.loc[normalized.index>=normalized.index.max()-pd.DateOffset(years=1)]
    for c in recent: ax.plot(recent.index,recent[c],label=c)
    ax.set(title="7. Latest One-Year Sector and Market Relative Trend",ylabel="Indexed close (full-window base = 100)",xlabel="Date")
    ax.legend()
    fig.text(.01,.005,"Caption: The one-year window shows recent relative momentum; it should be considered alongside earnings, valuations, and macro conditions.",fontsize=8)
    fig.tight_layout(rect=[0,.06,1,1]);p=CHARTS/"07_recent_relative_trend.png";fig.savefig(p,dpi=170);plt.close(fig);paths.append(p)
    return paths

def pct(x, digits=1):
    if pd.isna(x): return "n/a"
    return f"{x*100:.{digits}f}%"

def add_table(doc, headers, rows, widths=None):
    table=doc.add_table(rows=1,cols=len(headers));table.style="Light Shading Accent 1";table.alignment=WD_TABLE_ALIGNMENT.CENTER
    for c,h in zip(table.rows[0].cells,headers):
        c.text=str(h);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for run in c.paragraphs[0].runs:run.bold=True;run.font.size=Pt(8.5)
    for row in rows:
        cells=table.add_row().cells
        for i,(c,val) in enumerate(zip(cells,row)):
            c.text=str(val);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for para in c.paragraphs:
                for run in para.runs:run.font.size=Pt(8)
    return table

def add_source(doc, idx, title, url, note):
    p=doc.add_paragraph()
    r=p.add_run(f"[{idx}] {title}. ");r.bold=True
    p.add_run(note+" ")
    p.add_run(url)

def build_submission_doc():
    text=SUBMISSION_TXT.read_text(encoding="utf-8").strip()
    doc=Document()
    sec=doc.sections[0];sec.top_margin=Inches(.7);sec.bottom_margin=Inches(.7);sec.left_margin=Inches(.85);sec.right_margin=Inches(.85)
    doc.styles["Normal"].font.name="Aptos";doc.styles["Normal"].font.size=Pt(11)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("AUTOMOTIVE STOCK MARKET ANALYSIS INTERN");r.bold=True;r.font.size=Pt(11);r.font.color.rgb=RGBColor(230,100,0)
    p=doc.add_heading("Task 1 — Submission Description",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for block in text.split("\n\n"):
        p=doc.add_paragraph(block);p.paragraph_format.space_after=Pt(10);p.paragraph_format.line_spacing=1.12
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=p.add_run(f"Word count: {len(text.split())}");r.italic=True;r.font.size=Pt(9)
    doc.save(SUBMISSION_DOCX)

def build_report(prices,daily,drawdowns,metrics,returns,paths,source_info):
    doc=Document()
    sec=doc.sections[0];sec.top_margin=Inches(.55);sec.bottom_margin=Inches(.55);sec.left_margin=Inches(.65);sec.right_margin=Inches(.65)
    doc.styles["Normal"].font.name="Aptos";doc.styles["Normal"].font.size=Pt(9.3)
    doc.styles["Title"].font.name="Aptos Display";doc.styles["Title"].font.size=Pt(23)
    doc.styles["Heading 1"].font.name="Aptos Display";doc.styles["Heading 1"].font.size=Pt(15)
    doc.styles["Heading 2"].font.name="Aptos Display";doc.styles["Heading 2"].font.size=Pt(11.5)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("AUTOMOTIVE STOCK MARKET ANALYSIS INTERN");r.bold=True;r.font.size=Pt(11);r.font.color.rgb=RGBColor(230,100,0)
    p=doc.add_heading("Automotive Market Overview & Trend Analysis",0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p=doc.add_paragraph("India-listed automotive equities | Historical market data + public industry research");p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    generated=datetime.now().astimezone().strftime("%d %B %Y")
    latest=prices.index.max().strftime("%d %B %Y")
    # results summary
    sector=metrics.loc["Auto sample basket"] if "Auto sample basket" in metrics.index else metrics.iloc[0]
    nifty=metrics.loc["Nifty 50"] if "Nifty 50" in metrics.index else metrics.iloc[1]
    one_y_auto=sector.get("1Y return",np.nan);one_y_nifty=nifty.get("1Y return",np.nan)
    five_y_auto=sector.get("5Y return",np.nan);five_y_nifty=nifty.get("5Y return",np.nan)
    best_col=metrics.loc[[x for x in metrics.index if x not in ["Nifty Auto","Nifty 50","Auto sample basket"]],"1Y return"].dropna()
    best_name=best_col.idxmax() if len(best_col) else "Not available"
    worst_name=best_col.idxmin() if len(best_col) else "Not available"
    best_val=best_col.max() if len(best_col) else np.nan
    worst_val=best_col.min() if len(best_col) else np.nan
    count_equities=sum(1 for x in metrics.index if x not in ["Nifty Auto","Nifty 50","Auto sample basket"])
    doc.add_heading("Executive Summary",1)
    doc.add_paragraph(
      f"This report reviews the Indian automotive equity sector using daily historical price data downloaded on {generated}. "
      f"The market-data analysis includes Nifty 50 and an equal-weighted auto-stock sample basket built from {count_equities} listed manufacturers/component businesses where sufficient data were available. Yahoo Finance returned only one usable ^CNXAUTO observation during this run, so the official Nifty Auto index was not used for historical-return calculations. "
      f"The most recent common price observation in the downloaded data is {latest}. Over the one-year measured window, the Auto sample basket returned {pct(one_y_auto)} versus {pct(one_y_nifty)} for Nifty 50; over the available five-year window, the respective returns were {pct(five_y_auto)} and {pct(five_y_nifty)}. The basket is a descriptive equal-weighted proxy, not the official Nifty Auto index. "
      f"Among included individual shares, the highest observed one-year return was {best_name} ({pct(best_val)}), and the lowest was {worst_name} ({pct(worst_val)}). These are historical, sample-dependent outcomes, not forecasts or recommendations."
    )
    doc.add_paragraph(
      "The industry context is constructive but not one-directional. SIAM's August 2026 data reported record-high August sales in covered passenger-vehicle, three-wheeler and two-wheeler categories, while explicitly noting the lower prior-year base. Policy changes to GST in September 2025 improved price affordability in selected vehicle categories and auto parts. The PM E-DRIVE scheme continues to evolve, while fresh investment plans and export ambitions reinforce India's manufacturing opportunity. Against these positives, equity outcomes remain sensitive to valuations, commodity and energy costs, financing conditions, currency movements, competition, model launches and execution."
    )
    doc.add_heading("Key Conclusions at a Glance",2)
    for text in [
      f"Benchmark check: Auto sample basket 1Y return {pct(one_y_auto)} compared with Nifty 50 {pct(one_y_nifty)}; the gap helps distinguish the selected sample's performance from the broad market.",
      f"Cross-section: {best_name} was the strongest and {worst_name} the weakest included share over the observed one-year window; rank alone does not establish value or future performance.",
      "Demand: SIAM's August 2026 release reported 439,309 passenger vehicles, 93,764 three-wheelers and 2,034,698 two-wheelers sold domestically in the covered categories, with a lower-base caveat [S1].",
      "Profitability: growth must translate into margins, cash flow and return on capital; volume growth alone is not enough.",
      "Investment posture: use this as a screening/reporting framework, not personalized investment advice. Check current valuations, latest exchange filings and risk tolerance before any decision."
    ]: doc.add_paragraph(text,style="List Bullet")

    doc.add_heading("1. Objective and Scope",1)
    doc.add_paragraph("The task is to analyze public information about India's automotive stock market, identify trends and industry factors, visualize evidence, and explain how those factors can influence listed auto companies. The research reviews Nifty Auto as the official sector-index reference and Nifty 50 as the broad-market benchmark. For five-year historical comparison, the report uses an equal-weighted basket of the included listed auto shares because Yahoo's ^CNXAUTO feed supplied only one usable observation. The sample covers passenger vehicles, two- and three-wheelers, commercial vehicles, selected component suppliers and batteries and is not an exhaustive stock universe or a recommendation list.")
    doc.add_paragraph("Companies in the downloaded sample include Maruti Suzuki, Mahindra & Mahindra, Bajaj Auto, Eicher Motors, Hero MotoCorp, TVS Motor, Ashok Leyland, Bharat Forge, Bosch and Exide Industries, subject to market-data availability. Each business has different exposure to vehicle segments, exports, premiumization, fleet demand, commodity input costs and technology transition.")
    doc.add_heading("2. Data Sources and Methodology",1)
    doc.add_paragraph(f"Price-source method: yfinance retrieval from Yahoo Finance's public historical market-data interface. The script queried a five-year daily window on {generated}; source tickers, first/last observations and valid row counts are recorded in outputs/run_metadata.json and outputs/historical_prices.csv. Indexes use supplied close levels, while equity returns use adjusted close when available, falling back to close where not. That treatment aims to account for stock splits and distributions, but it does not make stock returns and index levels perfectly identical concepts.")
    doc.add_paragraph("Returns are computed as P_end/P_start − 1. Daily volatility is the sample standard deviation of daily percentage returns multiplied by sqrt(252). Maximum drawdown is the minimum of P_t / max(P_0 ... P_t) − 1. The annualized return shown for risk-return comparisons compounds the observed period return to an annual rate based on calendar days and must be read cautiously when history is less than five years. Daily return correlations use overlapping observations. The code does not fill missing prices backwards or synthesize prices.")
    doc.add_paragraph("Limitations: public data providers can revise adjusted histories, have occasional gaps, or change symbol availability. The five-year period is rolling and ends at the latest price observation available at run time, not necessarily the current calendar day. Individual-stock prices are affected by corporate actions and idiosyncratic news. Index price levels are not total-return indices. Conclusions therefore emphasize relative context, risk, and drivers rather than precise investable total-return estimates.")
    doc.add_heading("3. Historical Market Performance",1)
    doc.add_paragraph(f"The sample-basket return is computed from equal-weighted daily returns of the included automotive companies, not copied from an official index. It helps illustrate the collective behavior of the selected sample relative to Nifty 50 but can differ from a free-float market-cap-weighted index.")
    doc.add_paragraph(f"In the individual-stock sample, {best_name} led the one-year table at {pct(best_val)}, while {worst_name} was lowest at {pct(worst_val)}. That range illustrates why treating 'auto stocks' as one uniform trade is misleading. A two-wheeler exporter, a domestic SUV manufacturer, a commercial-vehicle producer and a component maker have distinct revenue mixes, margin drivers and valuation expectations.")
    doc.add_paragraph("Return ranking is a starting point, not an investment conclusion. A high historical return may be accompanied by a higher valuation, greater drawdown or elevated expectations; a low return may reflect structural problems, temporary conditions, or a valuation reset. Fundamental context is needed to distinguish those cases.")
    doc.add_heading("4. Industry Conditions and Demand Drivers",1)
    doc.add_heading("4.1 Sales volumes and base effects",2)
    doc.add_paragraph("SIAM's release dated 15 September 2026 reports August domestic sales of 439,309 passenger vehicles, 93,764 three-wheelers and 2,034,698 two-wheelers for its covered categories. It reports year-on-year growth rates of 36.5%, 22.8% and 10.5%, respectively, and also cautions that the comparison benefits from a lower base in August 2025 [S1]. For analysts, this means that sequential momentum, retail registrations, inventory, model mix and company disclosures should be evaluated alongside the headline annual growth rate.")
    doc.add_paragraph("Passenger vehicles are influenced by household income, urban consumption, replacement demand, SUV/premium mix, launch cycles and availability. Two-wheelers are particularly sensitive to affordability, rural cash flows, entry-level demand, finance availability and export markets. Commercial vehicles depend more heavily on freight activity, infrastructure investment, fleet utilization, replacement cycles and financing. Tractors add exposure to monsoon performance and farm economics.")
    doc.add_heading("4.2 Tax and policy transmission",2)
    doc.add_paragraph("The Government of India's September 2025 GST rationalization included rate changes for selected motorcycles up to 350 cc, small cars and auto components, with official communications describing the aim of improving affordability and supporting upstream suppliers [S2, S3]. The policy transmission is not automatic: lower tax can support unit demand, but the effect on manufacturers also depends on price pass-through, discounting, supplier pricing, mix and incremental fixed costs. Large-vehicle categories have different treatment, so the policy should not be summarized as one uniform cut for all vehicles.")
    doc.add_heading("4.3 Electrification and technology",2)
    doc.add_paragraph("The PM E-DRIVE scheme notification page lists updates during 2026, including an enhancement/extension for e-two-wheelers in August and amendments related to e-trucks and e-buses in September [S4]. Such support can influence adoption and economics by segment, but investors should distinguish scheme duration, category eligibility, incentive exhaustion, charging infrastructure, battery cost, localization requirements and post-subsidy product competitiveness. EV transition is also a capital-allocation question: incumbent OEMs and suppliers may benefit from new products but can face write-downs or margin pressure on legacy platforms.")
    doc.add_heading("4.4 Exports and investment",2)
    doc.add_paragraph("On 13 October 2026, Reuters reported Hyundai Motor Group's announced $15.5 billion investment in India through 2032, including research and production-related plans [S6]. This reinforces the country's role in international manufacturing networks while also raising the competitive bar for product development, scale, localization, cost and exports. Export potential diversifies revenue, but it introduces currency, tariff, regulatory, shipping and demand-cycle risks in destination markets.")
    doc.add_heading("5. Company and Segment Context",1)
    doc.add_paragraph("Maruti Suzuki offers a high-volume passenger-vehicle and export lens; analysis should monitor market share, channel inventory, product mix, capacity utilization and small-car demand. Mahindra & Mahindra spans SUVs, light commercial vehicles, farm equipment and other businesses, so consolidated earnings require careful segment reading. Its FY26 release reported consolidated revenue of Rs 198,639 crore, consolidated PAT of Rs 17,099 crore, FY26 auto volumes growth of 19% and UV volumes growth of 20% [S5]. These company-reported numbers show why segment leadership and execution can matter, but consolidated metrics should not be compared directly with a pure-play OEM without accounting for business mix.")
    doc.add_paragraph("Bajaj Auto, Eicher Motors, Hero MotoCorp and TVS Motor cover different points across motorcycles, scooters, premiumization, exports and three-wheelers. Their earnings sensitivities vary with the premium vs commuter mix and geography. Ashok Leyland is more exposed to commercial-vehicle replacement, freight and fleet economics. Bharat Forge and Bosch depend on manufacturing scale, technology content and global OEM order patterns; Exide has exposure to replacement batteries and evolving EV/energy-storage supply chains. These descriptions are analytical context, not current valuation recommendations.")
    doc.add_heading("6. Reading the Charts",1)
    chart_notes=[
      ("Figure 1. Five-Year Indexed Trend: Auto Sample Basket vs Broad Market","Rebases the equal-weighted automotive sample basket and Nifty 50 to a common starting value so relative direction and compounding can be compared. It does not compare nominal index levels or represent the official Nifty Auto index."),
      ("Figure 2. Individual Auto Stocks: One-Year vs Available Five-Year Return","Shows dispersion across listed companies. The cross-section supports avoiding one-size-fits-all sector conclusions; history lengths and corporate actions may differ."),
      ("Figure 3. Observed Return vs Annualized Volatility","Maps return against realized risk. A high return/high volatility point is not automatically better than a lower-risk outcome, and past return is not expected return."),
      ("Figure 4. Benchmark Drawdowns","Shows how far the equal-weighted auto-stock sample basket and Nifty 50 moved below prior peaks. Drawdowns help communicate the lived path of risk, which a single return number hides."),
      ("Figure 5. Daily-Return Correlation","Shows which stocks tended to co-move over the overlapping daily window. Correlation can change in stress periods and does not itself make a diversified portfolio."),
      ("Figure 6. Auto Sample Basket Monthly Return Heatmap","Highlights calendar periods of positive and negative returns in the selected auto-stock basket without assuming that a seasonality pattern will repeat."),
      ("Figure 7. Latest One-Year Relative Trend","Provides more recent context for the sector-vs-market relationship; it remains backward-looking and should be connected to valuation and fundamentals.")
    ]
    for i,(img,(title,note)) in enumerate(zip(paths,chart_notes)):
      heading_para=doc.add_heading(title,2)
      heading_para.paragraph_format.keep_with_next=True
      doc.add_picture(str(img),width=Inches(5.9));doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
      note_para=doc.add_paragraph(note)
      note_para.paragraph_format.keep_together=True
    doc.add_heading("7. Key Trends and What Could Move Stocks",1)
    for head,body in [
      ("Demand rebound vs normalization","Monthly growth can look unusually strong after a weak comparison period or tax/policy change. Analysts should compare multi-month trends, wholesales and retail registrations, inventory, booking conversion and management guidance."),
      ("Mix and margins matter as much as volume","A shift toward SUVs, premium motorcycles or higher-spec vehicles can lift revenue per unit, but competition, discounts, warranty provisions, freight and raw-material costs can offset those gains."),
      ("Financing and interest rates","Vehicle purchases often depend on monthly payments and lender availability. Higher borrowing costs can hurt affordability; looser financing can support demand but may increase household or dealer leverage risk."),
      ("Commodities, energy and currency","Steel, aluminium, copper, rubber, plastics, semiconductors and battery materials affect input costs. Currency weakness can raise imported-input cost while benefiting exporters; the net effect depends on hedging and localization. Reuters has recently highlighted possible commodity pressure even amid a more positive earnings outlook [S7, S8]."),
      ("EV investment and policy execution","The long-term opportunity involves electric two-wheelers, buses, trucks and passenger vehicles, software, charging and components. Risks include subsidy changes, low utilization, price competition, battery life, warranty liability and asset obsolescence."),
      ("Valuation and expectations","A good business can be a weak stock investment if the price already assumes exceptional growth. Compare valuation multiples with growth, ROCE, operating cash flow, capital intensity, net cash/debt and peers rather than selecting only the best price chart."),
    ]:
      doc.add_heading(head,2);doc.add_paragraph(body)
    doc.add_heading("8. Strategic Recommendations for an Analyst",1)
    recs=[
      ("Use a two-layer research process","Start with the sector index and segment-level demand; then test each company’s market share, realizations, product mix, operating margin, cash conversion, capital expenditure and valuation. This reduces the risk of treating a sector tailwind as proof that every company benefits equally."),
      ("Maintain a balanced watchlist rather than a buy list","Group companies by exposure: passenger vehicles, two/three-wheelers, commercial vehicles, diversified OEMs, components and batteries. Compare peers within a similar business model and record the catalyst, valuation concern and invalidation condition for each thesis."),
      ("Track leading indicators monthly","Monitor SIAM wholesales, FADA retail registrations when published, exports, production, inventory, dealer commentary, new launches, financing rates, commodity prices and policy notifications. Separate industry volume growth from company market-share gains."),
      ("Measure downside and concentration","Use drawdown and volatility charts alongside return. Avoid assuming that high historical performers will lead in future; sector correlations can rise during risk-off events. Define acceptable loss, diversification and position-sizing rules independently of this sector report."),
      ("Update views after disclosures","Reconcile company monthly sales with quarterly financials, consolidated vs standalone reporting, segment profit, free cash flow, debt, capex, warranty/quality news and capacity utilisation. Use official exchange filings and annual reports for final decisions."),
      ("Keep an explicit evidence log","For each conclusion, record source, date, measured variable, counterargument and next checkpoint. This supports reproducibility and stops a plausible macro narrative from turning into an unsupported stock forecast.")
    ]
    for head,body in recs:
      doc.add_heading(head,2);doc.add_paragraph(body)
    doc.add_heading("9. Risks, Limitations and Further Research",1)
    for item in [
      "Market-data availability and corporate-action adjustments vary by ticker and vendor. The report records the observed source window and does not fabricate missing prices.",
      "A sample of selected liquid listed companies does not represent every OEM, supplier, EV pure-play or small-cap stock. Survivorship and selection bias are possible.",
      "Past price correlations are not stable structural parameters. Short windows, crisis periods and index composition changes can alter estimates.",
      "Market prices incorporate expectations before reported sales or earnings. A strong industry metric does not guarantee a positive share-price response if expectations were higher.",
      "Industry data (wholesales, retails, production and exports) use different measurement points. Comparisons must preserve the source's definition and coverage.",
      "Policy benefits differ by vehicle category and company position. Confirm current eligibility and notifications at the official source before interpreting an incentive as a forecast of future earnings.",
      "This report is educational research, not personalized financial advice, and does not set a buy/sell target. Validate current quotes, filings, valuation and suitability independently."
    ]:doc.add_paragraph(item,style="List Bullet")
    doc.add_heading("10. Conclusion",1)
    doc.add_paragraph("India's automotive sector offers multiple demand and technology themes, but the listed-equity opportunity is heterogeneous. Recent official industry figures point to strong August 2026 volumes in covered categories, government tax changes and evolving EV support affect demand economics, and new investment plans underline strategic manufacturing importance. However, stock performance depends on the earnings and cash flow captured by each firm, valuation paid, cost control and execution. The most robust analyst approach combines sector data with company fundamentals, uses return and risk together, keeps sources dated, and updates the view after each reporting cycle. The quantitative charts in this report are descriptive evidence to support that process—not predictions.")
    doc.add_heading("11. References and Public Sources",1)
    for idx,title,url,note in SOURCES:
      doc.add_paragraph(f"[{idx}] {title}. {note} URL: {url}")
    doc.add_paragraph(f"Market-data retrieval timestamp: {source_info['generated_at_local']}. Requested window: {source_info['download_period_requested']}; observed date range: {source_info['observed_start']} to {source_info['observed_end']}.")
    doc.add_paragraph(f"Usable source histories: {len(source_info['tickers_available'])}. No synthetic market prices were used. Detailed ticker coverage, row counts and generated chart names are preserved in outputs/run_metadata.json.")

    doc.save(REPORT_PATH)

def main():
    OUT.mkdir(parents=True,exist_ok=True);CHARTS.mkdir(parents=True,exist_ok=True)
    prices,downloaded=download_market_data()
    prices.to_csv(OUT/"historical_prices.csv",index_label="Date")
    daily,drawdowns,metrics,period_returns=build_metrics(prices)
    metrics.to_csv(OUT/"return_risk_metrics.csv")
    daily.to_csv(OUT/"daily_returns.csv",index_label="Date")
    drawdowns.to_csv(OUT/"drawdowns.csv",index_label="Date")
    paths=make_charts(prices,daily,drawdowns,metrics,period_returns)
    rows={}
    for ticker,(name,category) in TICKERS.items():
        if name in prices.columns:
            rows[ticker]={"name":name,"category":category,"first_observation":prices[name].first_valid_index().strftime("%Y-%m-%d"),"last_observation":prices[name].last_valid_index().strftime("%Y-%m-%d"),"n_observations":int(prices[name].count())}
    info={"generated_at_local":datetime.now().astimezone().isoformat(timespec="seconds"),"data_source":"Yahoo Finance via yfinance public historical interface","download_period_requested":START_PERIOD,"observed_start":prices.index.min().strftime("%Y-%m-%d"),"observed_end":prices.index.max().strftime("%Y-%m-%d"),"tickers_available":rows,"synthetic_data_used":False,"charts_generated":[p.name for p in paths]}
    (OUT/"run_metadata.json").write_text(json.dumps(info,indent=2),encoding="utf-8")
    (OUT/"sources.json").write_text(json.dumps([{"id":a,"title":b,"url":c,"note":d} for a,b,c,d in SOURCES],indent=2),encoding="utf-8")
    build_report(prices,daily,drawdowns,metrics,period_returns,paths,info)
    build_submission_doc()
    print(f"Downloaded {len(prices.columns)} usable market series from {prices.index.min().date()} to {prices.index.max().date()}.")
    print("Synthetic data used: False")
    print("Benchmark returns:")
    print(metrics.loc[[x for x in ["Auto sample basket","Nifty 50"] if x in metrics.index],["1Y return","3Y return","5Y return","Annualized daily volatility","Maximum drawdown"]].to_string())
    print(f"Charts created: {len(paths)}")
    print(f"Report created: {REPORT_PATH}")

if __name__ == "__main__":
    main()
