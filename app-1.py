"""
NEPSE FloorSheet Intelligence — Institutional Order Flow Analyzer
Full-featured Streamlit app with column mapping, symbol-wise analysis,
metaorder detection, Smart Money Score, and export capabilities.
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
import warnings
import io
import csv
from datetime import datetime

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="NEPSE FloorSheet Intelligence",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# THEME / CSS
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #070a0f;
    color: #cdd6e0;
}
.stApp { background: radial-gradient(ellipse at 20% 10%, #0d1f35 0%, #070a0f 60%); }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #0b0f18;
    border-right: 1px solid #1a2535;
}

/* Headings */
h1,h2,h3 {
    font-family: 'IBM Plex Mono', monospace !important;
    color: #38bdf8 !important;
    letter-spacing: -0.5px;
}
h4,h5,h6 { color: #94a3b8 !important; }

/* Metrics */
[data-testid="metric-container"] {
    background: linear-gradient(135deg,#0f1c2e,#111827);
    border: 1px solid #1e3a5f;
    border-top: 2px solid #38bdf8;
    border-radius: 10px;
    padding: 14px;
    box-shadow: 0 4px 20px rgba(0,0,0,.5);
}
[data-testid="stMetricValue"] {
    font-family: 'IBM Plex Mono', monospace !important;
    color: #38bdf8 !important;
    font-size: 1.8rem !important;
}
[data-testid="stMetricLabel"] { color: #64748b !important; font-size: .78rem !important; }
[data-testid="stMetricDelta"]  { font-family: 'IBM Plex Mono', monospace !important; }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg,#0ea5e9,#0284c7);
    color: #fff;
    font-weight: 600;
    border: none;
    border-radius: 7px;
    padding: 9px 22px;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: .3px;
    transition: all .2s;
}
.stButton > button:hover {
    background: linear-gradient(135deg,#38bdf8,#0ea5e9);
    box-shadow: 0 0 18px rgba(56,189,248,.35);
}

/* File uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed #1e3a5f;
    border-radius: 10px;
    padding: 18px;
    background: #0b1220;
}
[data-testid="stFileUploader"]:hover { border-color: #38bdf8; }

/* Selectbox */
.stSelectbox label { color: #94a3b8 !important; font-size: .82rem !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: #0b1220; border-radius: 8px; padding: 3px; }
.stTabs [data-baseweb="tab"] {
    color: #64748b;
    font-family: 'IBM Plex Mono', monospace;
    font-size: .82rem;
}
.stTabs [aria-selected="true"] {
    background: #0f2036 !important;
    color: #38bdf8 !important;
    border-radius: 6px;
}

/* Custom cards */
.card {
    background: linear-gradient(135deg,#0f1c2e,#111827);
    border: 1px solid #1e3a5f;
    border-radius: 10px;
    padding: 18px 20px;
    margin: 6px 0;
}
.score-card {
    background: linear-gradient(135deg,#0f1c2e,#0d2540);
    border: 2px solid #1e3a5f;
    border-radius: 12px;
    padding: 22px;
    text-align: center;
    box-shadow: 0 8px 32px rgba(0,0,0,.6);
}
.score-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 3.6rem;
    font-weight: 700;
    line-height: 1;
}
.insight { 
    background: #0b1525;
    border: 1px solid #1e3a5f;
    border-left: 4px solid #38bdf8;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 6px 0;
    font-size: .88rem;
    line-height: 1.65;
}
.flag {
    background: rgba(239,68,68,.07);
    border: 1px solid rgba(239,68,68,.25);
    border-left: 4px solid #ef4444;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 6px 0;
    font-size: .88rem;
}
.success {
    background: rgba(34,197,94,.07);
    border: 1px solid rgba(34,197,94,.25);
    border-left: 4px solid #22c55e;
    border-radius: 8px;
    padding: 14px 16px;
    margin: 6px 0;
    font-size: .88rem;
}
.sec-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: .7rem;
    letter-spacing: 2px;
    color: #334155;
    text-transform: uppercase;
    margin: 22px 0 8px;
    padding-bottom: 6px;
    border-bottom: 1px solid #0f2036;
}
/* Mapping grid labels */
.map-label {
    font-size: .75rem;
    color: #64748b;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: .5px;
    margin-bottom: 2px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHART THEME CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

BG      = "#070a0f"
PANEL   = "#0b1220"
GRID    = "#1a2535"
TEXT    = "#94a3b8"
BLUE    = "#38bdf8"
GREEN   = "#22c55e"
RED     = "#ef4444"
AMBER   = "#f59e0b"
PURPLE  = "#a78bfa"

def _theme(fig, axes=None):
    fig.patch.set_facecolor(PANEL)
    for ax in (axes or fig.get_axes()):
        ax.set_facecolor(BG)
        ax.tick_params(colors=TEXT, labelsize=8.5)
        ax.xaxis.label.set_color(TEXT)
        ax.yaxis.label.set_color(TEXT)
        ax.title.set_color(BLUE)
        for sp in ax.spines.values():
            sp.set_edgecolor(GRID)
        ax.grid(color=GRID, linestyle="--", linewidth=0.5, alpha=0.6)

def fmt_vol(ax):
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE KEYS
# ══════════════════════════════════════════════════════════════════════════════

MAPPING_KEYS = ["col_sn","col_contract","col_symbol","col_buyer",
                "col_seller","col_qty","col_rate","col_amount"]

def init_state():
    for k in MAPPING_KEYS:
        if k not in st.session_state:
            st.session_state[k] = None
    if "mapping_confirmed" not in st.session_state:
        st.session_state.mapping_confirmed = False
    if "df_raw" not in st.session_state:
        st.session_state.df_raw = None

init_state()

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_raw(file_bytes: bytes, fname: str) -> tuple[pd.DataFrame | None, str]:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        df = df.dropna(how="all").reset_index(drop=True)
        return df, ""
    except Exception as e:
        return None, str(e)

# ══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING HEURISTICS  (auto-suggest)
# ══════════════════════════════════════════════════════════════════════════════

HINTS = {
    "col_sn":       ["sn","serial","s.n","s.n.","no","#"],
    "col_contract": ["contract","contract no","contractno","trade id","tradeid"],
    "col_symbol":   ["symbol","stock","scrip","ticker","company","script","stock symbol","scrip symbol"],
    "col_buyer":    ["buyer","buyer broker","buyer broker no","buy broker","buyer no","buyerbroker"],
    "col_seller":   ["seller","seller broker","seller broker no","sell broker","seller no","sellerbroker"],
    "col_qty":      ["quantity","qty","shares","volume","units","no of shares","num"],
    "col_rate":     ["rate","price","ltp","last price","traded price","rate(rs)","rate (rs)"],
    "col_amount":   ["amount","total","value","total value","turnover","amount(rs)","amount (rs)"],
}

def guess_mapping(cols: list[str]) -> dict[str, str | None]:
    result = {}
    lower_cols = {c.strip().lower(): c for c in cols}
    for key, hints in HINTS.items():
        matched = None
        for h in hints:
            if h in lower_cols:
                matched = lower_cols[h]
                break
        if matched is None:
            for h in hints:
                for lc, orig in lower_cols.items():
                    if h in lc or lc in h:
                        matched = orig
                        break
                if matched:
                    break
        result[key] = matched
    return result

# ══════════════════════════════════════════════════════════════════════════════
# DATA PREPARATION (post-mapping)
# ══════════════════════════════════════════════════════════════════════════════

def prepare_df(df_raw: pd.DataFrame, mapping: dict) -> tuple[pd.DataFrame | None, list[str]]:
    """Rename columns per mapping, coerce types, clean."""
    warns = []
    rename = {v: k.replace("col_", "") for k, v in mapping.items() if v}
    df = df_raw.rename(columns=rename).copy()

    needed = ["buyer", "seller", "qty", "rate"]
    missing = [n for n in needed if n not in df.columns]
    if missing:
        return None, [f"❌ Required columns missing after mapping: {missing}"]

    for c in ["qty", "rate", "amount"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    for c in ["buyer", "seller"]:
        df[c] = df[c].astype(str).str.strip().str.upper()

    before = len(df)
    df = df.dropna(subset=["qty","rate"]).copy()
    df = df[df["qty"] > 0].copy()
    dropped = before - len(df)
    if dropped:
        warns.append(f"⚠️ Dropped {dropped} rows with invalid qty/rate.")

    if "sn" in df.columns:
        df["sn"] = pd.to_numeric(df["sn"], errors="coerce")
        df = df.sort_values("sn").reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    if "symbol" not in df.columns:
        df["symbol"] = "ALL"
        warns.append("⚠️ No symbol column mapped — treating all trades as one stock.")

    df["symbol"] = df["symbol"].astype(str).str.strip().str.upper()
    df["trade_sign"] = 1

    if "amount" not in df.columns:
        df["amount"] = df["qty"] * df["rate"]

    return df, warns

# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def broker_analysis(df: pd.DataFrame) -> pd.DataFrame:
    buy  = df.groupby("buyer")["qty"].sum().reset_index().rename(columns={"buyer":"broker","qty":"buy_qty"})
    sell = df.groupby("seller")["qty"].sum().reset_index().rename(columns={"seller":"broker","qty":"sell_qty"})
    merged = pd.merge(buy, sell, on="broker", how="outer").fillna(0)
    merged["net_qty"]   = merged["buy_qty"] - merged["sell_qty"]
    merged["total_qty"] = merged["buy_qty"] + merged["sell_qty"]
    merged["net_pct"]   = merged["net_qty"] / merged["total_qty"].replace(0, np.nan) * 100
    return merged.sort_values("total_qty", ascending=False).reset_index(drop=True)

def acf_analysis(df: pd.DataFrame, max_lags: int = 20) -> dict:
    signs = df["trade_sign"].values.astype(float)
    n = len(signs)
    mu = signs.mean(); var = signs.var()
    acf_vals = []
    for lag in range(1, min(max_lags + 1, n)):
        cov = np.mean((signs[:n-lag] - mu) * (signs[lag:] - mu))
        acf_vals.append(cov / var if var > 0 else 0.0)
    lags = list(range(1, len(acf_vals) + 1))
    ci   = 1.96 / np.sqrt(n)
    sig  = [l for l, v in zip(lags, acf_vals) if abs(v) > ci]
    pos5 = sum(1 for v in acf_vals[:5] if v > ci)
    return {"acf": acf_vals, "lags": lags, "ci": ci,
            "sig_lags": sig, "pos5": pos5,
            "mean5": float(np.mean(acf_vals[:5])) if acf_vals else 0.0}

def detect_metaorders(df: pd.DataFrame, min_run: int = 3) -> pd.DataFrame:
    rows = []
    ref_col = "sn" if "sn" in df.columns else None

    def _scan(grp, direction: str, broker: str):
        grp = grp.reset_index(drop=True)
        run, vol, prices = 1, grp.loc[0,"qty"], [grp.loc[0,"rate"]]
        for i in range(1, len(grp)):
            gap = 1
            if ref_col:
                g1 = grp.loc[i-1, ref_col]; g2 = grp.loc[i, ref_col]
                if pd.notna(g1) and pd.notna(g2):
                    gap = abs(g2 - g1)
            if gap <= 15:
                run += 1; vol += grp.loc[i,"qty"]; prices.append(grp.loc[i,"rate"])
            else:
                if run >= min_run:
                    cv = np.std(prices)/np.mean(prices)*100 if np.mean(prices) else 0
                    rows.append({"broker":broker,"direction":direction,"run_len":run,
                                 "total_qty":vol,"avg_rate":np.mean(prices),
                                 "price_cv":round(cv,2),
                                 "suspicion":"HIGH" if run>=5 and cv<1 else "MODERATE"})
                run, vol, prices = 1, grp.loc[i,"qty"], [grp.loc[i,"rate"]]
        if run >= min_run:
            cv = np.std(prices)/np.mean(prices)*100 if np.mean(prices) else 0
            rows.append({"broker":broker,"direction":direction,"run_len":run,
                         "total_qty":vol,"avg_rate":np.mean(prices),
                         "price_cv":round(cv,2),
                         "suspicion":"HIGH" if run>=5 and cv<1 else "MODERATE"})

    for broker, grp in df.groupby("buyer"):
        grp = grp.sort_values(ref_col) if ref_col else grp
        _scan(grp.reset_index(drop=True), "BUY", broker)
    for broker, grp in df.groupby("seller"):
        grp = grp.sort_values(ref_col) if ref_col else grp
        _scan(grp.reset_index(drop=True), "SELL", broker)

    if not rows:
        return pd.DataFrame(columns=["broker","direction","run_len","total_qty",
                                      "avg_rate","price_cv","suspicion"])
    return pd.DataFrame(rows).sort_values("total_qty", ascending=False).reset_index(drop=True)

def size_analysis(df: pd.DataFrame) -> dict:
    qty = df["qty"].dropna()
    q25, q75 = qty.quantile(.25), qty.quantile(.75)
    iqr = q75 - q25
    threshold = max(q75 + 1.5*iqr, qty.mean() + 2.5*qty.std())
    large = df[df["qty"] > threshold]
    return {"mean":qty.mean(),"std":qty.std(),"median":qty.median(),
            "q25":q25,"q75":q75,"threshold":threshold,
            "large":large,"large_count":len(large),
            "large_vol":large["qty"].sum(),"qty":qty}

def volume_profile(df: pd.DataFrame) -> pd.DataFrame:
    vbp = df.groupby("rate").agg(
        total_qty =("qty","sum"),
        trades    =("qty","count"),
    ).reset_index().sort_values("rate")
    vbp["pct"] = vbp["total_qty"] / vbp["total_qty"].sum() * 100
    vbp["cum_qty"] = vbp["total_qty"].cumsum()
    return vbp

def pressure_stats(df: pd.DataFrame, broker_df: pd.DataFrame) -> dict:
    net_buy  = broker_df[broker_df["net_qty"]>0]["net_qty"].sum()
    net_sell = abs(broker_df[broker_df["net_qty"]<0]["net_qty"].sum())
    total = net_buy + net_sell
    ratio = net_buy/total*100 if total > 0 else 50
    return {
        "total_qty":   df["qty"].sum(),
        "trade_count": len(df),
        "buyers":      df["buyer"].nunique(),
        "sellers":     df["seller"].nunique(),
        "net_buy":     net_buy,
        "net_sell":    net_sell,
        "buy_ratio":   ratio,
    }

def smart_money_score(broker_df, acf_d, sz_d, meta_df, pr) -> dict:
    total_vol = pr["total_qty"]
    # Component 1 — broker net flow ±4
    total_net = pr["net_buy"] + pr["net_sell"]
    c1 = (pr["net_buy"]/total_net - 0.5)*8 if total_net > 0 else 0
    c1 = float(np.clip(c1, -4, 4))
    # Component 2 — ACF persistence ±2
    c2 = float(np.clip(acf_d["mean5"]*20, -2, 2))
    # Component 3 — trade size anomaly ±2
    lv_pct = sz_d["large_vol"]/(total_vol+1e-9)
    med_net = float(broker_df["net_pct"].median()) if len(broker_df)>0 else 0
    c3 = float(np.clip((med_net/100)*2*(1+lv_pct), -2, 2))
    # Component 4 — metaorder bias ±2
    if len(meta_df) > 0:
        bv = meta_df[meta_df["direction"]=="BUY"]["total_qty"].sum()
        sv = meta_df[meta_df["direction"]=="SELL"]["total_qty"].sum()
        mt = bv+sv
        c4 = float(np.clip((bv-sv)/mt*2, -2, 2)) if mt > 0 else 0
    else:
        c4 = 0.0
    score = float(np.clip(c1+c2+c3+c4, -10, 10))
    comps = {"Broker Net Flow":round(c1,2),"ACF Persistence":round(c2,2),
             "Size Anomaly":round(c3,2),"Metaorder Bias":round(c4,2)}
    signs = [np.sign(v) for v in comps.values() if v!=0]
    conf  = round(abs(sum(signs))/len(signs)*100) if signs else 50
    if score>=6:   interp,col="🟢 Strong Bullish — Accumulation",GREEN
    elif score>=3: interp,col="🟡 Weak Bullish — Mild Accumulation","#86efac"
    elif score>=-2:interp,col="⚪ Neutral — No Clear Bias",TEXT
    elif score>=-5:interp,col="🟠 Weak Bearish — Mild Distribution",AMBER
    else:          interp,col="🔴 Strong Bearish — Distribution",RED
    return {"score":score,"interp":interp,"color":col,"comps":comps,"conf":conf}

# ══════════════════════════════════════════════════════════════════════════════
# CHARTS
# ══════════════════════════════════════════════════════════════════════════════

def chart_broker(broker_df: pd.DataFrame, top_n: int) -> plt.Figure:
    top = broker_df.head(top_n).copy()
    x = np.arange(len(top)); w = .35
    fig, ax = plt.subplots(figsize=(12,5))
    _theme(fig,[ax])
    ax.bar(x-w/2, top["buy_qty"],  w, color=GREEN, alpha=.8, label="Buy Qty",  zorder=3)
    ax.bar(x+w/2, top["sell_qty"], w, color=RED,   alpha=.8, label="Sell Qty", zorder=3)
    for i,(b,s) in enumerate(zip(top["buy_qty"],top["sell_qty"])):
        net=b-s; c=GREEN if net>0 else RED
        ax.annotate("▲" if net>0 else "▼",(x[i],max(b,s)*1.02),ha="center",va="bottom",
                    fontsize=8,color=c)
    ax.set_xticks(x)
    ax.set_xticklabels(top["broker"].astype(str), rotation=45, ha="right", fontsize=8)
    ax.set_title(f"Top {top_n} Brokers — Buy vs Sell Quantity",fontsize=12,pad=12)
    ax.set_xlabel("Broker ID"); ax.set_ylabel("Quantity"); fmt_vol(ax)
    ax.legend(facecolor=PANEL,edgecolor=GRID,labelcolor=TEXT,fontsize=9)
    fig.tight_layout(); return fig

def chart_net_position(broker_df: pd.DataFrame, top_n: int) -> plt.Figure:
    accs  = broker_df[broker_df["net_qty"]>0].nlargest(top_n//2+2,"net_qty")
    dists = broker_df[broker_df["net_qty"]<0].nsmallest(top_n//2+2,"net_qty")
    top   = pd.concat([accs,dists]).sort_values("net_qty",ascending=False)
    colors= [GREEN if v>0 else RED for v in top["net_qty"]]
    fig,ax = plt.subplots(figsize=(12,4))
    _theme(fig,[ax])
    ax.barh(top["broker"].astype(str), top["net_qty"], color=colors, alpha=.8)
    ax.axvline(0, color=TEXT, linewidth=.8)
    ax.set_title("Net Position per Broker (Buy − Sell)",fontsize=12,pad=10)
    ax.set_xlabel("Net Quantity")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v,_:f"{v:,.0f}"))
    fig.tight_layout(); return fig

def chart_volume_profile(vbp: pd.DataFrame) -> plt.Figure:
    fig,(ax1,ax2) = plt.subplots(1,2,figsize=(12,5))
    _theme(fig,[ax1,ax2])
    ax1.barh(vbp["rate"].astype(str), vbp["total_qty"], color=BLUE, alpha=.7, height=.6)
    ax1.set_title("Volume Profile",fontsize=12); ax1.set_xlabel("Quantity"); ax1.set_ylabel("Price")
    ax1.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v,_:f"{v:,.0f}"))
    sz = (vbp["total_qty"]/vbp["total_qty"].max()*600).clip(15)
    sc = ax2.scatter(vbp["rate"], vbp["trades"], s=sz,
                     c=vbp["total_qty"], cmap="plasma", alpha=.7,
                     edgecolors=GRID, linewidths=.4)
    ax2.set_title("Price vs Trade Count (bubble=volume)",fontsize=12)
    ax2.set_xlabel("Price (Rate)"); ax2.set_ylabel("# Trades")
    cb = plt.colorbar(sc,ax=ax2)
    cb.ax.yaxis.set_tick_params(color=TEXT)
    plt.setp(cb.ax.yaxis.get_ticklabels(),color=TEXT,fontsize=7)
    cb.set_label("Total Qty",color=TEXT,fontsize=8)
    fig.tight_layout(); return fig

def chart_acf(acf_d: dict) -> plt.Figure:
    lags,acf,ci = acf_d["lags"],acf_d["acf"],acf_d["ci"]
    fig,ax = plt.subplots(figsize=(10,4))
    _theme(fig,[ax])
    colors=[GREEN if v>0 else RED for v in acf]
    ax.bar(lags, acf, color=colors, alpha=.75, width=.6, zorder=3)
    ax.axhline( ci, color=AMBER, ls="--", lw=1.2, label=f"+95% CI ±{ci:.3f}", alpha=.85)
    ax.axhline(-ci, color=AMBER, ls="--", lw=1.2, alpha=.85)
    ax.axhline( 0,  color=GRID,  lw=.8)
    ax.set_title("Order Flow Autocorrelation — Trade Sign Persistence",fontsize=12)
    ax.set_xlabel("Lag (trades)"); ax.set_ylabel("ACF")
    ax.set_xlim(.5, max(lags)+.5)
    ax.legend(facecolor=PANEL,edgecolor=GRID,labelcolor=TEXT,fontsize=9)
    fig.tight_layout(); return fig

def chart_trade_size(sz_d: dict) -> plt.Figure:
    fig,(ax1,ax2) = plt.subplots(1,2,figsize=(12,4))
    _theme(fig,[ax1,ax2])
    qty = sz_d["qty"]; thr = sz_d["threshold"]
    n,bins,patches = ax1.hist(qty, bins=min(50,len(qty)//5+1),
                               color=BLUE, alpha=.7, edgecolor=BG, lw=.3)
    for p,left in zip(patches,bins[:-1]):
        if left>=thr:
            p.set_facecolor(RED); p.set_alpha(.85)
    ax1.axvline(sz_d["mean"],   color=AMBER, ls="--", lw=1.5, label=f"Mean {sz_d['mean']:,.0f}")
    ax1.axvline(sz_d["median"], color=GREEN, ls=":",  lw=1.5, label=f"Median {sz_d['median']:,.0f}")
    ax1.axvline(thr,            color=RED,   ls="-",  lw=1.5, label=f"Large≥{thr:,.0f}")
    ax1.set_title("Trade Size (Qty) Distribution",fontsize=12)
    ax1.set_xlabel("Quantity"); ax1.set_ylabel("Frequency")
    ax1.legend(facecolor=PANEL,edgecolor=GRID,labelcolor=TEXT,fontsize=8)
    bp = ax2.boxplot(qty,patch_artist=True,vert=True,
                     boxprops=dict(facecolor=BLUE,color=GRID,alpha=.5),
                     medianprops=dict(color=AMBER,lw=2),
                     whiskerprops=dict(color=TEXT),
                     capprops=dict(color=TEXT),
                     flierprops=dict(marker="o",color=RED,alpha=.4,ms=4))
    ax2.set_title("Qty Box Plot",fontsize=12); ax2.set_ylabel("Quantity")
    ax2.set_xticks([1]); ax2.set_xticklabels(["All Trades"])
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v,_:f"{v:,.0f}"))
    fig.tight_layout(); return fig

def chart_gauge(score: float, color: str) -> plt.Figure:
    fig,ax = plt.subplots(figsize=(4.5,2.8),subplot_kw=dict(aspect="equal"))
    fig.patch.set_facecolor(PANEL); ax.set_facecolor(PANEL); ax.axis("off")
    segs = [(-10,-5,RED),(-5,-2,AMBER),(-2,2,TEXT),(2,5,"#86efac"),(5,10,GREEN)]
    for lo,hi,c in segs:
        t1 = 180-(lo+10)/20*180; t2 = 180-(hi+10)/20*180
        ax.add_patch(mpatches.Wedge((0,0),1.0,t2,t1,
                     width=.28,facecolor=c,alpha=.28,edgecolor=BG,lw=.5))
    ang = np.radians(180-(score+10)/20*180)
    nx,ny = .70*np.cos(ang),.70*np.sin(ang)
    ax.annotate("",xy=(nx,ny),xytext=(0,0),
                arrowprops=dict(arrowstyle="->",color=color,lw=2.5))
    ax.plot(0,0,"o",color=color,ms=5)
    ax.text(0,-.12,f"{score:+.1f}",ha="center",va="center",
            fontsize=20,color=color,fontfamily="monospace",fontweight="bold")
    ax.text(-1.1,0,"−10",ha="center",va="center",fontsize=7,color=RED)
    ax.text( 1.1,0,"+10",ha="center",va="center",fontsize=7,color=GREEN)
    ax.set_xlim(-1.3,1.3); ax.set_ylim(-.35,1.2)
    fig.tight_layout(pad=0); return fig

def chart_symbol_compare(df_all: pd.DataFrame, symbols: list[str]) -> plt.Figure:
    """Bar chart comparing Smart Money Score across multiple symbols."""
    scores = []
    for sym in symbols:
        sub = df_all[df_all["symbol"]==sym]
        if len(sub) < 5:
            scores.append((sym, 0))
            continue
        bd = broker_analysis(sub)
        ad = acf_analysis(sub)
        sd = size_analysis(sub)
        md = detect_metaorders(sub)
        pr = pressure_stats(sub, bd)
        sc = smart_money_score(bd,ad,sd,md,pr)
        scores.append((sym, sc["score"]))
    scores.sort(key=lambda x:x[1], reverse=True)
    syms = [s[0] for s in scores]; vals = [s[1] for s in scores]
    colors = [GREEN if v>=2 else (RED if v<=-2 else TEXT) for v in vals]
    fig,ax = plt.subplots(figsize=(max(8,len(syms)*0.6+2),4))
    _theme(fig,[ax])
    bars = ax.bar(syms, vals, color=colors, alpha=.8, zorder=3)
    ax.axhline(0,color=GRID,lw=.8)
    ax.set_title("Smart Money Score by Symbol",fontsize=12,pad=10)
    ax.set_ylabel("Score"); ax.set_ylim(-11,11)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_x()+bar.get_width()/2,
                v+(0.3 if v>=0 else -0.5),
                f"{v:+.1f}",ha="center",va="bottom",
                fontsize=8,color=TEXT,fontfamily="monospace")
    plt.xticks(rotation=45,ha="right",fontsize=8)
    fig.tight_layout(); return fig

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════

def build_csv_report(symbol, broker_df, meta_df, sz_d, acf_d, sc_d, pr) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["NEPSE FloorSheet Intelligence — Analysis Report"])
    w.writerow([f"Symbol: {symbol}",f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
    w.writerow([])
    w.writerow(["=== SMART MONEY SCORE ==="])
    w.writerow(["Score",sc_d["score"],"Interpretation",
                sc_d["interp"].replace("🟢","").replace("🟡","").replace("⚪","")
                              .replace("🟠","").replace("🔴","").strip()])
    w.writerow(["Confidence",f"{sc_d['conf']}%"])
    for k,v in sc_d["comps"].items():
        w.writerow([f"  {k}",v])
    w.writerow([])
    w.writerow(["=== FLOW SUMMARY ==="])
    w.writerow(["Total Qty",f"{pr['total_qty']:,.0f}","Trades",pr["trade_count"],
                "Buyers",pr["buyers"],"Sellers",pr["sellers"]])
    w.writerow(["Buy Ratio %",f"{pr['buy_ratio']:.1f}"])
    w.writerow([])
    w.writerow(["=== BROKER TABLE (top 30) ==="])
    w.writerow(["Broker","Buy Qty","Sell Qty","Net Qty","Total Qty","Net %"])
    for _,r in broker_df.head(30).iterrows():
        w.writerow([r["broker"],f"{r['buy_qty']:,.0f}",f"{r['sell_qty']:,.0f}",
                    f"{r['net_qty']:,.0f}",f"{r['total_qty']:,.0f}",f"{r['net_pct']:.1f}"])
    w.writerow([])
    w.writerow(["=== METAORDERS ==="])
    if len(meta_df):
        w.writerow(list(meta_df.columns))
        for _,r in meta_df.iterrows():
            w.writerow(list(r.values))
    else:
        w.writerow(["None detected"])
    w.writerow([])
    w.writerow(["=== TRADE SIZE ==="])
    w.writerow(["Mean",f"{sz_d['mean']:,.2f}","Median",f"{sz_d['median']:,.2f}",
                "Std",f"{sz_d['std']:,.2f}","LargeThreshold",f"{sz_d['threshold']:,.2f}",
                "LargeTrades",sz_d["large_count"],"LargeVol",f"{sz_d['large_vol']:,.0f}"])
    w.writerow([])
    w.writerow(["=== ACF ==="])
    w.writerow(["Lag","ACF"])
    for l,v in zip(acf_d["lags"],acf_d["acf"]):
        w.writerow([l,f"{v:.4f}"])
    return buf.getvalue().encode("utf-8")

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — UPLOAD & MAPPING
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🏔️ NEPSE FloorSheet\n**Intelligence System**")
    st.markdown("---")

    uploaded = st.file_uploader("Upload Floorsheet (.xlsx)", type=["xlsx"])

    if uploaded:
        raw_bytes = uploaded.read()
        df_raw, err = load_raw(raw_bytes, uploaded.name)
        if err:
            st.error(f"Could not read file: {err}")
            df_raw = None
        else:
            st.session_state.df_raw = df_raw
            st.success(f"✅ Loaded {len(df_raw):,} rows × {len(df_raw.columns)} columns")

    # ── COLUMN MAPPING UI ───────────────────────────────────────────────
    if st.session_state.df_raw is not None:
        df_raw = st.session_state.df_raw
        cols = list(df_raw.columns)
        opts = ["— not mapped —"] + cols

        st.markdown("---")
        st.markdown("### 🗺️ Column Mapping")
        st.caption("Match your file's columns to the required fields.")

        # Auto-suggest on first load
        if not st.session_state.mapping_confirmed:
            guessed = guess_mapping(cols)
            for k,v in guessed.items():
                if st.session_state[k] is None and v is not None:
                    st.session_state[k] = v

        def _sel(label, key, required=False):
            cur = st.session_state.get(key)
            idx = cols.index(cur)+1 if cur in cols else 0
            star = " *" if required else ""
            ch = st.selectbox(f"{label}{star}", opts, index=idx, key=f"_sel_{key}")
            st.session_state[key] = None if ch == "— not mapped —" else ch

        _sel("SN (Serial)",          "col_sn")
        _sel("Contract No",          "col_contract")
        _sel("Stock Symbol",         "col_symbol")
        _sel("Buyer Broker No",      "col_buyer",   required=True)
        _sel("Seller Broker No",     "col_seller",  required=True)
        _sel("Quantity",             "col_qty",     required=True)
        _sel("Rate (Price)",         "col_rate",    required=True)
        _sel("Amount (Trade Value)", "col_amount")

        st.markdown("")
        if st.button("✅ Confirm Mapping & Analyze"):
            st.session_state.mapping_confirmed = True
            st.rerun()

        if st.session_state.mapping_confirmed:
            st.success("Mapping confirmed!")
            if st.button("🔄 Re-map Columns"):
                st.session_state.mapping_confirmed = False
                st.rerun()

    # ── ANALYSIS PARAMS ─────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### ⚙️ Parameters")
    top_n       = st.slider("Top N Brokers", 5, 30, 10)
    min_run     = st.slider("Min Metaorder Run", 2, 10, 3)
    max_lags    = st.slider("ACF Max Lags", 5, 50, 20)

# ══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("# 🏔️ NEPSE FloorSheet Intelligence")
st.markdown("##### Institutional Order Flow · Smart Money Detection · NEPSE Market Analysis")
st.markdown("---")

# ── LANDING ──────────────────────────────────────────────────────────────────
if st.session_state.df_raw is None:
    c1,c2,c3 = st.columns(3)
    tiles = [
        ("📋 Column Mapping","Auto-detects NEPSE floorsheet columns and lets you fix any mismatches via dropdown selectors — works with any broker export format."),
        ("📊 Symbol-Wise Analysis","Filter and deep-dive into any individual stock. Compare Smart Money Scores across all symbols in a single chart."),
        ("🧠 Smart Money Score","−10 to +10 composite signal built from broker net flow, ACF persistence, trade size anomalies, and metaorder directional bias."),
    ]
    for col,(title,desc) in zip([c1,c2,c3],tiles):
        with col:
            st.markdown(f"<div class='card'><b>{title}</b><br><br><span style='color:#64748b;font-size:.87rem'>{desc}</span></div>",
                        unsafe_allow_html=True)
    st.info("👈 Upload your NEPSE floorsheet (.xlsx) in the sidebar to begin.")
    st.markdown("### Expected Column Names")
    demo = pd.DataFrame({"SN":[1,2],"Contract No":["C001","C002"],
                         "Stock Symbol":["NABIL","SCB"],
                         "Buyer Broker No":[12,34],"Seller Broker No":[45,67],
                         "Quantity":[500,1200],"Rate":[1250,823],"Amount":[625000,987600]})
    st.dataframe(demo,use_container_width=True)
    st.caption("Column names don't need to match exactly — the mapping UI handles variations.")
    st.stop()

# ── MAPPING NOT CONFIRMED YET ────────────────────────────────────────────────
if not st.session_state.mapping_confirmed:
    st.info("📋 Please complete the **Column Mapping** in the sidebar, then click **Confirm Mapping & Analyze**.")
    st.markdown("### Raw File Preview")
    st.dataframe(st.session_state.df_raw.head(20), use_container_width=True)
    st.stop()

# ── PREPARE DATA ─────────────────────────────────────────────────────────────

mapping = {k: st.session_state[k] for k in MAPPING_KEYS}
with st.spinner("🔧 Preparing data..."):
    df_all, prep_warns = prepare_df(st.session_state.df_raw, mapping)

if prep_warns:
    for w in prep_warns:
        st.warning(w) if not w.startswith("❌") else st.error(w)

if df_all is None:
    st.stop()

# ── SYMBOL SELECTOR ──────────────────────────────────────────────────────────

all_symbols = sorted(df_all["symbol"].unique().tolist())
has_symbols = len(all_symbols) > 1

st.markdown('<p class="sec-label">◆ Stock Selection</p>', unsafe_allow_html=True)

if has_symbols:
    sel_col1, sel_col2 = st.columns([2,3])
    with sel_col1:
        chosen_symbol = st.selectbox("Select Stock Symbol for Analysis",
                                     options=all_symbols, index=0)
    with sel_col2:
        st.markdown(f"""
        <div class='card' style='padding:12px 16px'>
        <span style='color:#64748b;font-size:.8rem'>ANALYZING</span><br>
        <span style='font-family:IBM Plex Mono,monospace;font-size:1.4rem;color:#38bdf8'>{chosen_symbol}</span>
        &nbsp;&nbsp;<span style='color:#64748b;font-size:.82rem'>
        {len(df_all[df_all['symbol']==chosen_symbol]):,} trades &nbsp;|&nbsp;
        {len(all_symbols)} symbols in file
        </span>
        </div>""", unsafe_allow_html=True)
else:
    chosen_symbol = all_symbols[0]
    st.info(f"Single symbol in file: **{chosen_symbol}**")

df = df_all[df_all["symbol"] == chosen_symbol].reset_index(drop=True)

if len(df) < 5:
    st.error(f"Only {len(df)} trades for {chosen_symbol} — insufficient data. Try a different symbol.")
    st.stop()

# ── RUN ANALYSIS ─────────────────────────────────────────────────────────────

with st.spinner(f"⚙️ Analyzing {chosen_symbol}..."):
    broker_df  = broker_analysis(df)
    acf_d      = acf_analysis(df, max_lags)
    sz_d       = size_analysis(df)
    meta_df    = detect_metaorders(df, min_run)
    vbp_df     = volume_profile(df)
    pr         = pressure_stats(df, broker_df)
    sc_d       = smart_money_score(broker_df, acf_d, sz_d, meta_df, pr)

# ── HEADER METRICS ───────────────────────────────────────────────────────────

st.markdown('<p class="sec-label">◆ Dashboard Overview</p>', unsafe_allow_html=True)
m1,m2,m3,m4,m5 = st.columns(5)
m1.metric("Total Trades",    f"{pr['trade_count']:,}")
m2.metric("Total Quantity",  f"{pr['total_qty']:,.0f}")
m3.metric("Active Brokers",  f"{max(pr['buyers'],pr['sellers']):,}")
m4.metric("Buy Pressure",    f"{pr['buy_ratio']:.1f}%",
          delta=f"{pr['buy_ratio']-50:+.1f}% vs neutral")
m5.metric("Large Trades",    f"{sz_d['large_count']:,}",
          delta=f"{sz_d['large_count']/pr['trade_count']*100:.1f}% of flow")

st.markdown("<br>",unsafe_allow_html=True)

# ── SMART MONEY SCORE ─────────────────────────────────────────────────────────

st.markdown('<p class="sec-label">◆ Smart Money Score</p>', unsafe_allow_html=True)
ga, gb, gc = st.columns([1,1.1,1.6])

with ga:
    st.pyplot(chart_gauge(sc_d["score"], sc_d["color"]), use_container_width=True)

with gb:
    c = sc_d["color"]
    st.markdown(f"""
    <div class='score-card'>
        <div style='font-size:.72rem;letter-spacing:2px;color:#334155;text-transform:uppercase'>Smart Money Score</div>
        <div class='score-num' style='color:{c}'>{sc_d['score']:+.1f}</div>
        <div style='font-size:.95rem;color:{c};margin:8px 0'>{sc_d['interp']}</div>
        <div style='font-size:.78rem;color:#475569'>Confidence: <b style='color:{c}'>{sc_d['conf']}%</b></div>
    </div>""", unsafe_allow_html=True)

with gc:
    st.markdown("**Score Components**")
    comp_rows = [{"Component":k,"Points":v,
                  "Signal":"▲ Bullish" if v>0 else ("▼ Bearish" if v<0 else "— Neutral")}
                 for k,v in sc_d["comps"].items()]
    st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)

# ── TABS ─────────────────────────────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
tabs = st.tabs(["🏦 Brokers","📈 Order Flow ACF","🔍 Metaorders",
                "📐 Trade Size","📊 Volume Profile",
                "📉 Multi-Symbol","💡 Insights","📁 Data & Export"])

# ── TAB 1: BROKERS ────────────────────────────────────────────────────────────
with tabs[0]:
    st.markdown('<p class="sec-label">◆ Broker-Level Activity</p>', unsafe_allow_html=True)
    st.pyplot(chart_broker(broker_df, top_n), use_container_width=True)
    st.pyplot(chart_net_position(broker_df, top_n*2), use_container_width=True)

    ca, cb = st.columns(2)
    with ca:
        st.markdown("#### 🟢 Top Accumulators (Net Buyers)")
        acc = broker_df[broker_df["net_qty"]>0].nlargest(10,"net_qty")[
              ["broker","buy_qty","sell_qty","net_qty","net_pct"]].copy()
        acc.columns=["Broker","Buy Qty","Sell Qty","Net Qty","Net %"]
        for col in ["Buy Qty","Sell Qty","Net Qty"]:
            acc[col] = acc[col].map("{:,.0f}".format)
        acc["Net %"] = acc["Net %"].map("{:.1f}%".format)
        st.dataframe(acc, use_container_width=True, hide_index=True)

    with cb:
        st.markdown("#### 🔴 Top Distributors (Net Sellers)")
        dist = broker_df[broker_df["net_qty"]<0].nsmallest(10,"net_qty")[
               ["broker","buy_qty","sell_qty","net_qty","net_pct"]].copy()
        dist.columns=["Broker","Buy Qty","Sell Qty","Net Qty","Net %"]
        for col in ["Buy Qty","Sell Qty","Net Qty"]:
            dist[col] = dist[col].map("{:,.0f}".format)
        dist["Net %"] = dist["Net %"].map("{:.1f}%".format)
        st.dataframe(dist, use_container_width=True, hide_index=True)

    st.markdown("#### Full Broker Table")
    disp_b = broker_df.copy()
    for c in ["buy_qty","sell_qty","net_qty","total_qty"]:
        disp_b[c] = disp_b[c].map("{:,.0f}".format)
    disp_b["net_pct"] = disp_b["net_pct"].map("{:.1f}%".format)
    disp_b.columns = ["Broker","Buy Qty","Sell Qty","Net Qty","Total Qty","Net %"]
    st.dataframe(disp_b, use_container_width=True, hide_index=True)

# ── TAB 2: ACF ────────────────────────────────────────────────────────────────
with tabs[1]:
    st.markdown('<p class="sec-label">◆ Order Flow Autocorrelation</p>', unsafe_allow_html=True)
    st.pyplot(chart_acf(acf_d), use_container_width=True)
    a1,a2,a3 = st.columns(3)
    a1.metric("Mean ACF (5 lags)", f"{acf_d['mean5']:.4f}")
    a2.metric("Significant Lags",   str(len(acf_d["sig_lags"])))
    a3.metric("Pos. Persist (≤5)",  str(acf_d["pos5"]))

    if acf_d["pos5"] >= 3:
        st.markdown("<div class='flag'>⚠️ <b>Strong buy-side autocorrelation</b> — consecutive buy trades indicate institutional order-splitting or momentum execution algorithm.</div>",
                    unsafe_allow_html=True)
    elif acf_d["mean5"] < -0.05:
        st.markdown("<div class='insight'>📊 <b>Negative autocorrelation</b> — trade direction alternates, consistent with market-making / two-sided liquidity provision.</div>",
                    unsafe_allow_html=True)
    else:
        st.markdown("<div class='insight'>📊 <b>Near-zero ACF</b> — no persistent directional bias in order flow. Normal retail two-sided trading.</div>",
                    unsafe_allow_html=True)
    if acf_d["sig_lags"]:
        st.info(f"Statistically significant lags (beyond 95% CI): {acf_d['sig_lags']}")

# ── TAB 3: METAORDERS ────────────────────────────────────────────────────────
with tabs[2]:
    st.markdown('<p class="sec-label">◆ Metaorder Detection — Possible Institutional Order Splitting</p>', unsafe_allow_html=True)

    if len(meta_df) == 0:
        st.info(f"No metaorders found with run ≥ {min_run}. Try lowering the minimum run in the sidebar.")
    else:
        high = meta_df[meta_df["suspicion"]=="HIGH"]
        mod  = meta_df[meta_df["suspicion"]=="MODERATE"]
        mc1,mc2,mc3 = st.columns(3)
        mc1.metric("Total Metaorders", len(meta_df))
        mc2.metric("HIGH Suspicion",   len(high))
        mc3.metric("MODERATE",         len(mod))

        if len(high) > 0:
            brokers_hi = high["broker"].unique()
            st.markdown(f"<div class='flag'>🚨 <b>HIGH-suspicion institutional activity</b> from brokers: <b>{list(brokers_hi)}</b> — consistent direction, similar sizes, minimal price drift. Classic algorithmic execution.</div>",
                        unsafe_allow_html=True)

        disp_m = meta_df.copy()
        disp_m["total_qty"] = disp_m["total_qty"].map("{:,.0f}".format)
        disp_m["avg_rate"]  = disp_m["avg_rate"].map("{:.2f}".format)
        disp_m["price_cv"]  = disp_m["price_cv"].map("{:.2f}%".format)
        disp_m.columns = ["Broker","Direction","Run Length","Total Qty","Avg Rate","Price CV%","Suspicion"]
        st.dataframe(disp_m, use_container_width=True, hide_index=True)

        st.markdown("#### Metaorder Summary by Broker")
        msumm = meta_df.groupby(["broker","direction"]).agg(
            Runs=("run_len","count"),
            Total_Qty=("total_qty","sum"),
            Avg_Run=("run_len","mean")).reset_index().sort_values("Total_Qty",ascending=False)
        st.dataframe(msumm, use_container_width=True, hide_index=True)

# ── TAB 4: TRADE SIZE ─────────────────────────────────────────────────────────
with tabs[3]:
    st.markdown('<p class="sec-label">◆ Trade Size Analysis</p>', unsafe_allow_html=True)
    st.pyplot(chart_trade_size(sz_d), use_container_width=True)
    ts1,ts2,ts3,ts4 = st.columns(4)
    ts1.metric("Mean Qty",      f"{sz_d['mean']:,.0f}")
    ts2.metric("Median Qty",    f"{sz_d['median']:,.0f}")
    ts3.metric("Std Dev",       f"{sz_d['std']:,.0f}")
    ts4.metric("Large Threshold",f"{sz_d['threshold']:,.0f}")

    if len(sz_d["large"]) > 0:
        lvp = sz_d["large_vol"]/pr["total_qty"]*100
        cls = "flag" if lvp>20 else "insight"
        st.markdown(f"<div class='{cls}'>{'⚠️' if lvp>20 else 'ℹ️'} Large trades = <b>{sz_d['large_count']}</b> trades, <b>{lvp:.1f}%</b> of total quantity {'— possible block/institutional activity.' if lvp>20 else '— within normal range.'}</div>",
                    unsafe_allow_html=True)
        st.markdown("#### Large Trade Records")
        show_cols = [c for c in ["sn","buyer","seller","rate","qty"] if c in sz_d["large"].columns]
        lt = sz_d["large"][show_cols].copy()
        if "qty" in lt.columns: lt["qty"] = lt["qty"].map("{:,.0f}".format)
        st.dataframe(lt, use_container_width=True, hide_index=True)
    else:
        st.info("No large trade outliers detected.")

# ── TAB 5: VOLUME PROFILE ─────────────────────────────────────────────────────
with tabs[4]:
    st.markdown('<p class="sec-label">◆ Volume Profile</p>', unsafe_allow_html=True)
    st.pyplot(chart_volume_profile(vbp_df), use_container_width=True)

    poc   = vbp_df.loc[vbp_df["total_qty"].idxmax()]
    cum   = vbp_df["cum_qty"].max()
    vah_r = vbp_df[vbp_df["cum_qty"] >= cum*0.70].iloc[0]
    val_r = vbp_df[vbp_df["cum_qty"] >= cum*0.30].iloc[0]
    vp1,vp2,vp3 = st.columns(3)
    vp1.metric("Point of Control (POC)", f"{poc['rate']:,.2f}")
    vp2.metric("Value Area High (VAH)",  f"{vah_r['rate']:,.2f}")
    vp3.metric("Value Area Low (VAL)",   f"{val_r['rate']:,.2f}")

    st.markdown("#### Top High-Volume Nodes")
    for _,row in vbp_df.nlargest(3,"total_qty").iterrows():
        pct = row["total_qty"]/vbp_df["total_qty"].sum()*100
        st.markdown(f"<div class='insight'>🎯 Price <b>{row['rate']:,.2f}</b> — Qty: <b>{row['total_qty']:,.0f}</b> ({pct:.1f}%) across <b>{row['trades']:,}</b> trades — significant volume node.</div>",
                    unsafe_allow_html=True)

    st.markdown("#### Price-Level Volume Table")
    disp_v = vbp_df.copy()
    disp_v["total_qty"] = disp_v["total_qty"].map("{:,.0f}".format)
    disp_v["pct"]       = disp_v["pct"].map("{:.2f}%".format)
    disp_v.columns      = ["Rate","Total Qty","Trades","Pct%","Cum Qty"]
    st.dataframe(disp_v, use_container_width=True, hide_index=True)

# ── TAB 6: MULTI-SYMBOL ───────────────────────────────────────────────────────
with tabs[5]:
    st.markdown('<p class="sec-label">◆ Multi-Symbol Comparison</p>', unsafe_allow_html=True)

    if not has_symbols:
        st.info("Only one symbol in the loaded file. Upload a full-day floorsheet to enable multi-symbol comparison.")
    else:
        min_trades = st.slider("Min trades per symbol", 5, 50, 10)
        eligible   = [s for s in all_symbols
                      if len(df_all[df_all["symbol"]==s]) >= min_trades]
        max_show   = st.slider("Max symbols to show", 5, min(50,len(eligible)), min(20,len(eligible)))
        show_syms  = eligible[:max_show]

        if len(show_syms) < 2:
            st.warning("Not enough eligible symbols. Lower the minimum trades threshold.")
        else:
            with st.spinner(f"Computing scores for {len(show_syms)} symbols..."):
                fig_cmp = chart_symbol_compare(df_all, show_syms)
            st.pyplot(fig_cmp, use_container_width=True)

            # Build comparison table
            rows = []
            for sym in show_syms:
                sub = df_all[df_all["symbol"]==sym]
                bd  = broker_analysis(sub)
                ad  = acf_analysis(sub)
                sd  = size_analysis(sub)
                md  = detect_metaorders(sub)
                prr = pressure_stats(sub, bd)
                sc  = smart_money_score(bd,ad,sd,md,prr)
                rows.append({
                    "Symbol":   sym,
                    "Trades":   len(sub),
                    "Score":    sc["score"],
                    "Signal":   sc["interp"].split("—")[1].strip() if "—" in sc["interp"] else sc["interp"],
                    "Buy%":     f"{prr['buy_ratio']:.1f}",
                    "Brokers":  max(prr["buyers"],prr["sellers"]),
                })
            cmp_df = pd.DataFrame(rows).sort_values("Score",ascending=False).reset_index(drop=True)
            st.dataframe(cmp_df, use_container_width=True, hide_index=True)

            cmp_csv = cmp_df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download Comparison CSV", cmp_csv,
                               file_name="symbol_comparison.csv", mime="text/csv")

# ── TAB 7: INSIGHTS ───────────────────────────────────────────────────────────
with tabs[6]:
    st.markdown('<p class="sec-label">◆ Automated Intelligence Insights</p>', unsafe_allow_html=True)

    st.markdown(f"### 🧠 Smart Money Assessment — {chosen_symbol}")
    c = sc_d["color"]
    st.markdown(f"<div class='score-card' style='text-align:left'><div style='font-size:1.05rem;color:{c};font-weight:600'>{sc_d['interp']}</div><div style='font-size:.82rem;color:#475569;margin-top:8px'>Score: <b style='color:{c}'>{sc_d['score']:+.2f}</b> / 10 &nbsp;|&nbsp; Confidence: <b>{sc_d['conf']}%</b></div></div>",
                unsafe_allow_html=True)

    st.markdown("### 🏦 Key Broker Findings")
    acc_all = broker_df[broker_df["net_qty"]>0]
    dis_all = broker_df[broker_df["net_qty"]<0]

    if len(acc_all)>0:
        top_a = acc_all.iloc[0]
        st.markdown(f"<div class='insight'>🟢 <b>Primary Accumulator — Broker {top_a['broker']}</b><br>Net buy: <b>{top_a['net_qty']:,.0f}</b> shares ({top_a['net_pct']:.1f}% directional bias). Buy: {top_a['buy_qty']:,.0f} | Sell: {top_a['sell_qty']:,.0f}</div>",
                    unsafe_allow_html=True)
    if len(dis_all)>0:
        top_d = dis_all.iloc[-1]
        st.markdown(f"<div class='insight'>🔴 <b>Primary Distributor — Broker {top_d['broker']}</b><br>Net sell: <b>{abs(top_d['net_qty']):,.0f}</b> shares ({abs(top_d['net_pct']):.1f}% directional bias). Buy: {top_d['buy_qty']:,.0f} | Sell: {top_d['sell_qty']:,.0f}</div>",
                    unsafe_allow_html=True)

    st.markdown("### 🔍 Institutional Activity Flags")
    flags = []

    if acf_d["pos5"] >= 3:
        flags.append(("⚠️ Order Flow Persistence",
                      f"Positive ACF at {acf_d['pos5']}/5 early lags (mean={acf_d['mean5']:.3f}). "
                      "Consistent with institutional order-splitting execution.",False))
    lvp = sz_d["large_vol"]/pr["total_qty"]*100
    if lvp > 15:
        flags.append(("⚠️ Abnormal Large Trade Volume",
                      f"{sz_d['large_count']} trades above size threshold ({sz_d['threshold']:,.0f}), "
                      f"comprising {lvp:.1f}% of total volume.",False))
    high_meta = meta_df[meta_df["suspicion"]=="HIGH"] if len(meta_df)>0 else pd.DataFrame()
    if len(high_meta)>0:
        flags.append(("🚨 HIGH-Suspicion Metaorders",
                      f"Brokers {list(high_meta['broker'].unique())} flagged for institutional order-splitting "
                      f"across {len(high_meta)} sequences.",True))
    if len(broker_df)>0:
        top_share = broker_df.iloc[0]["total_qty"]/broker_df["total_qty"].sum()*100
        if top_share>25:
            flags.append(("⚠️ Broker Concentration",
                          f"Broker {broker_df.iloc[0]['broker']} controls {top_share:.1f}% of volume — highly concentrated.",False))
    if pr["buy_ratio"]>70:
        flags.append(("📈 Strong Buy Imbalance",f"Buy ratio {pr['buy_ratio']:.1f}% — significantly bullish.",False))
    elif pr["buy_ratio"]<30:
        flags.append(("📉 Strong Sell Imbalance",f"Buy ratio {pr['buy_ratio']:.1f}% — significantly bearish.",False))

    if flags:
        for title,desc,critical in flags:
            cls = "flag" if critical else "insight"
            st.markdown(f"<div class='{cls}'><b>{title}</b><br>{desc}</div>",
                        unsafe_allow_html=True)
    else:
        st.markdown("<div class='success'>✅ No significant institutional activity flags. Market appears normally balanced.</div>",
                    unsafe_allow_html=True)

    st.markdown("### 📊 Market Bias Summary")
    bias = [
        ("Buy/Sell Pressure",   f"{pr['buy_ratio']:.1f}% buy-side",   pr["buy_ratio"]>55, pr["buy_ratio"]<45),
        ("Broker Net Flow",     f"{len(acc_all)} accumulators vs {len(dis_all)} distributors",
                                len(acc_all)>len(dis_all),len(dis_all)>len(acc_all)),
        ("Order Flow Momentum", f"ACF mean(5)={acf_d['mean5']:.3f}",
                                acf_d["mean5"]>0.05,acf_d["mean5"]<-0.05),
        ("Metaorder Direction", f"Buy: {len(meta_df[meta_df['direction']=='BUY']) if len(meta_df)>0 else 0} | "
                                 f"Sell: {len(meta_df[meta_df['direction']=='SELL']) if len(meta_df)>0 else 0}",
                                (len(meta_df[meta_df['direction']=='BUY']) if len(meta_df)>0 else 0) >
                                (len(meta_df[meta_df['direction']=='SELL']) if len(meta_df)>0 else 0),
                                False),
    ]
    for label,value,is_bull,is_bear in bias:
        icon  = "🟢" if is_bull else ("🔴" if is_bear else "⚪")
        color = GREEN  if is_bull else (RED if is_bear else "#475569")
        st.markdown(f"""<div style='display:flex;justify-content:space-between;padding:10px 16px;
                    background:#0b1220;border:1px solid #1a2535;border-radius:6px;margin:4px 0'>
                    <span>{icon} <b style='color:{color}'>{label}</b></span>
                    <span style='color:#475569;font-family:monospace;font-size:.83rem'>{value}</span>
                    </div>""", unsafe_allow_html=True)

    st.markdown("### 📥 Export Report")
    csv_bytes = build_csv_report(chosen_symbol, broker_df, meta_df, sz_d, acf_d, sc_d, pr)
    st.download_button("⬇️ Download Full Analysis Report (CSV)", csv_bytes,
                       file_name=f"NEPSE_{chosen_symbol}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                       mime="text/csv")

# ── TAB 8: DATA & EXPORT ─────────────────────────────────────────────────────
with tabs[7]:
    st.markdown('<p class="sec-label">◆ Data Preview & Export</p>', unsafe_allow_html=True)

    st.markdown(f"**{len(df):,} trades** loaded for **{chosen_symbol}**")
    st.dataframe(df.drop(columns=["trade_sign"], errors="ignore"), use_container_width=True)

    dl1,dl2,dl3 = st.columns(3)
    with dl1:
        st.download_button("⬇️ Cleaned Trades (CSV)",
                           df.drop(columns=["trade_sign"],errors="ignore")
                             .to_csv(index=False).encode("utf-8"),
                           f"{chosen_symbol}_trades.csv","text/csv")
    with dl2:
        st.download_button("⬇️ Broker Summary (CSV)",
                           broker_df.to_csv(index=False).encode("utf-8"),
                           f"{chosen_symbol}_brokers.csv","text/csv")
    with dl3:
        if len(meta_df)>0:
            st.download_button("⬇️ Metaorders (CSV)",
                               meta_df.to_csv(index=False).encode("utf-8"),
                               f"{chosen_symbol}_metaorders.csv","text/csv")
        else:
            st.button("⬇️ Metaorders (CSV)", disabled=True)

    st.markdown("#### Active Column Mapping")
    map_disp = pd.DataFrame([
        {"Field": k.replace("col_","").upper(), "Mapped To": v or "— not mapped —"}
        for k,v in mapping.items()
    ])
    st.dataframe(map_disp, use_container_width=True, hide_index=True)

# ── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""<div style='text-align:center;color:#1e3a5f;font-size:.76rem;padding:10px 0;font-family:monospace'>
NEPSE FloorSheet Intelligence &nbsp;|&nbsp; For research & analytical purposes only &nbsp;|&nbsp; Not financial advice
</div>""", unsafe_allow_html=True)
