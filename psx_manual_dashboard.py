"""
Enhanced PSX Manual Portfolio Dashboard with Robust Price Handling
================================================================

This version improves on the previous manual portfolio dashboard by
adding robust detection of the `Current Price` column and handling
non‑numeric or missing values gracefully. The app looks for a column
named `Current Price` (case insensitive, with or without spaces or
underscores) and warns the user if no such column is found. When
prices are missing or non‑numeric, the corresponding P/L and market
value are still computed (using zero) and the missing prices are
displayed as “N/A.”

Key Features
------------
* Automatically detects variations of the `Current Price` column name
  (e.g. `Current Price`, `current_price`, `currentprice`).
* Warns when no valid current price column is present.
* Displays `N/A` for positions with missing or unparseable current prices.
* Keeps all features from the previous version: login prompt, upload
  history, sector analysis and downloadable report.

Run with:

```
streamlit run psx_portfolio_manual_dashboard_v2.py
```

"""

import io
import os
from datetime import datetime
import pandas as pd
import streamlit as st

UPLOAD_HISTORY_DIR = "uploaded_history"


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Find a column in df matching any of the candidate names (case insensitive)."""
    lower_cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "")
        for col, orig_col in lower_cols.items():
            # remove spaces and underscores from df columns for comparison
            stripped = orig_col.lower().replace(" ", "").replace("_", "")
            if stripped == key:
                return orig_col
    return None


def load_manual_portfolio(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="My_Stocks")
    # Required columns
    required = {"symbol", "quantity", "buy price"}
    df.columns = [c.strip() for c in df.columns]
    lower_map = {c.lower(): c for c in df.columns}
    missing = [col for col in required if col not in lower_map]
    if missing:
        raise ValueError(
            "Sheet 'My_Stocks' must contain columns: "
            "Symbol, Quantity, Buy Price and a current price column"
        )
    # Determine current price column
    curr_col = find_column(df, ["current price", "current_price", "currentprice"])
    if curr_col is None:
        st.warning("No 'Current Price' column found. Prices will be treated as missing.")
        df["current_price"] = pd.NA
    else:
        df["current_price"] = pd.to_numeric(df[curr_col], errors="coerce")
    pf = pd.DataFrame()
    pf["symbol"] = df[lower_map.get("symbol", "Symbol")].astype(str).str.upper().str.strip()
    pf["quantity"] = pd.to_numeric(df[lower_map.get("quantity", "Quantity")], errors="coerce").fillna(0)
    pf["buy_price"] = pd.to_numeric(df[lower_map.get("buy price", "Buy Price")], errors="coerce").fillna(0.0)
    pf["current_price"] = df["current_price"]
    # Sector column detection
    sector_col = find_column(df, ["sector"])
    if sector_col:
        pf["sector"] = df[sector_col].astype(str).str.strip()
    else:
        pf["sector"] = "Unknown"
    return pf


def compute_portfolio(df: pd.DataFrame) -> pd.DataFrame:
    pf = df.copy()
    pf["cost"] = pf["quantity"] * pf["buy_price"]
    # Use 0 for market value when current price is missing
    pf["market_value"] = pf["quantity"] * pf["current_price"].fillna(0.0)
    pf["pnl"] = pf["market_value"] - pf["cost"]
    pf["pnl_pct"] = pf.apply(
        lambda row: (row["pnl"] / row["cost"] * 100) if row["cost"] != 0 else 0.0,
        axis=1,
    )
    return pf


def compute_sector_summary(pf: pd.DataFrame) -> pd.DataFrame:
    summary = (
        pf.groupby("sector")[ ["cost", "market_value", "pnl"] ]
        .sum()
        .reset_index()
    )
    summary["pnl_pct"] = summary.apply(
        lambda row: (row["pnl"] / row["cost"] * 100) if row["cost"] != 0 else 0.0,
        axis=1,
    )
    return summary


def save_upload(file_buffer: bytes, original_name: str) -> str:
    os.makedirs(UPLOAD_HISTORY_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(original_name)
    safe_base = base.replace(" ", "_")
    filename = f"{safe_base}_{timestamp}{ext}"
    filepath = os.path.join(UPLOAD_HISTORY_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(file_buffer)
    return filepath


# UI
st.set_page_config(page_title="PSX Manual Portfolio Dashboard", page_icon="📊", layout="wide")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.title("Login")
    st.write(
        "Enter your email to access the portfolio dashboard (no external authentication)."
    )
    email_input = st.text_input("Email")
    if st.button("Login"):
        if email_input:
            st.session_state["user"] = email_input
            st.session_state["logged_in"] = True
    st.stop()

st.title("📊 PSX Portfolio Dashboard (Manual Prices)")
st.caption(
    "Upload an Excel file with a 'My_Stocks' sheet containing Symbol, "
    "Quantity, Buy Price, a Current Price column, and an optional Sector column."
)

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel portfolio file", type=["xlsx"], help="Upload your manual PSX portfolio."
)

st.sidebar.subheader("Upload History")
if os.path.exists(UPLOAD_HISTORY_DIR):
    for hfile in sorted(os.listdir(UPLOAD_HISTORY_DIR), reverse=True):
        path = os.path.join(UPLOAD_HISTORY_DIR, hfile)
        with open(path, "rb") as f:
            st.sidebar.download_button(hfile, data=f.read(), file_name=hfile)
else:
    st.sidebar.write("No uploads yet.")

with st.expander("Sample structure of 'My_Stocks' sheet"):
    st.table(
        pd.DataFrame(
            {
                "Symbol": ["EFERT", "HUBC", "SYS"],
                "Quantity": [100, 250, 50],
                "Buy Price": [90.50, 110.00, 420.00],
                "Current Price": [95.00, 118.25, 410.00],
                "Sector": ["Fertilizer", "Power", "Technology"],
            }
        )
    )

if not uploaded_file:
    st.info("⬆️ Upload your Excel file to see the analysis dashboard.")
    st.stop()

# Save file to history
save_upload(uploaded_file.getbuffer(), uploaded_file.name)

try:
    portfolio_df = load_manual_portfolio(uploaded_file)
except Exception as e:
    st.error(f"Error reading portfolio: {e}")
    st.stop()

pf = compute_portfolio(portfolio_df)
sector_summary = compute_sector_summary(pf)

# Totals
sum_cost = pf["cost"].sum()
sum_value = pf["market_value"].sum()
sum_pnl = pf["pnl"].sum()
sum_pnl_pct = (sum_pnl / sum_cost * 100) if sum_cost else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Cost (PKR)", f"{sum_cost:,.0f}")
k2.metric("Market Value (PKR)", f"{sum_value:,.0f}")
k3.metric("Total P/L (PKR)", f"{sum_pnl:,.0f}")
k4.metric("Total P/L (%)", f"{sum_pnl_pct:,.2f}%")

st.caption(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Winners and losers
st.subheader("Performance Summary")
win = pf[pf["pnl"] > 0]
lose = pf[pf["pnl"] < 0]

c_win, c_lose = st.columns(2)

with c_win:
    st.markdown("✅ **Winners (Positive P/L)**")
    if win.empty:
        st.write("No winning positions.")
    else:
        w_df = win[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        w_df["pnl"] = w_df["pnl"].map(lambda x: f"{x:,.0f}")
        w_df["pnl_pct"] = w_df["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        w_df["market_value"] = w_df["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(w_df, use_container_width=True)

with c_lose:
    st.markdown("❌ **Losers (Negative P/L)**")
    if lose.empty:
        st.write("No losing positions.")
    else:
        l_df = lose[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        l_df["pnl"] = l_df["pnl"].map(lambda x: f"{x:,.0f}")
        l_df["pnl_pct"] = l_df["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        l_df["market_value"] = l_df["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(l_df, use_container_width=True)

# Top movers
st.subheader("Top Movers")
if not pf.empty:
    top_gain = pf.sort_values("pnl_pct", ascending=False).head(5)
    top_lose = pf.sort_values("pnl_pct").head(5)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("🚀 **Top Gainers (by % P/L)**")
        tg = top_gain[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        tg["pnl"] = tg["pnl"].map(lambda x: f"{x:,.0f}")
        tg["pnl_pct"] = tg["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        tg["market_value"] = tg["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(tg, use_container_width=True)
    with c2:
        st.markdown("📉 **Top Losers (by % P/L)**")
        tl = top_lose[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        tl["pnl"] = tl["pnl"].map(lambda x: f"{x:,.0f}")
        tl["pnl_pct"] = tl["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        tl["market_value"] = tl["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(tl, use_container_width=True)

# Full details
st.subheader("Full Holdings Detail")

view_df = pf.copy()
view_df["buy_price"] = view_df["buy_price"].map(lambda x: f"{x:,.2f}")
# For current price: show 'N/A' when missing
view_df["current_price"] = view_df["current_price"].apply(
    lambda x: f"{x:,.2f}" if pd.notna(x) else "N/A"
)
view_df["cost"] = view_df["cost"].map(lambda x: f"{x:,.0f}")
view_df["market_value"] = view_df["market_value"].map(lambda x: f"{x:,.0f}")
view_df["pnl"] = view_df["pnl"].map(lambda x: f"{x:,.0f}")
view_df["pnl_pct"] = view_df["pnl_pct"].map(lambda x: f"{x:,.2f}%")

st.dataframe(view_df, use_container_width=True)

# Sector summary
st.subheader("Sector Summary")
st.dataframe(sector_summary, use_container_width=True)

# Sector charts
sec_col1, sec_col2 = st.columns(2)
with sec_col1:
    st.subheader("Allocation by Sector (Market Value)")
    if not sector_summary.empty:
        st.bar_chart(sector_summary.set_index("sector")["market_value"])
with sec_col2:
    st.subheader("P/L by Sector")
    if not sector_summary.empty:
        st.bar_chart(sector_summary.set_index("sector")["pnl"])

# Symbol charts
sym_col1, sym_col2 = st.columns(2)
with sym_col1:
    st.subheader("Allocation by Market Value (Symbol)")
    st.bar_chart(pf[["symbol", "market_value"]].set_index("symbol"))
with sym_col2:
    st.subheader("P/L by Symbol")
    st.bar_chart(pf[["symbol", "pnl"]].set_index("symbol"))

# Download
st.subheader("Download Analysis")
out = io.BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
    pf.to_excel(writer, index=False, sheet_name="Portfolio Details")
    sector_summary.to_excel(writer, index=False, sheet_name="Sector Summary")
out.seek(0)
st.download_button(
    label="📥 Download Portfolio Analysis Excel",
    data=out,
    file_name="psx_portfolio_manual_analysis_v2.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
