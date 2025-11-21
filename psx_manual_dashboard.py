"""
PSX Portfolio Dashboard (Manual Prices) with Sector Analysis and Upload History
=============================================================================

This Streamlit application is an enhanced version of a manual PSX portfolio
dashboard. It reads an Excel file (uploaded by the user) containing
holdings with manually specified current prices and generates portfolio
metrics, sector summaries and allocation charts. Additional features
include a simple login prompt and a history of uploaded files.

## Expected Excel format

The uploaded workbook must include a sheet named `My_Stocks` with the
following columns (case insensitive):

* `Symbol` – ticker without the `.PSX` suffix
* `Quantity` – number of shares owned
* `Buy Price` – the cost per share at purchase
* `Current Price` – the current market price per share (provided by
  the user; no live data fetching is performed)
* `Sector` – (optional) sector classification for each holding

If the `Sector` column is absent, the app will assign the value
"Unknown" to all positions.

## Features

* **Login prompt.** A minimal login page asks for an email address to
  personalize the session. No external authentication is required.
* **Upload history.** Each uploaded file is saved in a local
  `uploaded_history` directory with a timestamped filename. The sidebar
  lists historical uploads for download.
* **Per‑stock metrics.** Calculates cost, market value, profit/loss
  (absolute and percent) for each holding.
* **Summary KPIs.** Displays total cost, total market value and total
  profit/loss across the portfolio.
* **Performance tables.** Highlights winning and losing positions and
  shows top gainers and losers by percentage P/L.
* **Sector analysis.** Groups holdings by sector to compute aggregate
  cost, market value, P/L and P/L%. Presents a summary table and
  visualizes market value allocation and sector P/L via bar charts.
* **Downloadable report.** Users can download the analyzed portfolio
  (including sector summary) as an Excel file.

Install dependencies using:

```
pip install streamlit pandas openpyxl xlsxwriter
```

Run the app with:

```
streamlit run psx_portfolio_manual_dashboard_updated.py
```

"""

import io
import os
from datetime import datetime
import pandas as pd
import streamlit as st

# Directory for storing uploaded files history
UPLOAD_HISTORY_DIR = "uploaded_history"


def load_manual_portfolio(file) -> pd.DataFrame:
    """Load holdings from the user‑provided Excel file.

    The function expects a sheet named 'My_Stocks' with at least the
    columns: Symbol, Quantity, Buy Price, Current Price. It also reads
    an optional Sector column; if it is missing, sectors are set to
    'Unknown'. Column names are case insensitive.

    Parameters
    ----------
    file : file‑like object
        The uploaded Excel file from Streamlit.

    Returns
    -------
    DataFrame
        The cleaned portfolio data.
    """
    df = pd.read_excel(file, sheet_name="My_Stocks")
    # Normalize column names: strip and lower case
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"symbol", "quantity", "buy price", "current price"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "Sheet 'My_Stocks' must have columns: "
            "Symbol, Quantity, Buy Price, Current Price"
        )

    pf = pd.DataFrame()
    pf["symbol"] = df["symbol"].astype(str).str.upper().str.strip()
    pf["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
    pf["buy_price"] = pd.to_numeric(df["buy price"], errors="coerce").fillna(0.0)
    pf["current_price"] = pd.to_numeric(df["current price"], errors="coerce").fillna(0.0)
    # Sector column
    if "sector" in df.columns:
        pf["sector"] = df["sector"].astype(str).str.strip()
    else:
        pf["sector"] = "Unknown"
    return pf


def compute_portfolio(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate cost, market value, P/L and P/L% per position."""
    pf = df.copy()
    pf["cost"] = pf["quantity"] * pf["buy_price"]
    pf["market_value"] = pf["quantity"] * pf["current_price"]
    pf["pnl"] = pf["market_value"] - pf["cost"]
    pf["pnl_pct"] = (pf["pnl"] / pf["cost"]) * 100
    return pf


def compute_sector_summary(pf: pd.DataFrame) -> pd.DataFrame:
    """Aggregate cost, market value and P/L by sector."""
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
    """Persist the uploaded file to a history directory with timestamp."""
    os.makedirs(UPLOAD_HISTORY_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(original_name)
    safe_base = base.replace(" ", "_")
    filename = f"{safe_base}_{timestamp}{ext}"
    filepath = os.path.join(UPLOAD_HISTORY_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(file_buffer)
    return filepath


# ----------------- Streamlit UI -----------------

st.set_page_config(
    page_title="PSX Manual Portfolio Dashboard",
    page_icon="📊",
    layout="wide",
)

# Simple login
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.title("Login")
    st.write(
        "Enter your email to access the portfolio dashboard. "
        "This is a simple login and does not perform external authentication."
    )
    email_input = st.text_input("Email")
    if st.button("Login"):
        if email_input:
            st.session_state["user"] = email_input
            st.session_state["logged_in"] = True
    st.stop()

st.title("📊 PSX Portfolio Dashboard (Manual Prices)")
st.caption(
    "Upload an Excel file with a 'My_Stocks' sheet containing: "
    "Symbol, Quantity, Buy Price, Current Price and an optional Sector column."
)

# Sidebar: file uploader and history display
st.sidebar.header("Upload Portfolio")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel portfolio file",
    type=["xlsx"],
    help=(
        "File must contain a sheet named 'My_Stocks' with columns: "
        "Symbol, Quantity, Buy Price, Current Price, Sector (optional)."
    ),
)

st.sidebar.subheader("Upload History")
if os.path.exists(UPLOAD_HISTORY_DIR):
    history_files = sorted(os.listdir(UPLOAD_HISTORY_DIR), reverse=True)
    for hfile in history_files:
        filepath = os.path.join(UPLOAD_HISTORY_DIR, hfile)
        with open(filepath, "rb") as f:
            data = f.read()
        st.sidebar.download_button(
            label=hfile,
            data=data,
            file_name=hfile,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.sidebar.write("No uploads yet.")

# Example of the expected sheet structure
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

# Save uploaded file to history
save_upload(uploaded_file.getbuffer(), uploaded_file.name)

# Load and process the portfolio
try:
    portfolio_df = load_manual_portfolio(uploaded_file)
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

pf = compute_portfolio(portfolio_df)
sector_summary = compute_sector_summary(pf)

# Compute total metrics
total_cost = pf["cost"].sum()
total_value = pf["market_value"].sum()
total_pnl = pf["pnl"].sum()
total_pnl_pct = (total_pnl / total_cost * 100) if total_cost else 0.0

# ---------- KPIs ----------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Cost (PKR)", f"{total_cost:,.0f}")
k2.metric("Market Value (PKR)", f"{total_value:,.0f}")
k3.metric("Total P/L (PKR)", f"{total_pnl:,.0f}")
k4.metric("Total P/L (%)", f"{total_pnl_pct:,.2f}%")

st.caption(
    f"Last updated (based on your file): "
    f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ---------- Winners / Losers ----------
st.subheader("Performance Summary")
winners = pf[pf["pnl"] > 0].copy()
losers = pf[pf["pnl"] < 0].copy()

col_win, col_lose = st.columns(2)

with col_win:
    st.markdown("✅ **Winners (Positive P/L)**")
    if winners.empty:
        st.write("No winning positions.")
    else:
        w_disp = winners[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        w_disp["pnl"] = w_disp["pnl"].map(lambda x: f"{x:,.0f}")
        w_disp["pnl_pct"] = w_disp["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        w_disp["market_value"] = w_disp["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(w_disp, use_container_width=True)

with col_lose:
    st.markdown("❌ **Losers (Negative P/L)**")
    if losers.empty:
        st.write("No losing positions.")
    else:
        l_disp = losers[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        l_disp["pnl"] = l_disp["pnl"].map(lambda x: f"{x:,.0f}")
        l_disp["pnl_pct"] = l_disp["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        l_disp["market_value"] = l_disp["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(l_disp, use_container_width=True)

# ---------- Top gainers / losers ----------
st.subheader("Top Movers")

if not pf.empty:
    top_gainers = pf.sort_values("pnl_pct", ascending=False).head(5)
    top_losers = pf.sort_values("pnl_pct", ascending=True).head(5)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("🚀 **Top Gainers (by % P/L)**")
        tg = top_gainers[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        tg["pnl"] = tg["pnl"].map(lambda x: f"{x:,.0f}")
        tg["pnl_pct"] = tg["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        tg["market_value"] = tg["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(tg, use_container_width=True)

    with c2:
        st.markdown("📉 **Top Losers (by % P/L)**")
        tl = top_losers[["symbol", "pnl", "pnl_pct", "market_value"]].copy()
        tl["pnl"] = tl["pnl"].map(lambda x: f"{x:,.0f}")
        tl["pnl_pct"] = tl["pnl_pct"].map(lambda x: f"{x:,.2f}%")
        tl["market_value"] = tl["market_value"].map(lambda x: f"{x:,.0f}")
        st.dataframe(tl, use_container_width=True)

# ---------- Full Detail Table ----------
st.subheader("Full Holdings Detail")

disp = pf.copy()
disp["buy_price"] = disp["buy_price"].map(lambda x: f"{x:,.2f}")
disp["current_price"] = disp["current_price"].map(lambda x: f"{x:,.2f}")
disp["cost"] = disp["cost"].map(lambda x: f"{x:,.0f}")
disp["market_value"] = disp["market_value"].map(lambda x: f"{x:,.0f}")
disp["pnl"] = disp["pnl"].map(lambda x: f"{x:,.0f}")
disp["pnl_pct"] = disp["pnl_pct"].map(lambda x: f"{x:,.2f}%")

st.dataframe(disp, use_container_width=True)

# ---------- Sector Analysis ----------
st.subheader("Sector Summary")
st.dataframe(sector_summary, use_container_width=True)

# Sector charts
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.subheader("Allocation by Sector (Market Value)")
    if not sector_summary.empty:
        st.bar_chart(
            sector_summary.set_index("sector")["market_value"],
            use_container_width=True,
        )

with chart_col2:
    st.subheader("P/L by Sector")
    if not sector_summary.empty:
        st.bar_chart(
            sector_summary.set_index("sector")["pnl"],
            use_container_width=True,
        )

# ---------- Charts by Symbol ----------
c_chart1, c_chart2 = st.columns(2)

with c_chart1:
    st.subheader("Allocation by Market Value (Symbol)")
    alloc_df = pf[["symbol", "market_value"]].set_index("symbol")
    st.bar_chart(alloc_df)

with c_chart2:
    st.subheader("P/L by Symbol")
    pl_df = pf[["symbol", "pnl"]].set_index("symbol")
    st.bar_chart(pl_df)

# ---------- Download analyzed portfolio ----------
st.subheader("Download Analyzed Portfolio as Excel")

output = io.BytesIO()
export_df = pf.copy()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    export_df.to_excel(writer, index=False, sheet_name="Analyzed Portfolio")
    sector_summary.to_excel(writer, index=False, sheet_name="Sector Summary")
output.seek(0)

st.download_button(
    label="📥 Download Portfolio Analysis Excel",
    data=output,
    file_name="psx_portfolio_analysis_manual_with_sector.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
