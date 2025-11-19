import io
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------- DATA LOADING ----------

def load_manual_portfolio(file) -> pd.DataFrame:
    """
    Expects an Excel file with a sheet named 'My_Stocks'
    containing columns:
      - Symbol
      - Quantity
      - Buy Price
      - Current Price
    """
    df = pd.read_excel(file, sheet_name="My_Stocks")

    # Normalize column names
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"symbol", "quantity", "buy price", "current price"}
    if not required.issubset(set(df.columns)):
        raise ValueError(
            "Sheet 'My_Stocks' must have columns: "
            "Symbol, Quantity, Buy Price, Current Price"
        )

    pf = pd.DataFrame()
    pf["symbol"] = df["symbol"].astype(str).str.upper().str.strip()
    pf["quantity"] = df["quantity"].astype(float)
    pf["buy_price"] = df["buy price"].astype(float)
    pf["current_price"] = df["current price"].astype(float)

    return pf


# ---------- CALCULATIONS ----------

def compute_portfolio(df: pd.DataFrame) -> pd.DataFrame:
    pf = df.copy()
    pf["cost"] = pf["quantity"] * pf["buy_price"]
    pf["market_value"] = pf["quantity"] * pf["current_price"]
    pf["pnl"] = pf["market_value"] - pf["cost"]
    pf["pnl_pct"] = (pf["pnl"] / pf["cost"]) * 100
    return pf


# ---------- STREAMLIT UI ----------

st.set_page_config(
    page_title="PSX Manual Portfolio Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PSX Portfolio Dashboard (Manual Prices)")
st.caption(
    "Upload an Excel file with a 'My_Stocks' sheet containing: "
    "Symbol, Quantity, Buy Price, Current Price."
)

st.sidebar.header("Upload Portfolio")

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel portfolio file",
    type=["xlsx"],
    help="Must contain a sheet named 'My_Stocks' with columns: Symbol, Quantity, Buy Price, Current Price.",
)

with st.expander("Sample structure of 'My_Stocks' sheet"):
    st.table(
        pd.DataFrame(
            {
                "Symbol": ["EFERT", "HUBC", "SYS"],
                "Quantity": [100, 250, 50],
                "Buy Price": [90.50, 110.00, 420.00],
                "Current Price": [95.00, 118.25, 410.00],
            }
        )
    )

if not uploaded_file:
    st.info("⬆️ Upload your Excel file to see the analysis dashboard.")
    st.stop()

# Load portfolio
try:
    portfolio_df = load_manual_portfolio(uploaded_file)
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# Compute metrics
pf = compute_portfolio(portfolio_df)

total_cost = pf["cost"].sum()
total_value = pf["market_value"].sum()
total_pnl = pf["pnl"].sum()
total_pnl_pct = (total_pnl / total_cost) * 100 if total_cost else 0.0

# ---------- KPIs ----------

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Cost (PKR)", f"{total_cost:,.0f}")
k2.metric("Market Value (PKR)", f"{total_value:,.0f}")
k3.metric("Total P/L (PKR)", f"{total_pnl:,.0f}")
k4.metric("Total P/L (%)", f"{total_pnl_pct:,.2f}%")

st.caption(f"Last updated (based on your file): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

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

# Top gainers / losers
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

# ---------- Charts ----------

c_chart1, c_chart2 = st.columns(2)

with c_chart1:
    st.subheader("Allocation by Market Value")
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
output.seek(0)

st.download_button(
    label="📥 Download Portfolio Analysis Excel",
    data=output,
    file_name="psx_portfolio_analysis_manual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
