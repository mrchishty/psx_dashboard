"""
Updated PSX Portfolio Dashboard with Sector Analysis and Upload History
====================================================================

This Streamlit app extends the original PSX portfolio dashboard to
include sector information, sector allocation charts, and sector‑wise
summary statistics. It also maintains a history of uploaded Excel files
for later reference.

Features
--------

* **Login screen.** A simple email prompt is used in lieu of full
  Google authentication. Entering your email allows you to proceed to
  the dashboard. (Full Google OAuth integration requires external
  configuration and is not included here.)

* **Excel upload.** Upload a `.xlsx` file with a sheet named
  `My_Stocks` that contains the columns `Symbol`, `Quantity`,
  `Buy Price` and `Sector`. Additional columns are ignored. A copy of
  every uploaded file is saved to the `uploaded_history` directory with
  a timestamped filename. Links to download prior uploads appear in
  the sidebar.

* **Live price retrieval.** Latest prices for PSX symbols are fetched
  using Yahoo Finance (`yfinance` library). Symbols are expected
  without the `.PSX` suffix; the suffix is appended automatically.

* **Portfolio metrics.** For each holding, the app calculates cost,
  market value, profit/loss (PKR) and profit/loss percent.

* **Sector analysis.** Holdings are grouped by the `Sector` column to
  compute total cost, market value and P/L for each sector. Summary
  tables and bar charts display the allocation and performance across
  sectors.

* **Download report.** A button allows users to download an Excel file
  that contains two sheets: `Portfolio_Details` (per‑stock metrics) and
  `Sector_Summary` (sector‑level metrics).

Dependencies
------------

The script uses the following Python libraries:

* `streamlit`
* `pandas`
* `yfinance`
* `openpyxl` (for reading Excel files)
* `xlsxwriter` (for writing Excel files)

Install them via pip:

```
pip install streamlit pandas yfinance openpyxl xlsxwriter
```

Usage
-----

Run the application with Streamlit:

```
streamlit run psx_excel_dashboard_app_updated.py
```

Upon launching, the app prompts for an email. After logging in,
upload your Excel file and explore the portfolio and sector analysis.

"""

import io
import os
from datetime import datetime
from typing import Dict, List

import pandas as pd
import streamlit as st
import yfinance as yf

# Yahoo Finance suffix for PSX symbols
PSX_SUFFIX = ".PSX"
# Directory to store uploaded files history
UPLOAD_HISTORY_DIR = "uploaded_history"


def get_live_psx_prices(symbols: List[str]) -> Dict[str, float]:
    """Fetch latest close price for a list of PSX symbols via Yahoo Finance.

    Parameters
    ----------
    symbols : list of str
        Tickers without the `.PSX` suffix.

    Returns
    -------
    dict
        Mapping of symbol to current price. Missing symbols will be
        absent from the dictionary.
    """
    prices: Dict[str, float] = {}
    if not symbols:
        return prices
    # Append the PSX suffix to each symbol
    yf_symbols = [s + PSX_SUFFIX for s in symbols]
    try:
        data = yf.download(
            " ".join(yf_symbols),
            period="1d",
            interval="1d",
            progress=False,
            auto_adjust=False,
        )
        # If multiple tickers, data columns will be a MultiIndex
        if isinstance(data.columns, pd.MultiIndex):
            last_close = data["Close"].iloc[-1]
            for yf_sym, price in last_close.items():
                if pd.isna(price):
                    continue
                base_symbol = yf_sym.split(".")[0]
                prices[base_symbol] = float(price)
        else:
            # Single ticker case
            last_close = data["Close"].iloc[-1]
            base_symbol = symbols[0]
            prices[base_symbol] = float(last_close)
    except Exception as e:
        st.warning(f"Error while fetching prices: {e}")
    return prices


def compute_sector_summary(pf: pd.DataFrame) -> pd.DataFrame:
    """Aggregate portfolio metrics by sector.

    Parameters
    ----------
    pf : DataFrame
        DataFrame with per‑stock metrics including `Sector`, `cost`,
        `market_value` and `pnl` columns.

    Returns
    -------
    DataFrame
        Summary table with total cost, market value, P/L and P/L % per
        sector.
    """
    summary = (
        pf.groupby("Sector")[["cost", "market_value", "pnl"]]
        .sum()
        .reset_index()
    )
    summary["pnl_pct"] = summary.apply(
        lambda row: (row["pnl"] / row["cost"] * 100) if row["cost"] != 0 else 0,
        axis=1,
    )
    return summary


def save_upload(file_buffer: bytes, original_name: str) -> str:
    """Save uploaded file buffer to the history directory with a timestamp.

    Parameters
    ----------
    file_buffer : bytes
        Raw bytes of the uploaded file.
    original_name : str
        Original filename from the upload widget.

    Returns
    -------
    str
        The path where the file was saved.
    """
    os.makedirs(UPLOAD_HISTORY_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(original_name)
    safe_base = base.replace(" ", "_")
    filename = f"{safe_base}_{timestamp}{ext}"
    filepath = os.path.join(UPLOAD_HISTORY_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(file_buffer)
    return filepath


def load_portfolio_from_excel(uploaded_file) -> pd.DataFrame:
    """Load the user's portfolio from the uploaded Excel file.

    Expects a sheet named `My_Stocks` with at least the columns:
    `Symbol`, `Quantity`, `Buy Price` and `Sector`. If `Sector` is
    missing, a default value of 'Unknown' will be used.

    Parameters
    ----------
    uploaded_file : a file‑like object
        The uploaded Excel file from Streamlit.

    Returns
    -------
    DataFrame
        DataFrame of the holdings.
    """
    df = pd.read_excel(uploaded_file, sheet_name="My_Stocks")
    # Normalize column names by stripping and lowering
    df.columns = [c.strip() for c in df.columns]
    # Ensure required columns exist
    required_cols = {"Symbol", "Quantity", "Buy Price"}
    if not required_cols.issubset(set(df.columns)):
        missing = required_cols - set(df.columns)
        raise ValueError(f"Missing required columns in 'My_Stocks' sheet: {missing}")
    # Ensure Sector column exists
    if "Sector" not in df.columns:
        df["Sector"] = "Unknown"
    # Clean up data types
    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Buy Price"] = pd.to_numeric(df["Buy Price"], errors="coerce").fillna(0.0)
    df["Sector"] = df["Sector"].astype(str).str.strip()
    return df


def main():
    # Configure the page
    st.set_page_config(
        page_title="PSX Portfolio Dashboard with Sector Analysis",
        page_icon="📈",
        layout="wide",
    )

    # Simple login mechanism
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if not st.session_state["logged_in"]:
        st.title("Login")
        st.write(
            "Enter your email to proceed. (Full Google OAuth requires external configuration.)"
        )
        email = st.text_input("Email")
        if st.button("Login"):
            if not email:
                st.warning("Please enter a valid email.")
            else:
                st.session_state["user"] = email
                st.session_state["logged_in"] = True
        st.stop()

    st.title("📈 PSX Portfolio Dashboard with Sector Analysis")
    st.caption(
        "Upload an Excel file with a sheet named 'My_Stocks' containing columns "
        "Symbol, Quantity, Buy Price and Sector. Sector is optional but recommended."
    )

    # Sidebar: upload widget and history display
    st.sidebar.header("Portfolio Input")
    uploaded_file = st.sidebar.file_uploader(
        "Upload PSX portfolio Excel file", type=["xlsx"], help="Expected sheet: My_Stocks"
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

    # If no file is uploaded, show instructions and stop
    if not uploaded_file:
        st.info(
            "⬆️ Please upload your portfolio Excel file to see the dashboard."
        )
        return

    # Save uploaded file to history
    save_upload(uploaded_file.getbuffer(), uploaded_file.name)

    # Load and validate portfolio
    try:
        portfolio_df = load_portfolio_from_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading portfolio: {e}")
        return

    # Fetch live PSX prices
    symbols = portfolio_df["Symbol"].tolist()
    live_prices = get_live_psx_prices(symbols)
    portfolio_df["current_price"] = portfolio_df["Symbol"].map(live_prices)

    # Warn if some prices are missing
    missing_prices = portfolio_df[portfolio_df["current_price"].isna()]["Symbol"].tolist()
    if missing_prices:
        st.warning(
            "Could not fetch live price for: " + ", ".join(missing_prices)
            + ". Check if these symbols exist on Yahoo Finance with .PSX suffix."
        )
        # Fill missing prices with zero to avoid NaNs in calculations
        portfolio_df["current_price"] = portfolio_df["current_price"].fillna(0.0)

    # Compute per‑holding metrics
    portfolio_df["cost"] = portfolio_df["Quantity"] * portfolio_df["Buy Price"]
    portfolio_df["market_value"] = portfolio_df["Quantity"] * portfolio_df["current_price"]
    portfolio_df["pnl"] = portfolio_df["market_value"] - portfolio_df["cost"]
    portfolio_df["pnl_pct"] = portfolio_df.apply(
        lambda row: (row["pnl"] / row["cost"] * 100) if row["cost"] != 0 else 0.0,
        axis=1,
    )

    total_cost = portfolio_df["cost"].sum()
    total_value = portfolio_df["market_value"].sum()
    total_pnl = portfolio_df["pnl"].sum()
    total_pnl_pct = (total_pnl / total_cost * 100) if total_cost else 0.0

    # Compute sector summary
    sector_summary = compute_sector_summary(portfolio_df)

    # Display portfolio KPIs
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    kpi_col1.metric("Total Cost (PKR)", f"{total_cost:,.0f}")
    kpi_col2.metric("Market Value (PKR)", f"{total_value:,.0f}")
    kpi_col3.metric("Total P/L (PKR)", f"{total_pnl:,.0f}")
    kpi_col4.metric("Total P/L (%)", f"{total_pnl_pct:,.2f}%")
    st.caption(
        f"Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )

    # Detailed holdings table
    st.subheader("Holdings Detail")
    display_pf = portfolio_df.copy()
    # Format numeric columns for display
    display_pf["Buy Price"] = display_pf["Buy Price"].map(lambda x: f"{x:,.2f}")
    display_pf["current_price"] = display_pf["current_price"].map(lambda x: f"{x:,.2f}")
    display_pf["cost"] = display_pf["cost"].map(lambda x: f"{x:,.0f}")
    display_pf["market_value"] = display_pf["market_value"].map(lambda x: f"{x:,.0f}")
    display_pf["pnl"] = display_pf["pnl"].map(lambda x: f"{x:,.0f}")
    display_pf["pnl_pct"] = display_pf["pnl_pct"].map(lambda x: f"{x:,.2f}%")
    st.dataframe(display_pf, use_container_width=True)

    # Sector analysis
    st.subheader("Sector Allocation and Summary")
    st.dataframe(sector_summary, use_container_width=True)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.write("Allocation by Sector (Market Value)")
        if not sector_summary.empty:
            st.bar_chart(
                sector_summary.set_index("Sector")["market_value"],
                use_container_width=True,
            )
    with chart_col2:
        st.write("P/L by Sector")
        if not sector_summary.empty:
            st.bar_chart(
                sector_summary.set_index("Sector")["pnl"],
                use_container_width=True,
            )

    # Allow user to download a combined report as Excel
    st.subheader("Download Report")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        portfolio_df.to_excel(writer, index=False, sheet_name="Portfolio_Details")
        sector_summary.to_excel(writer, index=False, sheet_name="Sector_Summary")
    output.seek(0)
    st.download_button(
        label="📥 Download Excel Report",
        data=output,
        file_name="psx_portfolio_with_sector_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
