"""
Financial Ratio Analyzer
=========================
Pulls a company's financial statements via yfinance, computes Profitability,
Liquidity, Efficiency, and Leverage ratios across a user-selected timeframe,
and writes each category to its OWN sheet in an Excel workbook so they are
visually and structurally separate from one another.

Run this directly in PyCharm. It will prompt you for:
  1. A ticker symbol (e.g. AAPL, MSFT, TSLA)
  2. Whether you want ANNUAL or QUARTERLY statements
  3. How many periods back to analyze (e.g. 4 = last 4 years/quarters)

Output: <TICKER>_ratio_analysis.xlsx in the same folder as this script,
with 4 tabs: Profitability, Liquidity, Efficiency, Leverage — each with
its own color header, its own formulas, and a short definition column.
"""

import sys
import pandas as pd
import yfinance as yf
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------

def get_user_inputs():
    ticker = input("Enter the stock ticker (e.g. AAPL): ").strip().upper()

    freq = ""
    while freq not in ("A", "Q"):
        freq = input("Timeframe — Annual or Quarterly? (A/Q): ").strip().upper()

    periods = 0
    while periods <= 0:
        try:
            periods = int(input("How many periods back do you want to analyze? (e.g. 4): ").strip())
        except ValueError:
            print("Please enter a whole number.")

    return ticker, freq, periods


# ---------------------------------------------------------------------------

def fetch_statements(ticker, freq):
    stock = yf.Ticker(ticker)

    if freq == "A":
        income = stock.financials
        balance = stock.balance_sheet
        cashflow = stock.cashflow
    else:
        income = stock.quarterly_financials
        balance = stock.quarterly_balance_sheet
        cashflow = stock.quarterly_cashflow

    if income.empty or balance.empty:
        print(f"No data returned for '{ticker}'. Check the ticker and try again.")
        sys.exit(1)

    return income, balance, cashflow, stock


def safe_get(df, row_options, col):
    for name in row_options:
        if name in df.index:
            val = df.loc[name, col]
            if pd.notna(val):
                return val
    return None


# ---------------------------------------------------------------------------

def calc_profitability(income, balance, cols):
    rows = []
    for col in cols:
        revenue = safe_get(income, ["Total Revenue"], col)
        gross_profit = safe_get(income, ["Gross Profit"], col)
        operating_income = safe_get(income, ["Operating Income"], col)
        net_income = safe_get(income, ["Net Income"], col)
        total_assets = safe_get(balance, ["Total Assets"], col)
        total_equity = safe_get(balance, ["Total Stockholder Equity", "Stockholders Equity",
                                           "Total Equity Gross Minority Interest"], col)

        row = {"Period": col.strftime("%Y-%m-%d") if hasattr(col, "strftime") else str(col)}
        row["Gross Margin %"] = (gross_profit / revenue * 100) if gross_profit and revenue else None
        row["Operating Margin %"] = (operating_income / revenue * 100) if operating_income and revenue else None
        row["Net Profit Margin %"] = (net_income / revenue * 100) if net_income and revenue else None
        row["Return on Assets (ROA) %"] = (net_income / total_assets * 100) if net_income and total_assets else None
        row["Return on Equity (ROE) %"] = (net_income / total_equity * 100) if net_income and total_equity else None
        rows.append(row)
    return pd.DataFrame(rows)


def calc_liquidity(balance, cols):
    rows = []
    for col in cols:
        current_assets = safe_get(balance, ["Total Current Assets", "Current Assets"], col)
        current_liab = safe_get(balance, ["Total Current Liabilities", "Current Liabilities"], col)
        inventory = safe_get(balance, ["Inventory"], col)
        cash = safe_get(balance, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments",
                                   "Cash"], col)

        row = {"Period": col.strftime("%Y-%m-%d") if hasattr(col, "strftime") else str(col)}
        row["Current Ratio"] = (current_assets / current_liab) if current_assets and current_liab else None
        row["Quick Ratio"] = ((current_assets - (inventory or 0)) / current_liab) if current_assets and current_liab else None
        row["Cash Ratio"] = (cash / current_liab) if cash and current_liab else None
        rows.append(row)
    return pd.DataFrame(rows)


def calc_efficiency(income, balance, cols):
    rows = []
    for col in cols:
        revenue = safe_get(income, ["Total Revenue"], col)
        cogs = safe_get(income, ["Cost Of Revenue", "Cost Of Goods Sold"], col)
        total_assets = safe_get(balance, ["Total Assets"], col)
        inventory = safe_get(balance, ["Inventory"], col)
        receivables = safe_get(balance, ["Net Receivables", "Accounts Receivable", "Receivables"], col)

        row = {"Period": col.strftime("%Y-%m-%d") if hasattr(col, "strftime") else str(col)}
        row["Asset Turnover"] = (revenue / total_assets) if revenue and total_assets else None
        row["Inventory Turnover"] = (cogs / inventory) if cogs and inventory else None
        row["Receivables Turnover"] = (revenue / receivables) if revenue and receivables else None
        rows.append(row)
    return pd.DataFrame(rows)


def calc_leverage(income, balance, cols):
    rows = []
    for col in cols:
        total_debt = safe_get(balance, ["Total Debt", "Long Term Debt"], col)
        total_equity = safe_get(balance, ["Total Stockholder Equity", "Stockholders Equity",
                                           "Total Equity Gross Minority Interest"], col)
        total_assets = safe_get(balance, ["Total Assets"], col)
        total_liab = safe_get(balance, ["Total Liabilities Net Minority Interest", "Total Liab"], col)
        operating_income = safe_get(income, ["Operating Income"], col)
        interest_expense = safe_get(income, ["Interest Expense"], col)

        row = {"Period": col.strftime("%Y-%m-%d") if hasattr(col, "strftime") else str(col)}
        row["Debt-to-Equity"] = (total_debt / total_equity) if total_debt and total_equity else None
        row["Debt Ratio"] = (total_liab / total_assets) if total_liab and total_assets else None
        row["Equity Multiplier"] = (total_assets / total_equity) if total_assets and total_equity else None
        row["Interest Coverage"] = (operating_income / abs(interest_expense)) if operating_income and interest_expense else None
        rows.append(row)
    return pd.DataFrame(rows)


# Definitions shown alongside each ratio in the Excel output
DEFINITIONS = {
    "Gross Margin %": "Gross profit as a % of revenue — pricing power and production cost control.",
    "Operating Margin %": "Operating income as a % of revenue — profitability from core operations before interest/tax.",
    "Net Profit Margin %": "Net income as a % of revenue — bottom-line profitability after all expenses.",
    "Return on Assets (ROA) %": "Net income relative to total assets — how efficiently assets generate profit.",
    "Return on Equity (ROE) %": "Net income relative to shareholder equity — return generated on owners' capital.",
    "Current Ratio": "Current assets / current liabilities — ability to cover short-term obligations.",
    "Quick Ratio": "(Current assets - inventory) / current liabilities — liquidity excluding hard-to-sell inventory.",
    "Cash Ratio": "Cash & equivalents / current liabilities — most conservative liquidity measure.",
    "Asset Turnover": "Revenue / total assets — how efficiently assets generate sales.",
    "Inventory Turnover": "Cost of goods sold / inventory — how quickly inventory is sold and replaced.",
    "Receivables Turnover": "Revenue / accounts receivable — how quickly customer credit is collected.",
    "Debt-to-Equity": "Total debt / total equity — reliance on debt vs. equity financing.",
    "Debt Ratio": "Total liabilities / total assets — proportion of assets financed by debt.",
    "Equity Multiplier": "Total assets / total equity — degree of financial leverage.",
    "Interest Coverage": "Operating income / interest expense — ability to service debt interest from operations.",
}


# ---------------------------------------------------------------------------
# 4. EXCEL EXPORT — each category gets its own distinctly styled sheet
# ---------------------------------------------------------------------------

SHEET_COLORS = {
    "Profitability": "1F6F43",  # green
    "Liquidity": "1F4E79",      # blue
    "Efficiency": "7030A0",     # purple
    "Leverage": "B7472A",       # red/brown
}


def write_sheet(writer, sheet_name, df):
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    ws = writer.sheets[sheet_name]
    color = SHEET_COLORS.get(sheet_name, "444444")

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    title_cell = ws.cell(row=1, column=1, value=f"{sheet_name} Ratios")
    title_cell.font = Font(size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Header row styling
    header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=2, column=c_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Column widths
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = max(18, len(col_name) + 2)

    # Round numeric values for readability
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.00"

    # Append a definitions block below the data
    def_start_row = ws.max_row + 3
    def_title = ws.cell(row=def_start_row, column=1, value="Ratio Definitions")
    def_title.font = Font(bold=True, italic=True, color=color)
    for i, ratio_name in enumerate(df.columns[1:], start=1):
        ws.cell(row=def_start_row + i, column=1, value=ratio_name).font = Font(bold=True)
        ws.cell(row=def_start_row + i, column=2, value=DEFINITIONS.get(ratio_name, ""))


def export_to_excel(ticker, profitability, liquidity, efficiency, leverage):
    filename = f"{ticker}_ratio_analysis.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        write_sheet(writer, "Profitability", profitability)
        write_sheet(writer, "Liquidity", liquidity)
        write_sheet(writer, "Efficiency", efficiency)
        write_sheet(writer, "Leverage", leverage)
    print(f"\nDone. Workbook saved as: {filename}")
    print("Each ratio category is on its own tab, with its own color and a definitions key.")


# ---------------------------------------------------------------------------
# 5. MAIN
# ---------------------------------------------------------------------------

def main():
    ticker, freq, periods = get_user_inputs()
    income, balance, cashflow, stock = fetch_statements(ticker, freq)

    # Use only as many periods as the user asked for (and are actually available)
    available_periods = min(periods, len(income.columns), len(balance.columns))
    if available_periods < periods:
        print(f"Note: only {available_periods} period(s) of data are available for {ticker}.")

    cols = income.columns[:available_periods]

    print(f"\nCalculating ratios for {ticker} across {available_periods} "
          f"{'annual' if freq == 'A' else 'quarterly'} period(s)...")

    profitability = calc_profitability(income, balance, cols)
    liquidity = calc_liquidity(balance, cols)
    efficiency = calc_efficiency(income, balance, cols)
    leverage = calc_leverage(income, balance, cols)

    # Quick console preview
    for name, df in [("PROFITABILITY", profitability), ("LIQUIDITY", liquidity),
                      ("EFFICIENCY", efficiency), ("LEVERAGE", leverage)]:
        print(f"\n--- {name} ---")
        print(df.to_string(index=False))

    export_to_excel(ticker, profitability, liquidity, efficiency, leverage)


if __name__ == "__main__":
    main()