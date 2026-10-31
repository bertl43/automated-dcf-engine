"""
Automated DCF & Valuation Sensitivity Engine
---------------------------------------------
Author: Albert Nagy
Description: An institutional-grade Discounted Cash Flow (DCF) valuation and 
             sensitivity engine. Calculates dynamic WACC via CAPM, projects 5-year 
             Unlevered Free Cash Flows (FCFF), bridges Enterprise Value to Equity Value,
             and exports a fully formatted, board-ready Excel sensitivity model.
"""

import sys
import argparse
from typing import Dict, Any, Tuple
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CORE DATA STRUCTURES & DEFAULT FINANCIAL PARAMETERS
# ==============================================================================

DEFAULT_DEMO_DATA = {
    "company_name": "SAP SE (Demo Run)",
    "ticker": "SAP.DE",
    "currency": "EUR",
    "share_price": 182.50,
    "shares_outstanding": 1228.0,  # in millions
    "revenue_t0": 31207.0,         # in millions EUR
    "ebit_t0": 6512.0,
    "ebitda_t0": 8450.0,
    "tax_rate": 0.25,
    "risk_free_rate": 0.0260,       # 10Y German Bund Yield (~2.60%)
    "equity_risk_premium": 0.055,  # 5.5% ERP for DACH/Europe
    "beta": 1.08,
    "cost_of_debt_pre_tax": 0.038,
    "total_debt": 11500.0,
    "cash_and_equivalents": 8400.0,
    "projection_years": 5,
    "revenue_growth_rates": [0.085, 0.080, 0.075, 0.070, 0.065],
    "ebit_margins": [0.220, 0.230, 0.235, 0.240, 0.245],
    "da_pct_rev": 0.065,           # D&A as % of Revenue
    "capex_pct_rev": 0.045,        # CapEx as % of Revenue
    "nwc_pct_rev": 0.020,          # Net Working Capital change as % of Revenue
    "terminal_growth_rate": 0.020, # 2.0% perpetual growth
    "exit_ebitda_multiple": 14.5
}


def fetch_company_data(ticker: str) -> Dict[str, Any]:
    """
    Attempts to fetch live market and financial statement data using yfinance.
    Falls back gracefully to institutional demo parameters if offline or unavailable.
    """
    try:
        import yfinance as yf
        print(f"[*] Fetching live financial statement data for {ticker} via yfinance...")
        stock = yf.Ticker(ticker)
        info = stock.info
        
        income_stmt = stock.income_stmt
        balance_sheet = stock.balance_sheet
        
        if income_stmt.empty or balance_sheet.empty:
            raise ValueError(f"Insufficient financial statement data retrieved for {ticker}.")

        revenue_t0 = float(income_stmt.loc["Total Revenue"].iloc[0]) / 1e6
        ebit_t0 = float(income_stmt.loc["EBIT"].iloc[0]) / 1e6 if "EBIT" in income_stmt.index else revenue_t0 * 0.20
        ebitda_t0 = float(income_stmt.loc["EBITDA"].iloc[0]) / 1e6 if "EBITDA" in income_stmt.index else ebit_t0 * 1.25
        
        total_debt = float(info.get("totalDebt", 10000000000)) / 1e6
        total_cash = float(info.get("totalCash", 8000000000)) / 1e6
        shares = float(info.get("sharesOutstanding", 1000000000)) / 1e6
        curr_price = float(info.get("currentPrice", 100.0))
        beta = float(info.get("beta", 1.05)) if info.get("beta") else 1.05
        
        return {
            "company_name": info.get("longName", ticker),
            "ticker": ticker.upper(),
            "currency": info.get("currency", "EUR"),
            "share_price": curr_price,
            "shares_outstanding": shares,
            "revenue_t0": revenue_t0,
            "ebit_t0": ebit_t0,
            "ebitda_t0": ebitda_t0,
            "tax_rate": 0.25,
            "risk_free_rate": 0.0260,
            "equity_risk_premium": 0.055,
            "beta": beta,
            "cost_of_debt_pre_tax": 0.040,
            "total_debt": total_debt,
            "cash_and_equivalents": total_cash,
            "projection_years": 5,
            "revenue_growth_rates": [0.08, 0.075, 0.07, 0.065, 0.06],
            "ebit_margins": [ebit_t0 / revenue_t0] * 5,
            "da_pct_rev": 0.060,
            "capex_pct_rev": 0.045,
            "nwc_pct_rev": 0.015,
            "terminal_growth_rate": 0.020,
            "exit_ebitda_multiple": 14.0
        }
    except Exception as e:
        print(f"[!] Warning: Live data retrieval failed ({e}). Proceeding with verified institutional demo dataset.")
        return DEFAULT_DEMO_DATA.copy()


# ==============================================================================
# 2. VALUATION & SENSITIVITY ENGINE
# ==============================================================================

class DCFValuationEngine:
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.wacc, self.cost_of_equity, self.cost_of_debt_after_tax = self._calculate_wacc()
        self.projections_df = self._project_fcff()
        self.results = self._value_enterprise_and_equity()

    def _calculate_wacc(self) -> Tuple[float, float, float]:
        """Calculates WACC using standard CAPM for Cost of Equity."""
        rf = self.data["risk_free_rate"]
        erp = self.data["equity_risk_premium"]
        beta = self.data["beta"]
        cost_of_equity = rf + (beta * erp)

        tax_rate = self.data["tax_rate"]
        cost_of_debt = self.data["cost_of_debt_pre_tax"] * (1 - tax_rate)

        market_equity = self.data["share_price"] * self.data["shares_outstanding"]
        total_debt = self.data["total_debt"]
        total_capital = market_equity + total_debt

        w_e = market_equity / total_capital
        w_d = total_debt / total_capital

        wacc = (w_e * cost_of_equity) + (w_d * cost_of_debt)
        return wacc, cost_of_equity, cost_of_debt

    def _project_fcff(self) -> pd.DataFrame:
        """Projects 5-Year Unlevered Free Cash Flow (FCFF)."""
        years = [f"Year {t}" for t in range(1, self.data["projection_years"] + 1)]
        df = pd.DataFrame(index=[
            "Revenue", "Revenue Growth", "EBIT", "EBIT Margin",
            "Taxes on EBIT (NOPAT = EBIT*(1-t))", "D&A", "CapEx",
            "Change in NWC", "Unlevered Free Cash Flow (FCFF)",
            "Discount Period (Mid-Year)", "Discount Factor", "PV of FCFF"
        ], columns=years)

        curr_rev = self.data["revenue_t0"]
        for t, yr in enumerate(years):
            g = self.data["revenue_growth_rates"][t]
            curr_rev = curr_rev * (1 + g)
            ebit_margin = self.data["ebit_margins"][t]
            ebit = curr_rev * ebit_margin
            taxes = ebit * self.data["tax_rate"]
            nopat = ebit - taxes
            da = curr_rev * self.data["da_pct_rev"]
            capex = curr_rev * self.data["capex_pct_rev"]
            delta_nwc = curr_rev * self.data["nwc_pct_rev"]
            fcff = nopat + da - capex - delta_nwc
            
            mid_year = (t + 1) - 0.5
            discount_factor = 1.0 / ((1.0 + self.wacc) ** mid_year)
            pv_fcff = fcff * discount_factor

            df[yr] = [
                curr_rev, g, ebit, ebit_margin, taxes, da, capex, delta_nwc,
                fcff, mid_year, discount_factor, pv_fcff
            ]
        return df

    def _value_enterprise_and_equity(self) -> Dict[str, Any]:
        """Calculates Terminal Value, Enterprise Value, and Equity Value per Share."""
        pv_fcff_sum = float(self.projections_df.loc["PV of FCFF"].sum())
        last_fcff = float(self.projections_df.iloc[:, -1]["Unlevered Free Cash Flow (FCFF)"])
        last_ebitda = float(self.projections_df.iloc[:, -1]["EBIT"] + self.projections_df.iloc[:, -1]["D&A"])
        n = self.data["projection_years"]

        # 1. Perpetuity Growth Method (Gordon Growth)
        g = self.data["terminal_growth_rate"]
        terminal_value_pgm = (last_fcff * (1 + g)) / (self.wacc - g)
        pv_terminal_value_pgm = terminal_value_pgm / ((1 + self.wacc) ** n)
        ev_pgm = pv_fcff_sum + pv_terminal_value_pgm
        net_debt = self.data["total_debt"] - self.data["cash_and_equivalents"]
        equity_val_pgm = ev_pgm - net_debt
        share_price_pgm = equity_val_pgm / self.data["shares_outstanding"]

        # 2. Exit Multiple Method
        mult = self.data["exit_ebitda_multiple"]
        terminal_value_mult = last_ebitda * mult
        pv_terminal_value_mult = terminal_value_mult / ((1 + self.wacc) ** n)
        ev_mult = pv_fcff_sum + pv_terminal_value_mult
        equity_val_mult = ev_mult - net_debt
        share_price_mult = equity_val_mult / self.data["shares_outstanding"]

        return {
            "pv_fcff_sum": pv_fcff_sum,
            "terminal_value_pgm": terminal_value_pgm,
            "pv_terminal_value_pgm": pv_terminal_value_pgm,
            "enterprise_value_pgm": ev_pgm,
            "net_debt": net_debt,
            "equity_value_pgm": equity_val_pgm,
            "implied_share_price_pgm": share_price_pgm,
            "terminal_value_mult": terminal_value_mult,
            "pv_terminal_value_mult": pv_terminal_value_mult,
            "enterprise_value_mult": ev_mult,
            "equity_value_mult": equity_val_mult,
            "implied_share_price_mult": share_price_mult,
            "current_share_price": self.data["share_price"],
            "upside_pgm": (share_price_pgm / self.data["share_price"]) - 1.0,
            "upside_mult": (share_price_mult / self.data["share_price"]) - 1.0,
        }

    def generate_sensitivity_matrix(self) -> pd.DataFrame:
        """Generates a 5x5 sensitivity table of Implied Share Price across WACC and g."""
        wacc_base = self.wacc
        g_base = self.data["terminal_growth_rate"]

        wacc_range = [wacc_base - 0.010, wacc_base - 0.005, wacc_base, wacc_base + 0.005, wacc_base + 0.010]
        g_range = [g_base - 0.005, g_base - 0.0025, g_base, g_base + 0.0025, g_base + 0.005]

        last_fcff = float(self.projections_df.iloc[:, -1]["Unlevered Free Cash Flow (FCFF)"])
        n = self.data["projection_years"]
        net_debt = self.results["net_debt"]
        shares = self.data["shares_outstanding"]

        matrix = pd.DataFrame(index=[f"{w*100:.2f}%" for w in wacc_range],
                              columns=[f"{g*100:.2f}%" for g in g_range])

        for w_idx, w in enumerate(wacc_range):
            pv_fcff_sum = sum(
                self.projections_df.iloc[:, t]["Unlevered Free Cash Flow (FCFF)"] / ((1 + w) ** ((t + 1) - 0.5))
                for t in range(n)
            )
            for g_idx, g_val in enumerate(g_range):
                if w <= g_val:
                    matrix.iloc[w_idx, g_idx] = np.nan
                else:
                    tv = (last_fcff * (1 + g_val)) / (w - g_val)
                    pv_tv = tv / ((1 + w) ** n)
                    ev = pv_fcff_sum + pv_tv
                    eq_val = ev - net_debt
                    price = eq_val / shares
                    matrix.iloc[w_idx, g_idx] = price
        return matrix


# ==============================================================================
# 3. EXCEL EXPORT ENGINE (INVESTOR / IB GRADE FORMATTING)
# ==============================================================================

def export_to_excel(engine: DCFValuationEngine, filepath: str = "dcf_valuation_output.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DCF Valuation & Sensitivity"
    ws.views.sheetView[0].showGridLines = True

    # Styling Palette
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    dark_gray_fill = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
    light_blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    soft_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    soft_red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    font_title = Font(name="Calibri", size=15, bold=True, color="1F4E79")
    font_section = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)
    font_italic = Font(name="Calibri", size=9, italic=True, color="595959")

    thin_border_side = Side(style="thin", color="D3D3D3")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom_border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    curr_fmt = '#,##0.0;(#,##0.0);"-"'
    price_fmt = '#,##0.00;(#,##0.00);"-"'
    pct_fmt = "0.0%"

    # Title Banner
    ws["B1"] = f"{engine.data['company_name']} ({engine.data['ticker']}) — Institutional DCF Valuation"
    ws["B1"].font = font_title
    ws["B2"] = f"Reporting Currency: {engine.data['currency']} (in millions, except per share figures) | Valuation Standard: Gordon Growth & Exit Multiples"
    ws["B2"].font = font_italic

    # Section 1: Valuation Assumptions & WACC
    ws["B4"] = "1. KEY VALUATION ASSUMPTIONS & WACC (CAPM)"
    ws["B4"].font = font_section
    ws["B4"].fill = navy_fill
    for col in range(3, 8):
        c = ws.cell(row=4, column=col)
        c.fill = navy_fill

    assumptions = [
        ("Risk-Free Rate (10Y Bund)", engine.data["risk_free_rate"], pct_fmt, "Current Share Price", engine.data["share_price"], price_fmt),
        ("Equity Risk Premium (ERP)", engine.data["equity_risk_premium"], pct_fmt, "Shares Outstanding", engine.data["shares_outstanding"], curr_fmt),
        ("Beta (Unlevered / Adjusted)", engine.data["beta"], "0.00", "Total Debt", engine.data["total_debt"], curr_fmt),
        ("Cost of Equity (Ke)", engine.cost_of_equity, pct_fmt, "Cash & Equivalents", engine.data["cash_and_equivalents"], curr_fmt),
        ("Pre-Tax Cost of Debt (Kd)", engine.data["cost_of_debt_pre_tax"], pct_fmt, "Net Debt", engine.results["net_debt"], curr_fmt),
        ("Effective Tax Rate (t)", engine.data["tax_rate"], pct_fmt, "Terminal Perpetual Growth (g)", engine.data["terminal_growth_rate"], pct_fmt),
        ("Calculated WACC", engine.wacc, pct_fmt, "Exit EBITDA Multiple", engine.data["exit_ebitda_multiple"], "0.0x")
    ]

    for idx, (label1, val1, fmt1, label2, val2, fmt2) in enumerate(assumptions, start=5):
        ws[f"B{idx}"] = label1
        ws[f"B{idx}"].font = font_regular
        ws[f"C{idx}"] = val1
        ws[f"C{idx}"].font = font_bold if "WACC" in label1 else font_regular
        ws[f"C{idx}"].number_format = fmt1
        ws[f"C{idx}"].alignment = Alignment(horizontal="right")

        ws[f"E{idx}"] = label2
        ws[f"E{idx}"].font = font_regular
        ws[f"F{idx}"] = val2
        ws[f"F{idx}"].font = font_bold if "Price" in label2 else font_regular
        ws[f"F{idx}"].number_format = fmt2
        ws[f"F{idx}"].alignment = Alignment(horizontal="right")

    # Section 2: 5-Year FCFF Projections
    ws["B13"] = "2. 5-YEAR UNLEVERED FREE CASH FLOW (FCFF) PROJECTIONS"
    ws["B13"].font = font_section
    ws["B13"].fill = navy_fill
    for col in range(3, 9):
        ws.cell(row=13, column=col).fill = navy_fill

    ws["B14"] = "Line Item (in millions)"
    ws["B14"].font = font_bold
    for idx, yr in enumerate(engine.projections_df.columns, start=3):
        col_letter = get_column_letter(idx)
        ws[f"{col_letter}14"] = yr
        ws[f"{col_letter}14"].font = font_bold
        ws[f"{col_letter}14"].alignment = Alignment(horizontal="right")

    for r_idx, (metric_name, series) in enumerate(engine.projections_df.iterrows(), start=15):
        ws[f"B{r_idx}"] = metric_name
        is_key = metric_name in ["Revenue", "EBIT", "Unlevered Free Cash Flow (FCFF)", "PV of FCFF"]
        ws[f"B{r_idx}"].font = font_bold if is_key else font_regular
        
        for c_idx, val in enumerate(series, start=3):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_bold if is_key else font_regular
            cell.alignment = Alignment(horizontal="right")
            if "Margin" in metric_name or "Growth" in metric_name:
                cell.number_format = pct_fmt
            elif "Factor" in metric_name:
                cell.number_format = "0.000"
            elif "Period" in metric_name:
                cell.number_format = "0.0"
            else:
                cell.number_format = curr_fmt

            if metric_name == "PV of FCFF":
                cell.border = double_bottom_border
                cell.fill = light_blue_fill

    # Section 3: Valuation Bridge & Summary
    start_bridge = 29
    ws[f"B{start_bridge}"] = "3. VALUATION SUMMARY & BRIDGE"
    ws[f"B{start_bridge}"].font = font_section
    ws[f"B{start_bridge}"].fill = navy_fill
    for col in range(3, 9):
        ws.cell(row=start_bridge, column=col).fill = navy_fill

    bridge_headers = ["Metric", "Perpetuity Growth Method", "Exit Multiple Method"]
    for c_idx, h in enumerate(bridge_headers, start=2):
        col_letter = get_column_letter(c_idx if c_idx == 2 else (c_idx * 2 - 1))
        ws[f"{col_letter}{start_bridge+1}"] = h
        ws[f"{col_letter}{start_bridge+1}"].font = font_bold

    bridge_items = [
        ("Cumulative PV of 5-Yr FCFF", engine.results["pv_fcff_sum"], engine.results["pv_fcff_sum"], curr_fmt),
        ("Present Value of Terminal Value", engine.results["pv_terminal_value_pgm"], engine.results["pv_terminal_value_mult"], curr_fmt),
        ("Enterprise Value (EV)", engine.results["enterprise_value_pgm"], engine.results["enterprise_value_mult"], curr_fmt),
        ("Less: Net Debt (Total Debt - Cash)", engine.results["net_debt"], engine.results["net_debt"], curr_fmt),
        ("Equity Value", engine.results["equity_value_pgm"], engine.results["equity_value_mult"], curr_fmt),
        ("Implied Share Price", engine.results["implied_share_price_pgm"], engine.results["implied_share_price_mult"], price_fmt),
        ("Current Trading Price", engine.results["current_share_price"], engine.results["current_share_price"], price_fmt),
        ("Implied Upside / (Downside)", engine.results["upside_pgm"], engine.results["upside_mult"], pct_fmt)
    ]

    for idx, (label, val_pgm, val_mult, fmt) in enumerate(bridge_items, start=start_bridge+2):
        is_highlight = "Implied Share Price" in label or "Upside" in label
        ws[f"B{idx}"] = label
        ws[f"B{idx}"].font = font_bold if is_highlight else font_regular
        
        c_pgm = ws[f"C{idx}"]
        c_pgm.value = val_pgm
        c_pgm.font = font_bold if is_highlight else font_regular
        c_pgm.number_format = fmt
        c_pgm.alignment = Alignment(horizontal="right")

        c_mult = ws[f"E{idx}"]
        c_mult.value = val_mult
        c_mult.font = font_bold if is_highlight else font_regular
        c_mult.number_format = fmt
        c_mult.alignment = Alignment(horizontal="right")

        if "Implied Share Price" in label:
            c_pgm.border = double_bottom_border
            c_mult.border = double_bottom_border
            c_pgm.fill = soft_green_fill
            c_mult.fill = soft_green_fill

    # Section 4: 5x5 WACC vs Perpetual Growth Sensitivity Table
    sens_matrix = engine.generate_sensitivity_matrix()
    sens_start = 40
    ws[f"B{sens_start}"] = "4. 5x5 SENSITIVITY TABLE: IMPLIED SHARE PRICE (WACC vs. TERMINAL GROWTH)"
    ws[f"B{sens_start}"].font = font_section
    ws[f"B{sens_start}"].fill = dark_gray_fill
    for col in range(3, 9):
        ws.cell(row=sens_start, column=col).fill = dark_gray_fill

    ws[f"B{sens_start+1}"] = "WACC \\ g"
    ws[f"B{sens_start+1}"].font = font_bold
    ws[f"B{sens_start+1}"].alignment = Alignment(horizontal="center")

    for c_idx, g_col in enumerate(sens_matrix.columns, start=3):
        col_letter = get_column_letter(c_idx)
        ws[f"{col_letter}{sens_start+1}"] = f"g = {g_col}"
        ws[f"{col_letter}{sens_start+1}"].font = font_bold
        ws[f"{col_letter}{sens_start+1}"].alignment = Alignment(horizontal="right")

    curr_p = engine.data["share_price"]
    for r_idx, (wacc_idx, row_vals) in enumerate(sens_matrix.iterrows(), start=sens_start+2):
        ws[f"B{r_idx}"] = f"WACC {wacc_idx}"
        ws[f"B{r_idx}"].font = font_bold

        for c_idx, price_val in enumerate(row_vals, start=3):
            cell = ws.cell(row=r_idx, column=c_idx, value=price_val)
            cell.font = font_regular
            cell.number_format = price_fmt
            cell.alignment = Alignment(horizontal="right")
            cell.border = thin_border
            
            if price_val is not None and not np.isnan(price_val):
                if price_val > curr_p * 1.05:
                    cell.fill = soft_green_fill
                elif price_val < curr_p * 0.95:
                    cell.fill = soft_red_fill

    # Column Width Auto-Fitting
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 5
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 16

    wb.save(filepath)
    print(f"[+] Formatted Excel model exported successfully to: {filepath}")


# ==============================================================================
# 4. SCRIPT ENTRYPOINT
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="Automated Institutional DCF Valuation Engine")
    parser.add_argument("--ticker", type=str, default="SAP.DE", help="Stock ticker (e.g., SAP.DE, SIE.DE, MSFT)")
    parser.add_argument("--demo", action="store_true", help="Force institutional demo dataset run without network")
    parser.add_argument("--output", type=str, default="dcf_valuation_output.xlsx", help="Output Excel filename")
    args = parser.parse_args()

    print("=" * 75)
    print("      AUTOMATED DCF VALUATION & SENSITIVITY ENGINE")
    print("=" * 75)

    if args.demo:
        data = DEFAULT_DEMO_DATA.copy()
    else:
        data = fetch_company_data(args.ticker)

    engine = DCFValuationEngine(data)

    print(f"[*] Company: {data['company_name']} ({data['ticker']})")
    print(f"[*] Calculated WACC: {engine.wacc*100:.2f}% | Ke: {engine.cost_of_equity*100:.2f}% | Kd(after-tax): {engine.cost_of_debt_after_tax*100:.2f}%")
    print(f"[*] Current Share Price: {engine.results['current_share_price']:.2f} {data['currency']}")
    print("-" * 75)
    print(f"[>] Implied Share Price (Perpetuity Growth): {engine.results['implied_share_price_pgm']:.2f} {data['currency']} (Upside: {engine.results['upside_pgm']*100:+.1f}%)")
    print(f"[>] Implied Share Price (Exit Multiple):    {engine.results['implied_share_price_mult']:.2f} {data['currency']} (Upside: {engine.results['upside_mult']*100:+.1f}%)")
    print("-" * 75)

    export_to_excel(engine, args.output)
    print("=" * 75)
    print(f"[✓] Completed. Review model in {args.output}")

if __name__ == "__main__":
    main()