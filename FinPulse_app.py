import io
from datetime import date
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# 1. PAGE CONFIG & CUSTOM STYLING
st.set_page_config(
    page_title="FinPulse | Loan & Portfolio Analytics Suite",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #6c757d;
        font-size: 0.95rem;
        margin-bottom: 1.2rem;
    }
    div[data-testid="stMetric"] {
        background-color: rgba(28, 131, 225, 0.06);
        border: 1px solid rgba(28, 131, 225, 0.18);
        padding: 12px 16px;
        border-radius: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# 2. PURE-NUMPY VECTORIZED FINANCIAL ENGINE
class FinPulseEngine:
    """Zero-dependency vectorized financial engine mirroring Excel functions."""

    @staticmethod
    def pmt(rate: float, nper: int, pv: float, fv: float = 0.0) -> float:
        if rate == 0:
            return -(pv + fv) / nper
        pvif = (1.0 + rate) ** nper
        return -rate * (pv * pvif + fv) / (pvif - 1.0)

    @staticmethod
    def ipmt(rate: float, per: np.ndarray, nper: int, pv: float, fv: float = 0.0) -> np.ndarray:
        if rate == 0:
            return np.zeros_like(per, dtype=float)
        payment = FinPulseEngine.pmt(rate, nper, pv, fv)
        factor = (1.0 + rate) ** (per - 1)
        prev_balance = pv * factor + payment * (factor - 1.0) / rate
        return -prev_balance * rate

    @staticmethod
    def ppmt(rate: float, per: np.ndarray, nper: int, pv: float, fv: float = 0.0) -> np.ndarray:
        payment = FinPulseEngine.pmt(rate, nper, pv, fv)
        interest = FinPulseEngine.ipmt(rate, per, nper, pv, fv)
        return payment - interest

    @staticmethod
    def fv(rate: float, nper: int, pmt: float, pv: float) -> float:
        if rate == 0:
            return -(pv + pmt * nper)
        factor = (1.0 + rate) ** nper
        return -(pv * factor + pmt * (factor - 1.0) / rate)

    @staticmethod
    def pv(rate: float, nper: int, pmt: float, fv: float = 0.0) -> float:
        if rate == 0:
            return -(fv + pmt * nper)
        factor = (1.0 + rate) ** nper
        return -((fv + pmt * (factor - 1.0) / rate) / factor)

    @staticmethod
    def nper(rate: float, pmt: float, pv: float, fv: float = 0.0) -> Optional[float]:
        if rate == 0:
            return -(pv + fv) / pmt if pmt != 0 else None
        num = pmt - fv * rate
        den = pv * rate + pmt
        if den == 0 or (num / den) <= 0:
            return None
        return float(np.log(num / den) / np.log(1.0 + rate))

    
    @staticmethod
    def npv(rate: float, cashflows: List[float]) -> float:
        periods = np.arange(len(cashflows))
        cf = np.asarray(cashflows, dtype=float)
        return float(np.sum(cf / ((1.0 + rate) ** periods)))

    @staticmethod
    def irr(cashflows: List[float], guess: float = 0.1, tol: float = 1e-7, max_iter: int = 250) -> Optional[float]:
        cf = np.asarray(cashflows, dtype=float)
        if not (np.any(cf > 0) and np.any(cf < 0)):
            return None
        periods = np.arange(len(cf))
        r = guess
        for _ in range(max_iter):
            r = max(r, -0.99)
            f_val = np.sum(cf / ((1.0 + r) ** periods))
            f_der = np.sum(-periods * cf / ((1.0 + r) ** (periods + 1.0)))
            if abs(f_der) < 1e-12:
                break
            r_next = r - f_val / f_der
            if abs(r_next - r) < tol:
                return float(r_next)
            r = r_next
        return None

    @staticmethod
    def xnpv(rate: float, cashflows: List[float], dates: List[date]) -> float:
        if rate <= -1.0:
            return float("inf")
        d0 = min(dates)
        years = np.array([(d - d0).days / 365.0 for d in dates], dtype=float)
        cf = np.asarray(cashflows, dtype=float)
        return float(np.sum(cf / ((1.0 + rate) ** years)))

    @staticmethod
    def xirr(cashflows: List[float], dates: List[date], guess: float = 0.1, tol: float = 1e-7, max_iter: int = 250) -> Optional[float]:
        cf = np.asarray(cashflows, dtype=float)
        if not (np.any(cf > 0) and np.any(cf < 0)):
            return None
        d0 = min(dates)
        years = np.array([(d - d0).days / 365.0 for d in dates], dtype=float)
        r = guess
        for _ in range(max_iter):
            r = max(r, -0.99)
            f_val = np.sum(cf / ((1.0 + r) ** years))
            f_der = np.sum(-years * cf / ((1.0 + r) ** (years + 1.0)))
            if abs(f_der) < 1e-12:
                break
            r_next = r - f_val / f_der
            if abs(r_next - r) < tol:
                return float(r_next)
            r = r_next
        return None


# 3. SCHEDULE BUILDER & FORMATTED EXCEL EXPORTER
def build_amortization_schedule(
    principal: float, annual_rate: float, months: int, extra_monthly: float = 0.0
) -> Tuple[pd.DataFrame, float, float]:
    """Generates month-by-month amortization schedule supporting optional extra prepayments."""
    monthly_rate = (annual_rate / 100.0) / 12.0
    base_emi = -FinPulseEngine.pmt(monthly_rate, months, principal)

    if extra_monthly == 0.0:
        periods = np.arange(1, months + 1)
        ppmt_arr = -FinPulseEngine.ppmt(monthly_rate, periods, months, principal)
        ipmt_arr = -FinPulseEngine.ipmt(monthly_rate, periods, months, principal)
        emi_arr = np.full(months, base_emi)
        closing_bal = np.maximum(0.0, principal - np.cumsum(ppmt_arr))
        opening_bal = closing_bal + ppmt_arr

        df = pd.DataFrame(
            {
                "Month": periods,
                "Opening Balance": np.round(opening_bal, 2),
                "Scheduled EMI (PMT)": np.round(emi_arr, 2),
                "Principal Paid (PPMT)": np.round(ppmt_arr, 2),
                "Interest Paid (IPMT)": np.round(ipmt_arr, 2),
                "Closing Balance": np.round(closing_bal, 2),
            }
        )
        return df, base_emi, float(np.sum(ipmt_arr))

    # Iterative schedule when extra monthly prepayments shorten loan tenure
    rows = []
    balance = principal
    month = 1
    while balance > 0.01 and month <= months:
        interest = balance * monthly_rate
        total_outflow = min(balance + interest, base_emi + extra_monthly)
        principal_paid = total_outflow - interest
        closing = max(0.0, balance - principal_paid)
        rows.append(
            {
                "Month": month,
                "Opening Balance": round(balance, 2),
                "Scheduled EMI (PMT)": round(total_outflow, 2),
                "Principal Paid (PPMT)": round(principal_paid, 2),
                "Interest Paid (IPMT)": round(interest, 2),
                "Closing Balance": round(closing, 2),
            }
        )
        balance = closing
        month += 1

    df = pd.DataFrame(rows)
    return df, base_emi, float(df["Interest Paid (IPMT)"].sum())


def export_styled_excel(df_schedule: pd.DataFrame, summary: Dict[str, object]) -> bytes:
    """Writes a styled multi-sheet Excel report using OpenPyXL."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df = pd.DataFrame(list(summary.items()), columns=["Parameter / Metric", "Value"])
        summary_df.to_excel(writer, index=False, sheet_name="Executive Summary")
        df_schedule.to_excel(writer, index=False, sheet_name="Amortization Schedule")

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_side = Side(style="thin", color="D9D9D9")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for name, ws in writer.sheets.items():
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for idx, cell in enumerate(row):
                    cell.border = border
                    if name == "Amortization Schedule" and idx > 0:
                        cell.number_format = "#,##0.00"

            for col in ws.columns:
                max_len = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

    return buffer.getvalue()


# 4. SIDEBAR & HEADER
with st.sidebar:
    st.header("⚙️ FinPulse Settings")
    currency = st.selectbox("Display Currency", ["₹ (INR)", "$ (USD)", "€ (EUR)", "£ (GBP)"], index=0)
    sym = currency.split()[0]

    st.divider()
    st.markdown("**Developer Profile**")
    st.caption("Nikhil Kalia")
    st.divider()
    st.markdown("**Supported Formulas**")
    st.code("PMT, PPMT, IPMT\nPV, FV, NPER\nNPV, IRR, XNPV, XIRR", language="text")

st.markdown('<div class="main-header">📈 FinPulse: Loan & Portfolio Analytics Suite</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Interactive Debt Amortization, Irregular Cash-Flow Valuation (XNPV/XIRR), and Wealth Planning Engine</div>',
    unsafe_allow_html=True,
)

tab1, tab2, tab3 = st.tabs(
    [
        "🏦 Loan Amortization (PMT / PPMT / IPMT)",
        "📊 Portfolio Cash-Flow Evaluator (XNPV / XIRR)",
        "🎯 SIP & Goal Solver (FV / PV / NPER)",
    ]
)

# TAB 1: LOAN EMI & AMORTIZATION SIMULATOR
with tab1:
    ctrl_col, viz_col = st.columns([1, 2.1], gap="large")

    with ctrl_col:
        st.subheader("Loan Parameters")
        principal = st.slider(f"Loan Principal ({sym})", 50_000, 10_000_000, 1_500_000, step=50_000)
        annual_rate = st.slider("Annual Interest Rate (%)", 1.0, 24.0, 8.5, step=0.1)
        tenure_mode = st.radio("Tenure Input Mode", ["Years", "Months"], horizontal=True)

        if tenure_mode == "Years":
            years = st.slider("Tenure (Years)", 1, 30, 15)
            months = years * 12
        else:
            months = st.slider("Tenure (Months)", 6, 360, 180, step=6)

        extra_pmt = st.number_input(
            f"Optional Extra Monthly Prepayment ({sym})",
            min_value=0,
            max_value=500_000,
            value=0,
            step=1000,
            help="Simulate paying extra principal every month to reduce total interest and loan duration.",
        )

    df_amort, base_emi, total_interest = build_amortization_schedule(principal, annual_rate, months, extra_pmt)
    _, _, baseline_interest = build_amortization_schedule(principal, annual_rate, months, 0.0)
    total_payable = principal + total_interest
    actual_months = len(df_amort)

    with viz_col:
        st.subheader("Amortization Overview")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Base Monthly EMI", f"{sym}{base_emi:,.2f}")
        k2.metric("Total Interest", f"{sym}{total_interest:,.2f}", delta=f"-{sym}{baseline_interest - total_interest:,.0f} saved" if extra_pmt > 0 else None)
        k3.metric("Total Repayment", f"{sym}{total_payable:,.2f}")
        k4.metric("Payoff Duration", f"{actual_months} mos", delta=f"-{months - actual_months} mos" if extra_pmt > 0 else None)

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.1))

        # Balance Decay Curve
        ax1.fill_between(df_amort["Month"], df_amort["Closing Balance"], color="#1f77b4", alpha=0.22)
        ax1.plot(df_amort["Month"], df_amort["Closing Balance"], color="#1f77b4", lw=2, label="Remaining Balance")
        ax1.plot(df_amort["Month"], np.cumsum(df_amort["Interest Paid (IPMT)"]), color="#ff7f0e", ls="--", lw=2, label="Cumulative Interest")
        ax1.set_title("Principal Balance vs. Cumulative Interest")
        ax1.set_xlabel("Month")
        ax1.set_ylabel(f"Amount ({sym})")
        ax1.grid(True, ls=":", alpha=0.6)
        ax1.legend()

        # Stacked Area Breakdown
        ax2.stackplot(
            df_amort["Month"],
            df_amort["Principal Paid (PPMT)"],
            df_amort["Interest Paid (IPMT)"],
            labels=["Principal (PPMT)", "Interest (IPMT)"],
            colors=["#2ca02c", "#d62728"],
            alpha=0.82,
        )
        ax2.set_title("Monthly Payment Composition (PPMT vs. IPMT)")
        ax2.set_xlabel("Month")
        ax2.set_ylabel(f"Monthly Share ({sym})")
        ax2.grid(True, ls=":", alpha=0.6)
        ax2.legend(loc="upper right")

        plt.tight_layout()
        st.pyplot(fig)

    st.divider()
    dl_col, info_col = st.columns([1, 2])
    summary_dict = {
        f"Loan Principal ({sym})": round(principal, 2),
        "Annual Interest Rate (%)": annual_rate,
        "Original Tenure (Months)": months,
        "Actual Payoff Tenure (Months)": actual_months,
        f"Base Monthly EMI ({sym})": round(base_emi, 2),
        f"Extra Monthly Prepayment ({sym})": round(extra_pmt, 2),
        f"Total Interest Payable ({sym})": round(total_interest, 2),
        f"Total Amount Payable ({sym})": round(total_payable, 2),
    }
    with dl_col:
        st.download_button(
            label="📥 Export Styled Excel Report (.xlsx)",
            data=export_styled_excel(df_amort, summary_dict),
            file_name="FinPulse_Amortization_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with info_col:
        st.caption("Includes two formatted worksheets: **Executive Summary** and **Amortization Schedule**.")

    st.dataframe(df_amort, use_container_width=True, height=310)


# TAB 2: IRREGULAR CASH-FLOW EVALUATOR (XNPV & XIRR)
with tab2:
    st.subheader("Irregular Cash-Flow Valuation (XNPV & XIRR)")
    up_col, rate_col = st.columns([2, 1])

    with up_col:
        uploaded = st.file_uploader("Upload Cash-Flow Workbook (.xlsx with columns: Date, Cash_Flow)", type=["xlsx"])
    with rate_col:
        disc_rate = st.number_input("Annual Discount Rate (%)", min_value=0.0, max_value=60.0, value=10.0, step=0.5)

    if uploaded is not None:
        df_cf = pd.read_excel(uploaded)
    else:
        df_cf = pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-15", "2024-06-20", "2024-11-10", "2025-04-05", "2025-12-18", "2026-06-22"]),
                "Cash_Flow": [-500000, -150000, 120000, 210000, 260000, 310000],
            }
        )
        sample_buf = io.BytesIO()
        df_cf.to_excel(sample_buf, index=False, engine="openpyxl")
        st.download_button(
            label="📄 Download Sample Cash-Flow Template (.xlsx)",
            data=sample_buf.getvalue(),
            file_name="FinPulse_Sample_Cash_Flows.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if {"Date", "Cash_Flow"}.issubset(df_cf.columns):
        df_cf["Date"] = pd.to_datetime(df_cf["Date"])
        df_cf = df_cf.sort_values("Date").reset_index(drop=True)

        dates = df_cf["Date"].dt.date.tolist()
        flows = df_cf["Cash_Flow"].astype(float).tolist()

        xnpv_val = FinPulseEngine.xnpv(disc_rate / 100.0, flows, dates)
        xirr_val = FinPulseEngine.xirr(flows, dates)
        npv_val = FinPulseEngine.npv(disc_rate / 100.0, flows)
        irr_val = FinPulseEngine.irr(flows)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Exact XNPV", f"{sym}{xnpv_val:,.2f}")
        m2.metric("Exact XIRR", f"{xirr_val * 100:.2f}%" if xirr_val is not None else "No Convergence")
        m3.metric("Periodic NPV", f"{sym}{npv_val:,.2f}")
        m4.metric("Periodic IRR", f"{irr_val * 100:.2f}%" if irr_val is not None else "N/A")

        t_col, c_col = st.columns([1, 1.5], gap="large")
        with t_col:
            st.markdown("#### Transaction Ledger")
            view_df = df_cf.copy()
            view_df["Date"] = view_df["Date"].dt.strftime("%Y-%m-%d")
            st.dataframe(view_df, use_container_width=True, height=300)

        with c_col:
            st.markdown("#### Cash Outflows vs. Inflows")
            fig_cf, ax_cf = plt.subplots(figsize=(7.5, 3.8))
            bar_colors = ["#d62728" if x < 0 else "#2ca02c" for x in flows]
            ax_cf.bar(view_df["Date"], flows, color=bar_colors, width=0.5)
            ax_cf.axhline(0, color="black", lw=1)
            ax_cf.set_ylabel(f"Cash Flow ({sym})")
            ax_cf.set_xlabel("Transaction Date")
            plt.xticks(rotation=25)
            ax_cf.grid(True, axis="y", ls=":", alpha=0.6)
            plt.tight_layout()
            st.pyplot(fig_cf)
    else:
        st.error("Invalid format: Ensure your Excel sheet contains exact column headers `Date` and `Cash_Flow`.")


# TAB 3: GOAL & WEALTH PLANNER (FV, PV, NPER)
with tab3:
    st.subheader("Time-Value-of-Money Wealth Planner")
    col_fv, col_pv, col_nper = st.columns(3, gap="medium")

    with col_fv:
        st.markdown("#### 1. Future Value (FV)")
        sip_amt = st.number_input(f"Monthly SIP ({sym})", value=10000, step=1000)
        init_pv = st.number_input(f"Initial Lump Sum ({sym})", value=50000, step=5000)
        fv_rate = st.slider("Expected Annual Return (%)", 1.0, 25.0, 12.0, 0.5, key="fv_r")
        fv_yrs = st.slider("Duration (Years)", 1, 40, 15, key="fv_y")

        fv_out = FinPulseEngine.fv((fv_rate / 100.0) / 12.0, fv_yrs * 12, -sip_amt, -init_pv)
        invested = init_pv + sip_amt * fv_yrs * 12
        st.success(f"**Future Corpus:** {sym}{fv_out:,.2f}")
        st.caption(f"Invested: {sym}{invested:,.0f} | Gain: {sym}{fv_out - invested:,.0f}")

    with col_pv:
        st.markdown("#### 2. Present Value (PV)")
        target_corpus = st.number_input(f"Target Goal Amount ({sym})", value=5000000, step=100000)
        pv_rate = st.slider("Annual Discount Rate (%)", 1.0, 20.0, 9.0, 0.5, key="pv_r")
        pv_yrs = st.slider("Years Until Goal", 1, 35, 10, key="pv_y")

        pv_out = -FinPulseEngine.pv(pv_rate / 100.0, pv_yrs, 0.0, target_corpus)
        st.info(f"**Invest Today (PV):** {sym}{pv_out:,.2f}")
        st.caption("Single lump sum required today to reach your target.")

    with col_nper:
        st.markdown("#### 3. Time to Goal (NPER)")
        goal_target = st.number_input(f"Target Wealth ({sym})", value=2500000, step=100000, key="np_fv")
        monthly_contrib = st.number_input(f"Monthly Contribution ({sym})", value=15000, step=1000, key="np_pmt")
        nper_rate = st.slider("Annual Return (%)", 1.0, 25.0, 11.0, 0.5, key="np_r")

        months_needed = FinPulseEngine.nper((nper_rate / 100.0) / 12.0, -monthly_contrib, 0.0, goal_target)
        if months_needed is not None:
            st.warning(f"**Time Required:** {months_needed / 12.0:.1f} Years")
            st.caption(f"Equivalent to {int(np.ceil(months_needed))} monthly installments.")
        else:
            st.error("Target unreachable with current inputs.")
