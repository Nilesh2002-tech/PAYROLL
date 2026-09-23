import pandas as pd
import streamlit as st

from payroll import Config, compute, read_workbook, slip_pdf

st.set_page_config(page_title="Attendance & Salary Console", layout="wide")
st.title("Attendance & Salary Console")

with st.sidebar:
    st.header("Settings")
    company = st.text_input("Company name", "Square One Communications")
    cfg = Config(
        late_allowed=st.number_input("Late marks allowed", 0, 31, 4),
        early_allowed=st.number_input("Early leaves allowed", 0, 31, 2),
        basis={"Calendar days": "calendar", "30 days": "30", "Working days": "working"}[
            st.selectbox("Per-day rate basis", ["Calendar days", "30 days", "Working days"])],
        deduct_absent=st.checkbox("Deduct absent days (1 day each)", True),
    )

up = st.file_uploader("Upload monthly workbook (punch sheet + 'Salary Data' sheet)", type=["xlsx", "xls"])
if not up:
    st.info("Upload the month's workbook to generate results and salary slips.")
    st.stop()

try:
    res = compute(*read_workbook(up), cfg)
except Exception as ex:
    st.error(f"Could not process file: {ex}")
    st.stop()

rows = res["rows"]
df = pd.DataFrame(rows).drop(columns="_ded")
c = st.columns(5)
c[0].metric("Pay period", res["period"])
c[1].metric("Working days", res["work_days"])
c[2].metric("Deductions", f"₹{df.Deductions.sum():,.0f}")
c[3].metric("Overtime", f"₹{df.Overtime.sum():,.0f}")
c[4].metric("Net payable", f"₹{df.Net.sum():,.0f}")

show = df.drop(columns=["Rate"])
money = ["Salary", "Deductions", "Overtime", "Net"]
show[money] = show[money].round(0).astype(int)
st.dataframe(show, use_container_width=True, hide_index=True)
st.caption(f"Per-day rate = salary ÷ {res['basis_days']}. Each late/early mark beyond the allowance and each half day costs 0.5 day.")

a, b, _ = st.columns([1, 1, 4])
tag = res["period"].replace(" ", "_")
a.download_button("All slips (PDF)", slip_pdf(rows, res["period"], res["work_days"], company), f"Salary_Slips_{tag}.pdf")
b.download_button("Export CSV", show.to_csv(index=False), f"Salary_Summary_{tag}.csv")

st.subheader("Employee detail & slip")
sel = st.selectbox("Employee", rows, format_func=lambda r: f"{r['ID']} – {r['Employee']}")
st.dataframe(res["details"][sel["ID"]], use_container_width=True, hide_index=True)
st.download_button(f"Download slip – {sel['Employee']}", slip_pdf([sel], res["period"], res["work_days"], company),
                   f"Salary_Slip_{sel['ID']}_{tag}.pdf", key="one")