import streamlit as st
import pandas as pd
import io

# ===================================
# ✅ FULL KPI CONFIG (ALL KPIs)
# ===================================
kpi_config = {
    "SMR_Submitted_6to8": {
        "question": "Service Management Report submitted between 6th and 8th day",
        "valid": ["yes"]
    },
    "MSM_Conducted": {
        "question": "Monthly Service Meeting conducted before end of the month",
        "valid": ["yes", "cancelled"]
    },
    "Minutes_Within2Days": {
        "question": "Minutes circulated within 2 business days",
        "valid": ["yes", "cancelled"]
    },
    "Docs_Saved_5Days": {
        "question": "Documents saved within 5 days",
        "valid": ["yes"]
    },
    "QCSR_Conducted": {
        "question": "Quarterly Customer Service Review conducted",
        "valid": ["yes", "cancelled", "not_required"]
    },
    "QCSR_PrepWeekPrior": {
        "question": "Preparatory meeting conducted",
        "valid": ["yes", "not_required"]
    },
    "QCSR_Minutes2Days": {
        "question": "QCSR minutes within 2 days",
        "valid": ["yes", "not_required"]
    },
    "QCSR_DocsSaved5Days": {
        "question": "QCSR docs saved within 5 days",
        "valid": ["yes", "not_required"]
    },
    "WeeklyReport_SentByTue": {
        "question": "Weekly report sent by Tuesday",
        "valid": ["yes", "not_required"]
    },
    "SIPs_Updated": {
        "question": "SIPs updated",
        "valid": ["yes"]
    },
    "CSIR_Prepared_OnTime": {
        "question": "CSIR prepared on time",
        "valid": ["yes", "not_required"]
    },
    "CSIR_Meeting_5Days": {
        "question": "CSIR meeting within 5 days",
        "valid": ["yes", "not_required"]
    }
}

# ===================================
# ✅ NORMALISATION FUNCTION (ROBUST)
# ===================================
def normalize(val):
    val = str(val).strip().lower()

    if val in ["yes"]:
        return "yes"

    if val in ["no"]:
        return "no"

    if any(x in val for x in [
        "cancel", "did not attend", "no meeting", "not held"
    ]):
        return "cancelled"

    if any(x in val for x in [
        "not required", "no requirement", "not a requirement"
    ]):
        return "not_required"

    if val in ["nan", "", "none"]:
        return "blank"

    return "other"

# ===================================
# ✅ VALID POPULATION
# ===================================
VALID_SET = ["yes", "no", "cancelled", "not_required"]

# ===================================
# STREAMLIT
# ===================================
st.set_page_config(page_title="KPI Dashboard", layout="wide")
st.title("📊 Nexio KPI Dashboard")

file = st.file_uploader("Upload Excel", type=["xlsx"])

if file:
    df = pd.read_excel(file, sheet_name="tblNexioKPI", engine="openpyxl")

    month = st.selectbox("Select Month", sorted(df["KPIMonth"].dropna().unique()))
    df = df[df["KPIMonth"] == month]

    total_correct = 0
    total_possible = 0

    debug_data = []

    for col, cfg in kpi_config.items():

        st.subheader(cfg["question"])

        if col not in df.columns:
            st.error(f"Missing column: {col}")
            continue

        series = df[col].apply(normalize)

        # ✅ CLEAN population
        valid_pop = series[series.isin(VALID_SET)]

        total = len(valid_pop)

        # ✅ FIX: normalize VALID values ALSO
        valid_values = [v.lower() for v in cfg["valid"]]

        correct = valid_pop.isin(valid_values).sum()

        percent = round((correct / total) * 100, 2) if total else 0

        total_correct += correct
        total_possible += total

        st.write(f"Score: {percent}%")
        st.write(f"{correct} / {total}")

        # ✅ DEBUG (very important)
        debug_counts = series.value_counts()

        debug_data.append({
            "KPI": col,
            "Total Rows": len(series),
            "Valid Rows": total,
            "Breakdown": dict(debug_counts)
        })

    # ===================================
    # ✅ CORRECT OVERALL KPI
    # ===================================
    overall = round((total_correct / total_possible) * 100, 2) if total_possible else 0

    st.header("⭐ Overall KPI")
    st.metric("Overall Score", f"{overall}%")
    st.write(f"{total_correct} / {total_possible}")

    # ===================================
    # ✅ DEBUG VIEW (THIS WILL SHOW YOUR ISSUE CLEARLY)
    # ===================================
    with st.expander("🔍 Data Debug View"):
        st.write(pd.DataFrame(debug_data))

    # ===================================
    # EXPORT
    # ===================================
    output = io.BytesIO()

    pd.DataFrame(debug_data).to_excel(output, index=False)

    st.download_button("Download Debug Data", output.getvalue(), "debug.xlsx")
