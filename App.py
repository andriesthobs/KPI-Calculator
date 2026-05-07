import streamlit as st
import pandas as pd
import io

# ===================================
# KPI CONFIGURATION (STANDARDISED)
# ===================================
kpi_config = {
    "SMR_Submitted_6to8": {
        "question": "Service Management Report submitted between 6th and 8th day of the immediately following month",
        "valid": ["yes"]
    },
    "MSM_Conducted": {
        "question": "Monthly Service Meeting conducted before end of the month",
        "valid": ["yes", "cancelled"]
    },
    "Minutes_Within2Days": {
        "question": "Minutes circulated within two (2) business days from the date of the Service Management meeting",
        "valid": ["yes", "cancelled"]
    },
    "Docs_Saved_5Days": {
        "question": "Meeting Minutes and/or Monthly Report saved within 5 days",
        "valid": ["yes"]
    }
}

# ===================================
# NORMALISATION FUNCTION (CRITICAL FIX)
# ===================================
def normalize_value(val):
    val = str(val).strip().lower()

    if val in ["yes"]:
        return "yes"

    if val in ["no"]:
        return "no"

    if val in [
        "customer cancelled meeting",
        "customer cancelled meeting ",
        "customer cancelled meeting",
        "meeting cancelled",
        "cancelled",
        "customer did not attend",
        "no meeting held",
        "did not join"
    ]:
        return "cancelled"

    if val in [
        "not a customer requirement",
        "not required",
        "no requirement"
    ]:
        return "not_required"

    if val in ["", "nan", "none"]:
        return "blank"

    return "other"


# ===================================
# VALID POPULATION FILTER
# ===================================
def get_valid_population(series):
    return series[series.isin(["yes", "no", "cancelled", "not_required"])]


# ===================================
# STREAMLIT UI
# ===================================
st.set_page_config(page_title="📊 Nexio KPI Dashboard", layout="wide")

st.title("📊 Nexio KPI Analytics Dashboard")

uploaded_file = st.file_uploader("Upload KPI Excel File", type=["xlsx"])

# ===================================
# PROCESS FILE
# ===================================
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="tblNexioKPI", engine="openpyxl")

    months = sorted(df["KPIMonth"].dropna().unique())
    selected_month = st.selectbox("Select KPI Month", months)

    df_month = df[df["KPIMonth"] == selected_month]

    # ===================================
    # KPI CALCULATION
    # ===================================
    total_correct = 0
    total_possible = 0
    export_rows = []

    for col, config in kpi_config.items():

        with st.expander(f"🔹 {config['question']}"):

            if col in df_month.columns:

                # ✅ Normalize column values
                normalized = df_month[col].apply(normalize_value)

                # ✅ Filter valid population
                valid_population = get_valid_population(normalized)

                total = len(valid_population)

                # ✅ Count correct
                correct = valid_population.isin(config["valid"]).sum()

                percent = round((correct / total) * 100, 2) if total > 0 else 0

                # ✅ Aggregate totals (FIXED!)
                total_correct += correct
                total_possible += total

                # ✅ UI display
                st.metric("KPI Score", f"{percent}%")
                st.write(f"Valid PASS values: {config['valid']}")
                st.write(f"Correct: {correct} / {total}")

                # ✅ Store export data
                export_rows.append({
                    "KPI Question": config["question"],
                    "Score (%)": percent,
                    "Correct": correct,
                    "Total": total
                })

            else:
                st.error(f"Missing column: {col}")

    # ===================================
    # ✅ CORRECT OVERALL KPI (WEIGHTED)
    # ===================================
    overall = round((total_correct / total_possible) * 100, 2) if total_possible > 0 else 0

    st.subheader("⭐ Overall KPI Performance")
    st.metric("Overall KPI Score", f"{overall}%")
    st.write(f"Overall: {total_correct} / {total_possible}")

    # ===================================
    # EXPORT
    # ===================================
    export_df = pd.DataFrame(export_rows)
    export_df.loc[len(export_df.index)] = ["Overall Score", overall, total_correct, total_possible]

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False)

    st.download_button(
        label="⬇ Download KPI Results",
        data=buffer.getvalue(),
        file_name=f"KPI_Results_{selected_month}.xlsx"
    )

else:
    st.info("Upload your KPI Excel file to begin.")
