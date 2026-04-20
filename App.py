import streamlit as st
import pandas as pd
import io


kpi_config = {
    "SMR_Submitted_6to8": {
        "question": "Service Management Report submitted between 6th and 8th day of the immediately following month",
        "valid": ["Yes"]
    },
    "MSM_Conducted": {
        "question": "Monthly Service Meeting conducted before end of the month",
        "valid": ["Yes", "Customer cancelled meeting"]
    },
    "Minutes_Within2Days": {
        "question": "Minutes circulated within two (2) business days from the date of the Service Management meeting",
        "valid": ["Yes", "Customer cancelled Meeting"]
    },
    "Docs_Saved_5Days": {
        "question": "Meeting Minutes and/or Monthly Report saved on the Vodacom SharePoint Site within 5 days of meeting conclusion",
        "valid": ["Yes"]
    },
    "QCSR_Conducted": {
        "question": "Quarterly Customer Service Review (QCSR) conducted",
        "valid": [
            "During the current month",
            "During the previous 3 Months",
            "Scheduled in the next 2 months",
            "Customer Declined/Postponed QCSR",
            "QCSR not a customer requirement"
        ]
    },
    "QCSR_PrepWeekPrior": {
        "question": "Physical preparatory meeting conducted a week prior to the QCSR",
        "valid": ["Yes", "QCSR not Scheduled for the current Month"]
    },
    "QCSR_Minutes2Days": {
        "question": "Minutes circulated within two (2) business days from the date of the QCRS",
        "valid": ["Yes", "QCSR not scheduled for the current month"]
    },
    "QCSR_DocsSaved5Days": {
        "question": "Documents saved on the Vodacom SharePoint Site within 5 days of QCRS",
        "valid": ["Yes", "QCSR not scheduled for the current Month"]
    },
    "WeeklyReport_SentByTue": {
        "question": "Weekly Report forwarded electronically to the customer by no later than the Tuesday immediately following the end of the week",
        "valid": ["Yes", "Not a customer requirement"]
    },
    "SIPs_Updated": {
        "question": "SIPS initiated and updated as indicated by the process requirements",
        "valid": ["Yes", "No Missed SLA"]
    },
    "CSIR_Prepared_OnTime": {
        "question": "Customer Specific Incident Report (CSIR) prepared and distributed 72 calendar hours or 24 business hours",
        "valid": ["Yes", "No Incidents Reports for the month"]
    },
    "CSIR_Meeting_5Days": {
        "question": "Meeting conducted within 5 business days after CSIR release",
        "valid": ["Yes", "No Incidents Reports for the month"]
    }
}

# ===================================
# STREAMLIT APP LAYOUT
# ===================================
st.set_page_config(page_title="📊 Nexio KPI Dashboard", layout="wide")

st.title("📊 Nexio KPI Analytics Dashboard")
st.write("Upload the KPI Excel file and select the KPI Month to view performance details.")

uploaded_file = st.file_uploader("Upload tblNexioKPI Excel File", type=["xlsx"])

# ===================================
# PROCESS EXCEL FILE
# ===================================
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="tblNexioKPI")

    with st.expander("📅 Select KPI Month"):
        months = sorted(df["KPIMonth"].dropna().unique())
        selected_month = st.selectbox("KPI Month", months)

    df_month = df[df["KPIMonth"] == selected_month]

    if df_month.empty:
        st.warning("No KPI data found for the selected month.")
    else:
        st.success(f"KPI data loaded for **{selected_month}**")

        st.subheader("KPI Performance Breakdown (Full Question Descriptions)")

        total_score = 0
        question_count = len(kpi_config)
        export_rows = []  # For Excel export

        for col, config in kpi_config.items():
            long_question = config["question"]
            valid_values = config["valid"]

            with st.expander(f"🔹 {long_question}"):

                if col in df_month.columns:

                    # ==============================
                    # CASE-INSENSITIVE MATCHING
                    # ==============================
                    normalized_series = df_month[col].astype(str).str.strip().str.lower()
                    normalized_valid = [v.lower() for v in valid_values]

                    correct = normalized_series.isin(normalized_valid).sum()
                    total = normalized_series.count()

                    percent = round((correct / total) * 100, 2) if total > 0 else 0
                    total_score += percent

                    st.metric("KPI Score", f"{percent}%")
                    st.write(f"**Valid PASS values:** {valid_values}")
                    st.write(f"Correct: {correct} / {total}")

                    export_rows.append({
                        "KPI Question": long_question,
                        "Score (%)": percent,
                        "Correct": correct,
                        "Total": total
                    })

                else:
                    st.error(f"Column not found: {col}")

        # ==============================
        # OVERALL KPI SCORE
        # ==============================
        overall = round(total_score / question_count, 2)

        st.subheader("⭐ Overall KPI Performance")
        st.metric(label=f"Overall KPI Score for {selected_month}", value=f"{overall}%")

        # ==============================
        # EXPORT TO EXCEL
        # ==============================
        st.subheader("📁 Export KPI Results to Excel")

        export_df = pd.DataFrame(export_rows)
        export_df.loc[len(export_df.index)] = ["Overall Score", overall, "", ""]

        towrite = io.BytesIO()

        # Use openpyxl so NO installation required
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name="KPI Results")

        st.download_button(
            label="⬇ Download KPI Results Excel",
            data=towrite.getvalue(),
            file_name=f"KPI_Results_{selected_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload your Excel file (tblNexioKPI).")
