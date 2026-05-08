import streamlit as st
import pandas as pd
import io

# ===================================
# PAGE CONFIG
# ===================================
st.set_page_config(
    page_title="Nexio KPI Dashboard",
    page_icon="📊",
    layout="wide"
)

# ===================================
# KPI CONFIGURATION
# ===================================
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
            "QCSR not a customer requirement",
            "Customer Declined / Postponed QCSR"
        ]
    },
    "QCSR_PrepWeekPrior": {
        "question": "Physical preparatory meeting conducted a week prior to the QCSR",
        "valid": ["Yes", "QCSR not Scheduled for the current Month"]
    },
    "QCSR_Minutes2Days": {
        "question": "Minutes circulated within two (2) business days from the date of the QCRS",
        "valid": ["Yes", "QCRS not scheduled for current month"]
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
# TITLE
# ===================================
st.title("📊 Nexio KPI Analytics Dashboard")
st.write("Upload the KPI Excel file and analyze monthly KPI performance.")

# ===================================
# FILE UPLOAD
# ===================================
uploaded_file = st.file_uploader(
    "Upload KPI Excel File",
    type=["xlsx"]
)

# ===================================
# PROCESS FILE
# ===================================
if uploaded_file:

    try:
        df = pd.read_excel(
            uploaded_file,
            sheet_name="tblNexioKPI"
        )

    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    # ===================================
    # CLEAN KPIMonth
    # ===================================
    if "KPIMonth" not in df.columns:
        st.error("KPIMonth column not found.")
        st.stop()

    # Normalize KPIMonth
    df["KPIMonth"] = (
        df["KPIMonth"]
        .astype(str)
        .str.strip()
    )

    # ===================================
    # REMOVE DUPLICATES
    # ===================================
    duplicate_count = df.duplicated().sum()

    if duplicate_count > 0:
        st.warning(f"⚠ {duplicate_count} duplicate rows detected and removed.")
        df = df.drop_duplicates()

    # ===================================
    # MONTH SELECTION
    # ===================================
    months = sorted(df["KPIMonth"].dropna().unique())

    selected_month = st.selectbox(
        "📅 Select KPI Month",
        months
    )

    # ===================================
    # FILTER MONTH DATA
    # ===================================
    df_month = df[
        df["KPIMonth"] == selected_month
    ]

    if df_month.empty:
        st.warning("No KPI data found.")
        st.stop()

    st.success(f"Loaded KPI data for: {selected_month}")

    # ===================================
    # KPI ANALYSIS
    # ===================================
    st.subheader("📌 KPI Performance Breakdown")

    export_rows = []

    grand_correct = 0
    grand_total = 0

    for col, config in kpi_config.items():

        question = config["question"]
        valid_values = config["valid"]

        with st.expander(f"🔹 {question}"):

            if col not in df_month.columns:
                st.error(f"Column missing: {col}")
                continue

            # ===================================
            # CLEAN DATA
            # ===================================
            normalized_series = (
                df_month[col]
                .dropna()
                .astype(str)
                .str.strip()
                .str.replace(r'\s+', ' ', regex=True)
                .str.lower()
            )

            # ===================================
            # CLEAN VALID VALUES
            # ===================================
            normalized_valid = [
                v.lower().strip()
                for v in valid_values
            ]

            # ===================================
            # CALCULATE KPI
            # ===================================
            correct = normalized_series.isin(
                normalized_valid
            ).sum()

            total = len(normalized_series)

            percent = (
                round((correct / total) * 100, 2)
                if total > 0 else 0
            )

            # ===================================
            # OVERALL TOTALS
            # ===================================
            grand_correct += correct
            grand_total += total

            # ===================================
            # DISPLAY KPI
            # ===================================
            st.metric(
                label="KPI Score",
                value=f"{percent}%"
            )

            st.progress(percent / 100)

            st.write(f"✅ Valid PASS Values: {valid_values}")
            st.write(f"✔ Correct Records: {correct}")
            st.write(f"📄 Total Records: {total}")

            # ===================================
            # DEBUGGING SECTION
            # ===================================
            invalid_entries = normalized_series[
                ~normalized_series.isin(normalized_valid)
            ]

            if len(invalid_entries) > 0:

                with st.expander("⚠ View Invalid Entries"):

                    invalid_df = pd.DataFrame({
                        "Invalid Values": invalid_entries
                    })

                    st.dataframe(
                        invalid_df,
                        use_container_width=True
                    )

            # ===================================
            # EXPORT ROWS
            # ===================================
            export_rows.append({
                "KPI Question": question,
                "Score (%)": percent,
                "Correct": correct,
                "Total": total
            })

    # ===================================
    # OVERALL KPI
    # ===================================
    overall = (
        round((grand_correct / grand_total) * 100, 2)
        if grand_total > 0 else 0
    )

    st.divider()

    st.subheader("⭐ Overall KPI Performance")

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "Overall KPI Score",
            f"{overall}%"
        )

    with col2:
        st.metric(
            "Total KPI Records",
            grand_total
        )

    st.progress(overall / 100)

    # ===================================
    # KPI SUMMARY TABLE
    # ===================================
    st.subheader("📋 KPI Summary Table")

    export_df = pd.DataFrame(export_rows)

    st.dataframe(
        export_df,
        use_container_width=True
    )

    # ===================================
    # EXPORT TO EXCEL
    # ===================================
    st.subheader("📁 Export KPI Results")

    export_df.loc[len(export_df.index)] = [
        "OVERALL KPI SCORE",
        overall,
        grand_correct,
        grand_total
    ]

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        export_df.to_excel(
            writer,
            index=False,
            sheet_name="KPI Results"
        )

    st.download_button(
        label="⬇ Download KPI Results Excel",
        data=output.getvalue(),
        file_name=f"KPI_Results_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload your KPI Excel file.")
