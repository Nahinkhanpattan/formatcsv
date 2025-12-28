import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Event Data Formatter", layout="centered")
st.title("📄 Event Data – CSV to Excel Formatter")

uploaded_file = st.file_uploader("Upload total-part-errorErase.csv", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.success("CSV loaded successfully!")

    # ---- FIXED COLUMN NAMES (FROM YOUR CSV) ----
    team_col = "team_name"
    t1_col = "teammate1_name"
    t2_col = "teammate2_name"
    college_col = "college_name"
    phone_col = "phone_number"
    email_col = "email"

    required_cols = [team_col, t1_col, college_col, phone_col, email_col]
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"Missing columns in CSV: {missing}")
        st.stop()

    # ---- BUILD FINAL ROWS ----
    rows = []

    for _, row in df.iterrows():
        students = [row[t1_col]]

        if t2_col in df.columns and pd.notna(row[t2_col]) and row[t2_col] != "":
            students.append(row[t2_col])

        for idx, student in enumerate(students):
            record = {
                team_col: row[team_col],
                "Student Name": student
            }

            # TEAM-LEVEL FIELDS ONLY ON FIRST STUDENT
            if idx == 0:
                record[college_col] = row[college_col]
                record[phone_col] = str(row[phone_col]).replace(".0", "").strip()
                record[email_col] = row[email_col]
            else:
                record[college_col] = ""
                record[phone_col] = ""
                record[email_col] = ""

            rows.append(record)

    final_df = pd.DataFrame(rows)

    # ---- WRITE EXCEL WITH MERGES ----
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Event Data")
        ws = writer.sheets["Event Data"]

        def merge(col):
            idx = final_df.columns.get_loc(col)
            start = 1  # Excel row (after header)

            while start <= len(final_df):
                team = final_df.iloc[start - 1][team_col]
                end = start

                while end <= len(final_df) and final_df.iloc[end - 1][team_col] == team:
                    end += 1

                if end - start > 1:
                    ws.merge_range(start, idx, end - 1, idx, final_df.iloc[start - 1][col])

                start = end

        # ---- APPLY MERGES ----
        merge(team_col)
        merge(college_col)
        merge(phone_col)
        merge(email_col)

    st.download_button(
        "⬇️ Download formatted Excel",
        output.getvalue(),
        "event_data_final.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
