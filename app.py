import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CSV → Excel Formatter", layout="centered")

st.title("📄 CSV to Formatted Excel Converter")
st.write("Upload a CSV file and get a properly merged Excel output.")

uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

def find_col(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.success("CSV loaded successfully!")

    # -------- REQUIRED COLUMNS --------
    team_col = "team_name"
    t1_col = "teammate1_name"
    t2_col = "teammate2_name"

    if team_col not in df.columns or t1_col not in df.columns:
        st.error("CSV must contain: team_name, teammate1_name")
        st.stop()

    # -------- OPTIONAL TEAM-LEVEL COLUMNS --------
    college_col = find_col(df, ["college", "college_name"])
    event_col   = find_col(df, ["event", "event_name"])
    email_col   = find_col(df, ["email", "team_email"])
    phone_col   = find_col(df, ["phone", "mobile", "contact_number"])

    # -------- EXPAND STUDENTS --------
    rows = []
    base_cols = [c for c in df.columns if c not in [t1_col, t2_col]]

    for _, row in df.iterrows():
        base = row[base_cols].to_dict()

        if pd.notna(row[t1_col]):
            r = base.copy()
            r["Student Name"] = row[t1_col]
            rows.append(r)

        if t2_col in df.columns and pd.notna(row[t2_col]):
            r = base.copy()
            r["Student Name"] = row[t2_col]
            rows.append(r)

    final_df = pd.DataFrame(rows)

    # -------- NORMALIZE PHONE (IMPORTANT) --------
    if phone_col:
        final_df[phone_col] = (
            final_df[phone_col]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )

    # -------- COLUMN ORDER --------
    cols = list(final_df.columns)
    cols.remove(team_col)
    cols.remove("Student Name")
    final_df = final_df[[team_col, "Student Name"] + cols]

    # -------- WRITE EXCEL --------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Teams")
        ws = writer.sheets["Teams"]

        def merge_team_level(col_name):
            if col_name is None or col_name not in final_df.columns:
                return

            col_idx = final_df.columns.get_loc(col_name)
            start = 1

            while start <= len(final_df):
                team = final_df.iloc[start - 1][team_col]
                end = start

                while end <= len(final_df) and final_df.iloc[end - 1][team_col] == team:
                    end += 1

                # MERGE ONLY IF TEAM HAS MULTIPLE STUDENTS
                if end - start > 1:
                    value = final_df.iloc[start - 1][col_name]

                    if pd.notna(value):
                        # Clear duplicate cells
                        for r in range(start + 1, end):
                            ws.write_blank(r, col_idx, None)

                        ws.merge_range(start, col_idx, end - 1, col_idx, value)

                start = end

        # -------- APPLY MERGES --------
        merge_team_level(team_col)
        merge_team_level(college_col)
        merge_team_level(event_col)
        merge_team_level(email_col)   # TEAM-LEVEL
        merge_team_level(phone_col)   # TEAM-LEVEL

    st.success("Excel file generated successfully!")

    st.download_button(
        label="⬇️ Download Formatted Excel",
        data=output.getvalue(),
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
