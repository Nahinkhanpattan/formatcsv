import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CSV → Excel Formatter", layout="centered")
st.title("📄 CSV to Formatted Excel Converter")

uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

def find_col(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None

if uploaded_file:
    df = pd.read_csv(uploaded_file)

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

    # -------- EXPAND STUDENTS (1 ROW = 1 STUDENT) --------
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

    # -------- NORMALIZE PHONE --------
    if phone_col:
        final_df[phone_col] = (
            final_df[phone_col]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )

    # -------- REMOVE TEAM-LEVEL DUPLICATES (CRITICAL STEP) --------
    i = 0
    while i < len(final_df):
        team = final_df.loc[i, team_col]
        j = i

        while j < len(final_df) and final_df.loc[j, team_col] == team:
            j += 1

        # If team has multiple students, keep value ONLY in first row
        if j - i > 1:
            for col in [phone_col, email_col, college_col, event_col]:
                if col:
                    for r in range(i + 1, j):
                        final_df.loc[r, col] = ""

        i = j

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

        def merge(col):
            if not col:
                return

            idx = final_df.columns.get_loc(col)
            start = 1  # Excel row index (after header)

            while start <= len(final_df):
                team = final_df.iloc[start - 1][team_col]
                end = start

                while end <= len(final_df) and final_df.iloc[end - 1][team_col] == team:
                    end += 1

                # Merge whole team block
                if end - start > 1:
                    value = final_df.iloc[start - 1][col]
                    if value != "":
                        ws.merge_range(start, idx, end - 1, idx, value)

                start = end

        # -------- APPLY MERGES --------
        merge(team_col)
        merge(college_col)
        merge(event_col)
        merge(email_col)
        merge(phone_col)

    st.download_button(
        "⬇️ Download Formatted Excel",
        output.getvalue(),
        "formatted_output.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
