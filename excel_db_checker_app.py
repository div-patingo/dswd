#!/usr/bin/env python3
"""
Excel Cross-Checker - Cloud Web UI (Streamlit + Postgres)
-------------------------------------------------------------
Same app as before, but uses a cloud Postgres database (e.g. Supabase's
free tier) instead of a local SQLite file, so it works when deployed to
Streamlit Community Cloud and data persists + is shared across users.

LOCAL TESTING:
    You can still test this on your own machine by setting DB_URL in
    .streamlit/secrets.toml (see bottom of this file for format).

DEPLOYMENT:
    See the setup guide provided alongside this file.
"""

import re
import io
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
from sqlalchemy import create_engine, text

# ---- Database connection -----------------------------------------------------
@st.cache_resource
def get_engine():
    db_url = st.secrets.get("DB_URL")
    if not db_url:
        st.error(
            "No database connection configured. Add DB_URL to your Streamlit "
            "secrets (Settings -> Secrets on Streamlit Cloud, or "
            ".streamlit/secrets.toml locally). See the setup guide."
        )
        st.stop()
    return create_engine(db_url, pool_pre_ping=True)


engine = get_engine()


# ---- Helpers ----------------------------------------------------------------
def list_tables():
    with engine.connect() as conn:
        result = conn.execute(text(
            "SELECT table_name FROM information_schema.tables "
            "WHERE table_schema = 'public';"
        ))
        return [row[0] for row in result]


def table_row_count(table_name):
    with engine.connect() as conn:
        result = conn.execute(text(f'SELECT COUNT(*) FROM "{table_name}"'))
        return result.scalar()


def sanitize_table_name(name):
    name = re.sub(r"[^A-Za-z0-9_]", "_", name).lower()
    if name and name[0].isdigit():
        name = f"t_{name}"
    return name or "reference_data"


def build_key(row, cols, join_with_space=True):
    sep = " " if join_with_space else ""
    return sep.join(str(row[c]).strip() for c in cols if pd.notna(row[c]))


def classify(score, threshold):
    if score == 100:
        return "Exact Match"
    elif score >= threshold:
        return "Partial Match"
    else:
        return "No Match"


def to_excel_bytes(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
    return buffer.getvalue()


# ---- Page setup ---------------------------------------------------------------
st.set_page_config(page_title="Excel Cross-Checker", layout="wide")
st.title("📊 Excel Cross-Checker")
st.caption("Shared cloud database - changes are visible to your whole team")

page = st.sidebar.radio(
    "Menu",
    ["Import / Update Reference Data", "Check a File", "View Database"]
)


# ---- Page: Import ---------------------------------------------------------------
if page == "Import / Update Reference Data":
    st.header("Import / Update Reference Data")

    uploaded = st.file_uploader("Upload an Excel file to import", type=["xlsx", "xls"])
    if uploaded:
        df = pd.read_excel(uploaded)
        st.success(f"Loaded {len(df):,} rows, {len(df.columns)} columns")
        st.dataframe(df.head(10), use_container_width=True)

        existing = list_tables()
        default_name = sanitize_table_name(Path(uploaded.name).stem)
        table_name = st.text_input("Table name to store this data as", value=default_name)
        table_name = sanitize_table_name(table_name)

        mode_label = "Replace"
        if table_name in existing:
            st.warning(f"Table '{table_name}' already exists ({table_row_count(table_name):,} rows).")
            mode_label = st.radio("What should happen?", ["Append new rows", "Replace entirely"])

        if st.button("Import into database", type="primary"):
            mode = "append" if mode_label == "Append new rows" else "replace"
            with st.spinner("Uploading to database..."):
                df.to_sql(table_name, engine, if_exists=mode, index=False)
            st.success(f"Done! Table '{table_name}' now has {table_row_count(table_name):,} rows.")
            st.rerun()


# ---- Page: Check a File ---------------------------------------------------------
elif page == "Check a File":
    st.header("Check a File Against the Database")

    tables = list_tables()
    if not tables:
        st.info("Database is empty. Go to 'Import / Update Reference Data' first.")
    else:
        table_name = st.selectbox("Reference table to check against", tables)
        df_ref = pd.read_sql(f'SELECT * FROM "{table_name}"', engine)
        st.caption(f"{len(df_ref):,} rows in this table")

        uploaded = st.file_uploader("Upload the file that needs checking", type=["xlsx", "xls"])
        if uploaded:
            excel_file = pd.ExcelFile(uploaded)
            sheet_names = excel_file.sheet_names

            if len(sheet_names) > 1:
                sheet_choice = st.selectbox("This file has multiple tabs — which one do you want to check?", sheet_names)
            else:
                sheet_choice = sheet_names[0]

            df_check = pd.read_excel(excel_file, sheet_name=sheet_choice)
            st.success(f"Loaded {len(df_check):,} rows from tab '{sheet_choice}'")
            st.dataframe(df_check.head(5), use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                cols_check = st.multiselect(
                    "Column(s) from uploaded file to compare",
                    options=list(df_check.columns),
                    help="e.g. select Last Name, First Name, Middle Name together"
                )
                join_check_no_space = st.checkbox(
                    "Join these without spaces (e.g. LastFirstMiddle)",
                    key="join_check"
                )
            with col2:
                cols_ref = st.multiselect(
                    "Column(s) from reference table to compare",
                    options=list(df_ref.columns)
                )
                join_ref_no_space = st.checkbox(
                    "Join these without spaces",
                    key="join_ref"
                )

            threshold = st.slider("Partial match threshold (%)", 0, 100, 50)

            if cols_check and cols_ref and st.button("Run Cross-Check", type="primary"):
                df_ref_local = df_ref.copy()
                df_ref_local["_match_key"] = df_ref_local.apply(
                    lambda r: build_key(r, cols_ref, join_with_space=not join_ref_no_space), axis=1
                )

                results = []
                progress = st.progress(0, text="Comparing rows...")
                total = len(df_check)

                for i, (idx, row) in enumerate(df_check.iterrows()):
                    key1 = build_key(row, cols_check, join_with_space=not join_check_no_space)

                    if not key1.strip():
                        results.append({
                            "Row (Uploaded File)": idx + 2,
                            "Uploaded_Value": key1,
                            "Best Match (Database)": None,
                            "Match Score (%)": 0,
                            "Status": "No Match"
                        })
                    else:
                        best_score, best_key_ref = -1, None
                        for key_ref in df_ref_local["_match_key"]:
                            score = fuzz.token_sort_ratio(key1, key_ref)
                            if score > best_score:
                                best_score, best_key_ref = score, key_ref

                        final_score = best_score if best_score >= 0 else 0
                        results.append({
                            "Row (Uploaded File)": idx + 2,
                            "Uploaded_Value": key1,
                            "Best Match (Database)": best_key_ref,
                            "Match Score (%)": round(final_score, 1),
                            "Status": classify(final_score, threshold)
                        })

                    if total:
                        progress.progress((i + 1) / total, text=f"Comparing rows... {i+1}/{total}")

                progress.empty()
                results_df = pd.DataFrame(results)

                st.subheader("Results")
                summary = results_df["Status"].value_counts()
                cols = st.columns(len(summary))
                for c, (status, count) in zip(cols, summary.items()):
                    c.metric(status, count)

                def highlight_status(row):
                    color = {
                        "Exact Match": "background-color: #1b4332; color: #d8f3dc",
                        "Partial Match": "background-color: #4a3b00; color: #ffe8a3",
                        "No Match": "background-color: #4a1420; color: #ffd6de",
                    }.get(row["Status"], "")
                    return [color] * len(row)

                st.dataframe(results_df.style.apply(highlight_status, axis=1), use_container_width=True)

                st.download_button(
                    "Download results as Excel",
                    data=to_excel_bytes(results_df),
                    file_name="cross_check_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


# ---- Page: View Database ---------------------------------------------------------
elif page == "View Database":
    st.header("Database Contents")

    tables = list_tables()
    if not tables:
        st.info("Database is empty. Go to 'Import / Update Reference Data' first.")
    else:
        table_name = st.selectbox("Table to view", tables)
        total_rows = table_row_count(table_name)
        st.caption(f"{total_rows:,} total rows")

        search_term = st.text_input(
            "Search (matches any column, case-insensitive)",
            placeholder="e.g. a name, barangay, or project"
        )

        @st.cache_data(ttl=300)
        def load_full_table(name):
            return pd.read_sql(f'SELECT * FROM "{name}"', engine)

        df_full = load_full_table(table_name)

        if search_term:
            mask = df_full.apply(
                lambda col: col.astype(str).str.contains(search_term, case=False, na=False)
            ).any(axis=1)
            df_display = df_full[mask]
            st.caption(f"{len(df_display):,} rows match '{search_term}'")
        else:
            df_display = df_full

        st.dataframe(df_display, use_container_width=True, height=600)

        st.download_button(
            "Download this view as Excel",
            data=to_excel_bytes(df_display),
            file_name=f"{table_name}_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------------------------------------------------------------------
# secrets.toml format (for local testing only - on Streamlit Cloud this goes
# in the app's Settings -> Secrets instead, as TOML text):
#
# DB_URL = "postgresql://postgres:[YOUR-PASSWORD]@[YOUR-HOST]:5432/postgres"
# ---------------------------------------------------------------------------
