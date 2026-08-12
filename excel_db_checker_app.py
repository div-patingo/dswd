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
from streamlit_option_menu import option_menu

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
st.set_page_config(page_title="Excel Cross-Checker", layout="wide", page_icon="📊")

st.markdown("""
<style>
    #MainMenu, footer {visibility: hidden;}
    section[data-testid="stSidebar"] {
        background-color: #12141c;
        border-right: 1px solid #2a2d3a;
    }
    .app-title {
        font-size: 1.4rem;
        font-weight: 700;
        color: #fafafa;
        padding: 1.2rem 0 0.2rem 0;
        text-align: center;
    }
    .app-subtitle {
        font-size: 0.78rem;
        color: #8a8d98;
        text-align: center;
        padding-bottom: 1.2rem;
        border-bottom: 1px solid #2a2d3a;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div class="app-title">📊 Cross-Checker</div>', unsafe_allow_html=True)
    st.markdown('<div class="app-subtitle">Shared cloud database</div>', unsafe_allow_html=True)

    page = option_menu(
        menu_title=None,
        options=["Import Data", "Check a File", "View Database"],
        icons=["cloud-upload", "search", "table"],
        default_index=0,
        styles={
            "container": {"padding": "0", "background-color": "transparent"},
            "icon": {"color": "#f0a500", "font-size": "16px"},
            "nav-link": {
                "font-size": "14px",
                "text-align": "left",
                "margin": "4px 0",
                "border-radius": "8px",
                "color": "#c7c9d1",
                "padding": "10px 14px",
            },
            "nav-link-selected": {
                "background-color": "#f0a500",
                "color": "#12141c",
                "font-weight": "600",
            },
        }
    )

st.title("📊 Excel Cross-Checker")
st.caption("Shared cloud database - changes are visible to your whole team")


# ---- Page: Import ---------------------------------------------------------------
if page == "Import Data":
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

            st.markdown("**Match on separate name fields, weighted by importance**")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("*Uploaded file columns*")
                last_check = st.selectbox("Last Name column", df_check.columns, key="last_check")
                first_check = st.selectbox("First Name column", df_check.columns, key="first_check")
                middle_check = st.selectbox(
                    "Middle Name column (optional)",
                    ["(none)"] + list(df_check.columns), key="middle_check"
                )
            with col2:
                st.markdown("*Reference table columns*")
                last_ref = st.selectbox("Last Name column", df_ref.columns, key="last_ref")
                first_ref = st.selectbox("First Name column", df_ref.columns, key="first_ref")
                middle_ref = st.selectbox(
                    "Middle Name column (optional)",
                    ["(none)"] + list(df_ref.columns), key="middle_ref"
                )

            use_middle = middle_check != "(none)" and middle_ref != "(none)"

            st.markdown("**Field weights** (how much each field counts toward the final score)")
            w1, w2, w3 = st.columns(3)
            with w1:
                weight_last = st.slider("Last Name weight", 0, 100, 50)
            with w2:
                weight_first = st.slider("First Name weight", 0, 100, 30 if use_middle else 50)
            with w3:
                weight_middle = st.slider("Middle Name weight", 0, 100, 20 if use_middle else 0, disabled=not use_middle)

            weight_total = weight_last + weight_first + (weight_middle if use_middle else 0)
            if weight_total == 0:
                st.warning("Weights can't all be zero.")

            default_extra = ["PROJECT"] if "PROJECT" in df_ref.columns else []
            extra_cols = st.multiselect(
                "Additional column(s) from the database to include in the results",
                options=[c for c in df_ref.columns if c not in [last_ref, first_ref, middle_ref]],
                default=default_extra,
                help="e.g. Project, so you can see what each matched person is tied to"
            )

            threshold = st.slider("Partial match threshold (%)", 0, 100, 50)

            ready = weight_total > 0
            if ready and st.button("Run Cross-Check", type="primary"):
                df_ref_local = df_ref.copy()

                def norm(val):
                    return str(val).strip() if pd.notna(val) else ""

                results = []
                progress = st.progress(0, text="Comparing rows...")
                total = len(df_check)

                for i, (idx, row) in enumerate(df_check.iterrows()):
                    c_last = norm(row[last_check])
                    c_first = norm(row[first_check])
                    c_middle = norm(row[middle_check]) if use_middle else ""

                    display_name = " ".join(x for x in [c_last, c_first, c_middle] if x)

                    if not c_last and not c_first:
                        base_result = {
                            "Row (Uploaded File)": idx + 2,
                            "Uploaded_Name": display_name,
                            "Best Match (Last Name)": None,
                            "Best Match (First Name)": None,
                            "Match Score (%)": 0,
                            "Status": "No Match"
                        }
                        for c in extra_cols:
                            base_result[c] = None
                        results.append(base_result)
                    else:
                        best_score, best_idx_ref = -1, None
                        for idx_ref in df_ref_local.index:
                            r_last = norm(df_ref_local.at[idx_ref, last_ref])
                            r_first = norm(df_ref_local.at[idx_ref, first_ref])

                            score_last = fuzz.ratio(c_last, r_last)
                            score_first = fuzz.ratio(c_first, r_first)

                            weighted = (score_last * weight_last) + (score_first * weight_first)

                            if use_middle:
                                r_middle = norm(df_ref_local.at[idx_ref, middle_ref])
                                score_middle = fuzz.ratio(c_middle, r_middle)
                                weighted += score_middle * weight_middle

                            combined_score = weighted / weight_total

                            if combined_score > best_score:
                                best_score, best_idx_ref = combined_score, idx_ref

                        final_score = best_score if best_score >= 0 else 0
                        base_result = {
                            "Row (Uploaded File)": idx + 2,
                            "Uploaded_Name": display_name,
                            "Best Match (Last Name)": df_ref_local.at[best_idx_ref, last_ref] if best_idx_ref is not None else None,
                            "Best Match (First Name)": df_ref_local.at[best_idx_ref, first_ref] if best_idx_ref is not None else None,
                            "Match Score (%)": round(final_score, 1),
                            "Status": classify(final_score, threshold)
                        }
                        for c in extra_cols:
                            base_result[c] = df_ref_local.at[best_idx_ref, c] if best_idx_ref is not None else None
                        results.append(base_result)

                    if total:
                        progress.progress((i + 1) / total, text=f"Comparing rows... {i+1}/{total}")

                progress.empty()
                results_df = pd.DataFrame(results)
                st.session_state["results_df"] = results_df

            if "results_df" in st.session_state:
                results_df = st.session_state["results_df"]

                st.subheader("Results")

                filter_col1, filter_col2 = st.columns([2, 1])
                with filter_col1:
                    status_filter = st.multiselect(
                        "Filter by status",
                        options=["Exact Match", "Partial Match", "No Match"],
                        default=["Exact Match", "Partial Match", "No Match"]
                    )
                with filter_col2:
                    min_score = st.slider("Minimum match score (%)", 0, 100, 0)

                filtered_df = results_df[
                    results_df["Status"].isin(status_filter) &
                    (results_df["Match Score (%)"] >= min_score)
                ]

                summary = results_df["Status"].value_counts()
                cols = st.columns(len(summary))
                for c, (status, count) in zip(cols, summary.items()):
                    c.metric(status, count)

                st.caption(f"Showing {len(filtered_df):,} of {len(results_df):,} rows")

                def highlight_status(row):
                    color = {
                        "Exact Match": "background-color: #1b4332; color: #d8f3dc",
                        "Partial Match": "background-color: #4a3b00; color: #ffe8a3",
                        "No Match": "background-color: #4a1420; color: #ffd6de",
                    }.get(row["Status"], "")
                    return [color] * len(row)

                st.dataframe(filtered_df.style.apply(highlight_status, axis=1), use_container_width=True)

                st.download_button(
                    "Download filtered results as Excel",
                    data=to_excel_bytes(filtered_df),
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
