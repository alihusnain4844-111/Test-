"""
EFERT x UBL — Dealer Credit Finance (DCF) Pipeline Tracker
============================================================
A Streamlit app to track EFERT-recommended dealers through UBL's DCF
credit approval process: current stage, EFERT / UBL points of contact,
and document completeness against the Program Policy Manual (PPM)
checklist (Section 6: Documentation required for credit approvals).

HOW TO RUN
----------
1. Save this file as app.py
2. pip install streamlit pandas openpyxl
3. streamlit run app.py
4. In the sidebar, upload:
     - EFERT_DCF_Recommended_Dealers.xlsx   (required — the dealer list)
     - EFERT_DCF_POC_List.xlsx              (optional — UBL/EFERT POC directory)
   If you don't upload them, the app looks for both files in the same
   folder as app.py using the default filenames above.

PERSISTENCE
------------
All tracking edits (stage, assigned UBL owner, document checkboxes,
notes) are saved to a local file called "dcf_tracker_state.csv" next to
this script, so your progress survives closing and reopening the app.
New dealers found in the source Excel file are automatically added;
dealers you've already started tracking keep their saved status.
"""

import os
import io
import json
import datetime as dt

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------

st.set_page_config(page_title="EFERT–UBL DCF Pipeline Tracker", layout="wide")

DEALER_FILE_DEFAULT = "EFERT_DCF_Recommended_Dealers.xlsx"
POC_FILE_DEFAULT = "EFERT_DCF_POC_List.xlsx"
TRACKER_STATE_FILE = "dcf_tracker_state.csv"

# Pipeline stages, in order, reflecting the DCF credit-approval workflow
# described in the PPM (recommendation -> documentation -> credit approval
# -> facility offer -> disbursement).
STAGES = [
    "1. Recommended by EFERT",
    "2. Document Collection In Progress",
    "3. Documents Complete – Pending Credit Proposal",
    "4. Credit Proposal Under Approval",
    "5. Approved – Pending Facility Offer Letter",
    "6. Facility Offer Letter Issued",
    "7. Disbursed / Active",
    "8. On Hold / Declined",
]

# Document checklist, transcribed from PPM Section 6 (Documentation
# required for credit approvals).
DEALER_DOCS = [
    "Customer (Dealer) Request for Finance",
    "Loan Application Form (LAF) / BBFS (per PR)",
    "3 Years Financials (Audited / Management)",
    "Bank Statement (min. 1 year)",
    "Valid Fertilizer Dealership Certificate",
]

CREDIT_APPROVAL_DOCS = [
    "Short Credit Proposal (paper/email)",
    "Director Search (if applicable)",
    "eCIB Report (sponsor/firm/group)",
    "EFERT Anchor Recommendation Letter (with 2-yr sales)",
    "Obligor Risk Rating (ORR)",
    "Call / Visit Report",
    "Eligibility Criteria Checklist",
]

CONSTITUTION_DOCS_INDIVIDUAL = ["CNIC / NTN / Partnership Deed"]
CONSTITUTION_DOCS_COMPANY = [
    "COI / MOA / AOA / Form A & 29 / Board Resolution / SECP Search Report"
]

COMPANY_ENTITY_TYPES = {"Company/Corporation", "Limited Company", "Company"}


def get_constitution_docs(entity_type: str):
    if str(entity_type).strip() in COMPANY_ENTITY_TYPES:
        return CONSTITUTION_DOCS_COMPANY
    return CONSTITUTION_DOCS_INDIVIDUAL


def all_required_docs(entity_type: str):
    return DEALER_DOCS + CREDIT_APPROVAL_DOCS + get_constitution_docs(entity_type)


# ----------------------------------------------------------------------------
# DATA LOADING
# ----------------------------------------------------------------------------

def to_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("-", "", "NIL"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        import re
        m = re.search(r"-?\d+\.?\d*", s)
        return float(m.group()) if m else 0.0


@st.cache_data(show_spinner=False)
def load_dealers(file_bytes: bytes) -> pd.DataFrame:
    """Parse the 'DCF list 2026' sheet of the Recommended Dealers workbook."""
    wb = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = "DCF list 2026" if "DCF list 2026" in wb.sheet_names else wb.sheet_names[0]
    raw = pd.read_excel(wb, sheet_name=sheet, header=None)

    # Header row is row index 16 (0-indexed) -> data starts at row 17
    records = []
    for _, row in raw.iloc[17:].iterrows():
        dealer_name = row.get(16)
        if pd.isna(dealer_name):
            continue
        records.append({
            "sno": row.get(0),
            "zone": row.get(1),
            "region": row.get(2),
            "efert_poc": str(row.get(3)).strip() if pd.notna(row.get(3)) else "",
            "warehouse": row.get(4),
            "category": row.get(6),
            "recommended_limit_mn": to_num(row.get(7)),
            "entity_type": row.get(9),
            "years_with_engro": round(to_num(row.get(10)), 1),
            "overdue_flag": row.get(11),
            "dealer_name": str(dealer_name).strip(),
            "address": row.get(17),
            "sales_2025": to_num(row.get(27)),
            "sales_2026": to_num(row.get(29)),
        })
    df = pd.DataFrame(records)
    df["dealer_key"] = (
        df["dealer_name"].str.upper().str.strip() + " | " + df["region"].astype(str)
    )
    return df


@st.cache_data(show_spinner=False)
def load_poc_directory(file_bytes: bytes) -> pd.DataFrame:
    """Parse the UBL/EFERT POC directory, keyed by Zone + Region."""
    wb = pd.ExcelFile(io.BytesIO(file_bytes))
    raw = pd.read_excel(wb, sheet_name=wb.sheet_names[0], header=None)
    records = []
    for _, row in raw.iloc[2:].iterrows():
        zone = row.get(0)
        if pd.isna(zone):
            continue
        records.append({
            "zone": zone,
            "region": row.get(1),
            "territory": row.get(2),
            "asm_name": row.get(3),
            "asm_contact": row.get(4),
            "rc_name": row.get(6),
            "rsm_name": row.get(9),
            "ubl_ch_rh": row.get(13),
        })
    return pd.DataFrame(records)


def load_workbook_bytes(uploaded_file, default_path):
    if uploaded_file is not None:
        return uploaded_file.read()
    if os.path.exists(default_path):
        with open(default_path, "rb") as f:
            return f.read()
    return None


# ----------------------------------------------------------------------------
# TRACKER STATE (persistence layer)
# ----------------------------------------------------------------------------

def doc_column(doc_name: str) -> str:
    return "doc__" + doc_name.replace(" ", "_").replace("/", "_")[:60]


def build_blank_tracker_row(dealer_row) -> dict:
    row = {
        "dealer_key": dealer_row["dealer_key"],
        "stage": STAGES[0],
        "ubl_assigned_to": "",
        "notes": "",
        "last_updated": dt.date.today().isoformat(),
    }
    for doc in all_required_docs(dealer_row.get("entity_type", "")):
        row[doc_column(doc)] = False
    return row


def load_tracker_state() -> pd.DataFrame:
    if os.path.exists(TRACKER_STATE_FILE):
        return pd.read_csv(TRACKER_STATE_FILE)
    return pd.DataFrame()


def save_tracker_state(df: pd.DataFrame):
    df.to_csv(TRACKER_STATE_FILE, index=False)


def sync_tracker_with_dealers(dealers_df: pd.DataFrame) -> pd.DataFrame:
    """Merge saved tracker state with the latest dealer list — keep edits
    for existing dealers, add blank tracker rows for new ones."""
    tracker_df = load_tracker_state()

    if tracker_df.empty:
        new_rows = [build_blank_tracker_row(r) for _, r in dealers_df.iterrows()]
        tracker_df = pd.DataFrame(new_rows)
    else:
        existing_keys = set(tracker_df["dealer_key"])
        new_rows = [
            build_blank_tracker_row(r)
            for _, r in dealers_df.iterrows()
            if r["dealer_key"] not in existing_keys
        ]
        if new_rows:
            tracker_df = pd.concat([tracker_df, pd.DataFrame(new_rows)], ignore_index=True)

        # Make sure any newly-introduced document columns exist (e.g. if a
        # dealer's entity type implies a doc set not yet in the CSV).
        for _, r in dealers_df.iterrows():
            for doc in all_required_docs(r.get("entity_type", "")):
                col = doc_column(doc)
                if col not in tracker_df.columns:
                    tracker_df[col] = False
        tracker_df = tracker_df.fillna({c: False for c in tracker_df.columns if c.startswith("doc__")})

    save_tracker_state(tracker_df)
    return tracker_df


def compute_completion(row, entity_type):
    docs = all_required_docs(entity_type)
    cols = [doc_column(d) for d in docs]
    done = sum(bool(row.get(c, False)) for c in cols)
    total = len(cols)
    missing = [d for d, c in zip(docs, cols) if not bool(row.get(c, False))]
    return done, total, missing


# ----------------------------------------------------------------------------
# APP
# ----------------------------------------------------------------------------

def main():
    st.title("📋 EFERT → UBL Dealer Credit Finance (DCF) Pipeline Tracker")
    st.caption(
        "Track every EFERT-recommended dealer through UBL's credit approval "
        "process — stage, owners, and document completeness vs. the Program "
        "Policy Manual checklist."
    )

    with st.sidebar:
        st.header("Data Source")
        dealer_upload = st.file_uploader("Recommended Dealers (.xlsx)", type=["xlsx"])
        poc_upload = st.file_uploader("UBL/EFERT POC List (.xlsx, optional)", type=["xlsx"])
        st.caption(
            f"Falls back to `{DEALER_FILE_DEFAULT}` / `{POC_FILE_DEFAULT}` "
            "in the app folder if nothing is uploaded."
        )

    dealer_bytes = load_workbook_bytes(dealer_upload, DEALER_FILE_DEFAULT)
    if dealer_bytes is None:
        st.warning("Upload the Recommended Dealers Excel file to get started.")
        st.stop()

    dealers_df = load_dealers(dealer_bytes)

    poc_bytes = load_workbook_bytes(poc_upload, POC_FILE_DEFAULT)
    poc_df = load_poc_directory(poc_bytes) if poc_bytes is not None else pd.DataFrame()

    tracker_df = sync_tracker_with_dealers(dealers_df)

    # Merge dealer info + tracker state
    merged = dealers_df.merge(tracker_df, on="dealer_key", how="left")

    # Suggest a UBL contact (Cluster Head / RH) from the POC directory by
    # Zone + Region, only used as a placeholder when nothing has been
    # assigned yet.
    if not poc_df.empty:
        poc_lookup = (
            poc_df.dropna(subset=["region"])
            .groupby(["zone", "region"])["ubl_ch_rh"]
            .apply(lambda s: next((x for x in s if pd.notna(x)), ""))
            .to_dict()
        )
        merged["ubl_suggested"] = merged.apply(
            lambda r: poc_lookup.get((r["zone"], r["region"]), ""), axis=1
        )
    else:
        merged["ubl_suggested"] = ""

    # Completion metrics
    completion_data = merged.apply(
        lambda r: compute_completion(r, r.get("entity_type", "")), axis=1
    )
    merged["docs_done"] = completion_data.apply(lambda t: t[0])
    merged["docs_total"] = completion_data.apply(lambda t: t[1])
    merged["docs_missing_list"] = completion_data.apply(lambda t: t[2])
    merged["completion_pct"] = (merged["docs_done"] / merged["docs_total"] * 100).round(0)

    # ------------------------------------------------------------------
    # SIDEBAR FILTERS
    # ------------------------------------------------------------------
    with st.sidebar:
        st.header("Filters")
        zones = ["All"] + sorted(merged["zone"].dropna().unique().tolist())
        f_zone = st.selectbox("Zone", zones)
        stages_filter = ["All"] + STAGES
        f_stage = st.selectbox("Stage", stages_filter)
        only_missing = st.checkbox("Only show dealers with missing documents")
        search = st.text_input("Search dealer / POC")

    view = merged.copy()
    if f_zone != "All":
        view = view[view["zone"] == f_zone]
    if f_stage != "All":
        view = view[view["stage"] == f_stage]
    if only_missing:
        view = view[view["docs_done"] < view["docs_total"]]
    if search:
        s = search.lower()
        view = view[
            view["dealer_name"].str.lower().str.contains(s)
            | view["efert_poc"].str.lower().str.contains(s)
            | view["ubl_assigned_to"].astype(str).str.lower().str.contains(s)
        ]

    # ------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Dealers in Pipeline", len(view))
    c2.metric("Fully Documented", int((view["docs_done"] == view["docs_total"]).sum()))
    c3.metric("Avg. Doc Completion", f"{view['completion_pct'].mean():.0f}%" if len(view) else "0%")
    disbursed = (view["stage"] == "7. Disbursed / Active").sum()
    c4.metric("Disbursed / Active", int(disbursed))
    on_hold = (view["stage"] == "8. On Hold / Declined").sum()
    c5.metric("On Hold / Declined", int(on_hold))

    st.markdown("#### Pipeline by Stage")
    stage_counts = view["stage"].value_counts().reindex(STAGES, fill_value=0)
    st.bar_chart(stage_counts)

    st.divider()

    # ------------------------------------------------------------------
    # MAIN EDITABLE TABLE
    # ------------------------------------------------------------------
    st.markdown("#### Pipeline Overview (editable)")
    st.caption("Edit Stage / UBL Owner / Notes directly in the table, then click **Save Changes**.")

    edit_cols = [
        "dealer_name", "zone", "region", "efert_poc", "ubl_assigned_to",
        "stage", "docs_done", "docs_total", "completion_pct", "notes", "dealer_key",
    ]
    table_view = view[edit_cols].rename(columns={
        "dealer_name": "Dealer",
        "zone": "Zone",
        "region": "Region",
        "efert_poc": "EFERT POC",
        "ubl_assigned_to": "UBL Owner",
        "stage": "Stage",
        "docs_done": "Docs Done",
        "docs_total": "Docs Required",
        "completion_pct": "Completion %",
        "notes": "Notes",
    })

    edited = st.data_editor(
        table_view,
        column_config={
            "Stage": st.column_config.SelectboxColumn(options=STAGES, required=True),
            "Docs Done": st.column_config.NumberColumn(disabled=True),
            "Docs Required": st.column_config.NumberColumn(disabled=True),
            "Completion %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%d%%"),
            "dealer_key": None,  # hide join key
        },
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="main_editor",
    )

    if st.button("💾 Save Changes", type="primary"):
        full_tracker = tracker_df.set_index("dealer_key")
        for _, row in edited.iterrows():
            key = row["dealer_key"]
            full_tracker.loc[key, "stage"] = row["Stage"]
            full_tracker.loc[key, "ubl_assigned_to"] = row["UBL Owner"]
            full_tracker.loc[key, "notes"] = row["Notes"]
            full_tracker.loc[key, "last_updated"] = dt.date.today().isoformat()
        full_tracker = full_tracker.reset_index()
        save_tracker_state(full_tracker)
        st.success("Saved. Refresh filters or reload the page to see updates everywhere.")
        st.rerun()

    st.divider()

    # ------------------------------------------------------------------
    # DEALER DETAIL — document checklist + eligibility check
    # ------------------------------------------------------------------
    st.markdown("#### Dealer Detail — Document Checklist & Eligibility")

    if view.empty:
        st.info("No dealers match the current filters.")
        return

    selected_dealer = st.selectbox("Select a dealer", view["dealer_name"].tolist())
    d_row = view[view["dealer_name"] == selected_dealer].iloc[0]
    key = d_row["dealer_key"]

    colA, colB, colC = st.columns(3)
    colA.markdown(f"**Zone / Region:** {d_row['zone']} / {d_row['region']}")
    colA.markdown(f"**EFERT POC (RSM):** {d_row['efert_poc'] or '—'}")
    colB.markdown(f"**Category:** {d_row['category']}  ·  **Rec. Limit:** PKR {d_row['recommended_limit_mn']:.0f} Mn")
    colB.markdown(f"**Entity Type:** {d_row['entity_type']}")
    colC.markdown(f"**Years with EFERT:** {d_row['years_with_engro']}")
    colC.markdown(f"**2025 Sales:** PKR {d_row['sales_2025']:,.0f}")

    ubl_owner = st.text_input(
        "UBL Owner (RM / BDO assigned to this dealer)",
        value=d_row.get("ubl_assigned_to", "") or d_row.get("ubl_suggested", ""),
        key=f"owner_{key}",
    )

    st.markdown("##### Eligibility Criteria Check (per PPM Section 5)")
    elig_checks = {
        "≥ 3 years EFERT dealership": d_row["years_with_engro"] >= 3,
        "No overdue flag with EFERT": str(d_row["overdue_flag"]).strip().upper() in ("NIL", "NONE", ""),
        "Annual sales ≥ PKR 100 Mn": d_row["sales_2025"] >= 100_000_000,
    }
    ec1, ec2, ec3 = st.columns(3)
    for (label, passed), col in zip(elig_checks.items(), [ec1, ec2, ec3]):
        col.metric(label, "✅ Pass" if passed else "⚠️ Review")

    st.markdown("##### Document Checklist (per PPM Section 6)")
    docs = all_required_docs(d_row["entity_type"])
    sections = {
        "Documents from Customer/Dealer (6.1)": DEALER_DOCS,
        "Documents for Credit Approval (6.2)": CREDIT_APPROVAL_DOCS,
        "Constitution Documents (6.3)": get_constitution_docs(d_row["entity_type"]),
    }

    updated_values = {}
    for section_name, doc_list in sections.items():
        st.markdown(f"**{section_name}**")
        cols = st.columns(2)
        for i, doc in enumerate(doc_list):
            col = doc_column(doc)
            current = bool(d_row.get(col, False))
            new_val = cols[i % 2].checkbox(doc, value=current, key=f"{key}_{col}")
            updated_values[col] = new_val

    notes_val = st.text_area("Notes", value=d_row.get("notes", "") or "", key=f"notes_{key}")

    if st.button("💾 Save Dealer Detail", key=f"save_{key}"):
        full_tracker = tracker_df.set_index("dealer_key")
        full_tracker.loc[key, "ubl_assigned_to"] = ubl_owner
        full_tracker.loc[key, "notes"] = notes_val
        full_tracker.loc[key, "last_updated"] = dt.date.today().isoformat()
        for col, val in updated_values.items():
            full_tracker.loc[key, col] = val
        full_tracker = full_tracker.reset_index()
        save_tracker_state(full_tracker)
        st.success(f"Saved checklist for {selected_dealer}.")
        st.rerun()

    done, total, missing = compute_completion(
        {**d_row.to_dict(), **updated_values}, d_row["entity_type"]
    )
    if missing:
        st.warning(f"**Missing documents ({len(missing)}/{total}):** " + "; ".join(missing))
    else:
        st.success("All required documents are complete for this dealer. ✅")

    st.divider()

    # ------------------------------------------------------------------
    # EXPORT
    # ------------------------------------------------------------------
    st.markdown("#### Export")
    export_df = merged.drop(columns=["docs_missing_list"])
    csv_bytes = export_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download full pipeline (CSV)",
        data=csv_bytes,
        file_name=f"dcf_pipeline_export_{dt.date.today().isoformat()}.csv",
        mime="text/csv",
    )


if __name__ == "__main__":
    main()
