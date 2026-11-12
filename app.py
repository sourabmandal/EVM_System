import time
import random
import streamlit as st
import pandas as pd
import sys
import os
import hashlib
from datetime import datetime
from zoneinfo import ZoneInfo


sys.path.append(os.path.dirname(__file__))

from evm_logic import (
    run_phase1,
    run_phase2,
    get_phase1_excel,
    get_phase2_excel,
    get_phase1_pdf,
    get_phase2_pdf,
    get_phase1_capacity,
    get_phase1_current_assignments,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="EVM Randomization System",
    page_icon="assets/favicon.png" if False else "🗳️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------

st.markdown("""
<style>

/* Try to hide GitHub & Fork links */
a[href*="github"], a[href*="fork"] {
    display: none !important;
}

/* Keep header and menu working */
header {
    visibility: visible !important;
}

button[kind="header"] {
    display: block !important;
    visibility: visible !important;
}

/* Hide footer */
footer {
    visibility: hidden !important;
}

/* Hide bottom decoration */
div[data-testid="stDecoration"] {
    display: none !important;
}

/* Reduce spacing */
.block-container {
    padding-top: 1rem !important;
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

h1, h2, h3 {
    font-family: 'Rajdhani', sans-serif !important;
    letter-spacing: 0.04em;
}

.main-header {
    background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
    padding: 2rem 2.5rem;
    border-radius: 12px;
    margin-bottom: 2rem;
    color: white;
    border-left: 5px solid #00c6ff;
}

.main-header h1 {
    font-size: 2.4rem;
    margin: 0;
    color: #fff;
    font-weight: 700;
}

.main-header p {
    color: #a8c8d8;
    margin: 0.4rem 0 0;
    font-size: 0.95rem;
    letter-spacing: 0.05em;
}

.phase-card {
    background: #f7f9fc;
    border: 1px solid #dce3ee;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}

.stat-box {
    background: linear-gradient(135deg, #e8f4fd, #d0eafc);
    border-left: 4px solid #0077b6;
    border-radius: 8px;
    padding: 1rem 1.4rem;
    margin: 0.5rem 0;
}

.stat-box .label {
    font-size: 0.78rem;
    color: #555;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}

.stat-box .value {
    font-size: 2rem;
    font-weight: 700;
    color: #0077b6;
    font-family: 'Rajdhani', sans-serif;
}

.anim-frame {
    background: #0f2027;
    color: #00c6ff;
    border-radius: 8px;
    padding: 1.2rem 1.6rem;
    font-family: 'Courier New', monospace;
    font-size: 0.82rem;
    margin: 0.5rem 0;
    min-height: 80px;
    border: 1px solid #1e4d6b;
}

.warning-box {
    background: #fff8e6;
    border-left: 4px solid #f0a500;
    border-radius: 6px;
    padding: 0.8rem 1.2rem;
    margin: 0.5rem 0;
    color: #7a5200;
    font-size: 0.9rem;
}

.success-box {
    background: #edfaf1;
    border-left: 4px solid #27ae60;
    border-radius: 6px;
    padding: 0.8rem 1.2rem;
    margin: 0.5rem 0;
    color: #1a5c34;
    font-size: 0.9rem;
}

.section-title {
    font-family: 'Rajdhani', sans-serif;
    font-size: 1.3rem;
    font-weight: 700;
    color: #203a43;
    border-bottom: 2px solid #00c6ff;
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}

div[data-testid="stDataFrame"] {
    border-radius: 8px;
    overflow: hidden;
}

.stButton > button {
    background: linear-gradient(135deg, #0077b6, #00b4d8) !important;
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.55rem 2rem;
    font-family: 'Rajdhani', sans-serif;
    font-size: 1rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    transition: all 0.2s ease;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #005f8e, #0096b8) !important;
    box-shadow: 0 4px 15px rgba(0, 119, 182, 0.4);
    transform: translateY(-1px);
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #27ae60, #2ecc71);
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.55rem 2rem;
    font-family: 'Rajdhani', sans-serif;
    font-size: 1rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    transition: all 0.2s ease;
}

.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #2ecc71, #58d68d);
    box-shadow: 0 4px 15px rgba(46, 204, 113, 0.4);
    transform: translateY(-1px);
}

.sidebar-info {
    background: #e8f4fd;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    font-size: 0.83rem;
    color: #1a3a4a;
    line-height: 1.6;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

for key in [
    "phase1_result", "phase1_current_result", "phase1_summary", "phase1_frames", "phase1_warnings",
    "phase1_remaining_bu", "phase1_remaining_cu", "phase1_finalized", "phase1_randomization_date",
    "phase2_assigned", "phase2_reserved", "phase2_summary", "phase2_frames", "phase2_warnings",
    "phase2_finalized", "phase2_randomization_date", "phase2_selected_ro", "phase2_selected_constituencies", "phase2_ro_label", "phase2_constituency_label",
]:
    if key not in st.session_state:
        if key in ["phase2_selected_ro", "phase2_selected_constituencies"]:
            st.session_state[key] = []
        else:
            st.session_state[key] = None

if "login_error" not in st.session_state:
    st.session_state["login_error"] = ""

# ---------------------------------------------------------------------------
# Helper: animation display
# ---------------------------------------------------------------------------

def display_phase1_animation(frames):
    st.markdown('<div class="section-title">Randomization Preview</div>', unsafe_allow_html=True)
    placeholder = st.empty()
    for i, frame in enumerate(frames):
        progress = int((i + 1) / len(frames) * 100)
        lines = [f"[Step {i + 1}/{len(frames)}]  Shuffling EVMs... {progress}% complete"]
        for _, row in frame.head(6).iterrows():
            lines.append(
                f"  BU: {row['BU Unit ID']}  -->  CU: {row['CU Unit ID']}  -->  {row['Constituency']}"
            )
        if len(frame) > 6:
            lines.append(f"  ... and {len(frame) - 6} more assignments")
        with placeholder.container():
            st.markdown(
                f'<div class="anim-frame">' + "<br>".join(lines) + "</div>",
                unsafe_allow_html=True,
            )
        time.sleep(0.18)
    placeholder.markdown(
        '<div class="success-box">Randomization complete! All EVMs assigned to constituencies.</div>',
        unsafe_allow_html=True,
    )


def display_phase2_animation(frames):
    st.markdown('<div class="section-title">Randomization Preview</div>', unsafe_allow_html=True)
    placeholder = st.empty()
    for i, frame in enumerate(frames):
        const = frame["constituency"]
        evms = frame["evms"]
        ps_list = frame["polling_stations"]
        lines = [f"[Constituency: {const}]  Assigning polling stations..."]
        for j in range(min(5, len(evms), len(ps_list))):
            ev = evms[j]
            ps = ps_list[j]
            lines.append(
                f"  BU:{ev['BU Unit ID']} + CU:{ev['CU Unit ID']} --> PS {ps['Polling Station No.']}: {ps['Polling Station Name']}"
            )
        if len(evms) > 5:
            lines.append(f"  ... and {len(evms) - 5} more")
        with placeholder.container():
            st.markdown(
                f'<div class="anim-frame">' + "<br>".join(lines) + "</div>",
                unsafe_allow_html=True,
            )
        time.sleep(0.28)
    placeholder.markdown(
        '<div class="success-box">Phase 2 randomization complete!</div>',
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# CAPTCHA page
# ---------------------------------------------------------------------------

def captcha_page():
    st.markdown("""
    <div class="main-header">
        <h1>EVM Randomization System</h1>
        <p>Secure Randomization Portal</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Verify Access")
        st.markdown("Please enter User ID, Password, and complete the CAPTCHA to proceed.")

        if "captcha_a" not in st.session_state or st.session_state.get("refresh_captcha"):
            a = random.randint(10, 99)
            b = random.randint(10, 99)
            st.session_state["captcha_a"] = a
            st.session_state["captcha_b"] = b
            st.session_state["refresh_captcha"] = False

        a = st.session_state["captcha_a"]
        b = st.session_state["captcha_b"]

        if st.session_state["login_error"]:
            st.error(st.session_state["login_error"])
            st.session_state["login_error"] = ""
        
        user_id = st.text_input("User ID", key="login_user_id")
        password = st.text_input("Password", type="password", key="login_password")

        st.markdown(
            f"""
            <div style="background:#0f2027;color:#00c6ff;font-family:'Courier New',monospace;
            font-size:2rem;letter-spacing:0.3em;padding:1.2rem;border-radius:8px;
            text-align:center;margin:1rem 0;border:1px solid #1e4d6b;">
            {a} + {b} = ?
            </div>
            """,
            unsafe_allow_html=True,
        )

        user_ans = st.text_input("Enter CAPTCHA answer", key="captcha_input", placeholder="Type your answer here")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Login", use_container_width=True):
                try:
                    valid_user = user_id == st.secrets["APP_USER_ID"]
                    valid_pass = password == st.secrets["APP_PASSWORD"]
                    valid_captcha = int(user_ans) == a + b

                    errors = []

                    if not valid_user or not valid_pass:
                        errors.append("User ID and/or password")
                    if not valid_captcha:
                        errors.append("CAPTCHA")
                    
                    if not errors:
                        st.session_state["authenticated"] = True
                        st.rerun()
                    else:
                        st.session_state["login_error"] = "Incorrect: " + ", ".join(errors)
                        st.session_state["refresh_captcha"] = True
                        st.rerun()
                except ValueError:
                    st.session_state["login_error"] = "Please enter a valid CAPTCHA answer."
                    st.session_state["refresh_captcha"] = True
                    st.rerun()
        with c2:
            if st.button("Refresh CAPTCHA", use_container_width=True):
                st.session_state["refresh_captcha"] = True
                st.rerun()


def get_blank_ps_template():
    return pd.DataFrame(columns=[
        "Slno",
        "RO/ARO's Office",
        "Constituency Name",
        "Polling Station No.",
        "Polling Station Name",
    ])


def get_blank_bu_template():
    df = pd.DataFrame({
        "S. No.": [1, 2, 3],
        "Unit ID": ["", "", ""],
        "Status": ["Available for use", "Rejected", "Training"],
        "Box No.": ["", "", ""],
    })
    return df


def get_blank_cu_template():
    df = pd.DataFrame({
        "S. No.": [1, 2, 3],
        "Unit ID": ["", "", ""],
        "Status": ["Available for use", "Rejected", "Training"],
        "Box No.": ["", "", ""],
    })
    return df


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

def get_suggested_constituency_allocations(ps_df, available_bu, available_cu):
    """Generate editable, polling-station-proportional defaults for exact allocation."""
    rows = (ps_df.groupby(["RO/ARO's Office", "Constituency Name"], sort=False)
            .size().reset_index(name="Polling Stations"))
    total_ps = int(rows["Polling Stations"].sum())

    def distribute(units):
        units = max(0, int(units))
        exact = [units * int(ps) / total_ps for ps in rows["Polling Stations"]]
        base = [int(value) for value in exact]
        order = sorted(range(len(base)), key=lambda i: exact[i] - base[i], reverse=True)
        for i in order[:units - sum(base)]:
            base[i] += 1
        return base

    rows["Base BUs"] = distribute(available_bu)
    rows["Base CUs"] = distribute(available_cu)
    return rows


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

def main_app():
    st.markdown("""
    <div class="main-header">
        <h1>EVM Randomization System</h1>
        <p>Secure Randomization Portal</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### Navigation")
        page = st.radio(
            "Select Phase",
            ["Phase 1: Constituency Assignment", "Phase 2: Polling Station Assignment"],
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.markdown(
            """
            <div class="sidebar-info">
            <b>Phase 1:</b> Randomly assigns all available BUs and CUs to constituencies.<br><br>
            <b>Phase 2:</b> Randomly assigns polling stations to EVMs within each constituency. Surplus EVMs are marked as Reserved.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("---")
        if st.button("Logout"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state["authenticated"] = False
            st.rerun()

    # ------------------------------------------------------------------
    # Phase 1
    # ------------------------------------------------------------------
    if page == "Phase 1: Constituency Assignment":
        st.markdown("## Phase 1: EVM Constituency Randomization")
        st.markdown(
            "Upload the three input files below. The system will randomly assign "
            "available BUs and CUs to constituencies. You can specify exact, "
            "separate quantities for each constituency."
        )

        st.markdown("### Download Blank Templates")
        t1, t2, t3 = st.columns(3)

        with t1:
            st.download_button(
                label="Download PS Template",
                data=dataframe_to_excel_bytes(get_blank_ps_template()),
                file_name="polling_station_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with t2:
            st.download_button(
                label="Download BU Template",
                data=dataframe_to_excel_bytes(get_blank_bu_template()),
                file_name="bu_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with t3:
            st.download_button(
                label="Download CU Template",
                data=dataframe_to_excel_bytes(get_blank_cu_template()),
                file_name="cu_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.expander("Upload Input Files", expanded=True):
            col1, col2, col3 = st.columns(3)

            with col1:
                st.markdown("**Polling Station List**")
                ps_file = st.file_uploader(
                    "Columns: Slno, RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name",
                    type=["xlsx", "xls"],
                    key="p1_ps",
                )

            with col2:
                st.markdown("**Ballot Unit (BU) Details**")
                bu_file = st.file_uploader(
                    "Columns: S. No. (Optional), Unit ID, Status, Box No. (Optional)",
                    type=["xlsx", "xls"],
                    key="p1_bu",
                )

            with col3:
                st.markdown("**Control Unit (CU) Details**")
                cu_file = st.file_uploader(
                    "Columns: S. No. (Optional), Unit ID, Status, Box No. (Optional)",
                    type=["xlsx", "xls"],
                    key="p1_cu",
                )

        incremental_phase1 = st.checkbox(
            "These are additional units: append to an existing Phase 1 allocation",
            value=False,
            help=("Upload the earlier Phase 1 workbook and only the newly received BU/CU inventory. "
                  "Earlier assignments remain unchanged; the downloaded result includes both."),
            key="p1_incremental_mode",
        )
        existing_phase1_file = None
        if incremental_phase1:
            existing_phase1_file = st.file_uploader(
                "Existing Phase 1 Assigned EVMs workbook",
                type=["xlsx", "xls"],
                key="p1_existing_output",
            )
            st.caption(
                "The BU/CU input files above must contain the NEW units only. "
                "The updated Phase 1 Excel output will combine old and new assignments "
                "for use by Phase 2. On-screen results will display only the "
                "new assignments; the PDF will show those first and the "
                "complete list afterwards."
            )

        st.markdown("### EVM Allocation Mode")
        allocation_method = st.radio(
            "Choose how BUs and CUs will be distributed",
            ["Automatic - polling-station ratio", "Exact - constituency-wise quantities"],
            horizontal=True,
            index=1 if incremental_phase1 else 0,
            key="p1_allocation_method_additional" if incremental_phase1 else "p1_allocation_method",
        )
        use_all_available = True
        bu_percentage = None
        cu_percentage = None
        constituency_allocations = None
        capacity = None

        if ps_file and bu_file and cu_file:
            try:
                capacity = get_phase1_capacity(ps_file, bu_file, cu_file)
                total_ps = capacity["total_polling_stations"]
                st.info(
                    f"Total Polling Stations: {total_ps}  |  "
                    f"Available BUs: {capacity['available_bus']} "
                    f"(maximum {capacity['max_bu_percentage']:g}% of stations)  |  "
                    f"Available CUs: {capacity['available_cus']} "
                    f"(maximum {capacity['max_cu_percentage']:g}% of stations)"
                )
                if allocation_method == "Exact - constituency-wise quantities":
                    ps_preview = pd.read_excel(ps_file)
                    default_allocation = get_suggested_constituency_allocations(
                        ps_preview, capacity["available_bus"], capacity["available_cus"]
                    )
                    input_signature = hashlib.sha256(
                        ps_file.getvalue() + bu_file.getvalue() + cu_file.getvalue()
                    ).hexdigest()[:16]
                    st.markdown("**Enter exact, separate BU and CU quantities for each constituency.**")
                    st.caption(
                        "Do not enter BU/CU pairs as a single number. Specify both values "
                        "independently.
                    )
                    requested_rows = []
                    for i, row in default_allocation.iterrows():
                        name = str(row["Constituency Name"])
                        ro_office = str(row["RO/ARO's Office"])
                        st.markdown(
                            f"**{name}**  |  {int(row['Polling Stations'])} polling station(s)  |  "
                            f"{ro_office}"
                        )
                        col_bu, col_cu = st.columns(2)
                        with col_bu:
                            requested_bu = st.number_input(
                                f"BUs to allocate to {name}",
                                min_value=0,
                                max_value=int(capacity["available_bus"]),
                                value=int(row["Base BUs"]),
                                step=1,
                                key=f"p1_exact_bu_{input_signature}_{i}",
                            )
                        with col_cu:
                            requested_cu = st.number_input(
                                f"CUs to allocate to {name}",
                                min_value=0,
                                max_value=int(capacity["available_cus"]),
                                value=int(row["Base CUs"]),
                                step=1,
                                key=f"p1_exact_cu_{input_signature}_{i}",
                            )
                        requested_rows.append({
                            "Constituency Name": name,
                            "Base BUs": int(requested_bu),
                            "Base CUs": int(requested_cu),
                        })

                    constituency_allocations = pd.DataFrame(requested_rows)
                    planned_bu = int(constituency_allocations["Base BUs"].sum())
                    planned_cu = int(constituency_allocations["Base CUs"].sum())
                    st.markdown("**Confirm constituency-wise allocation before randomization:**")
                    preview = default_allocation[["Constituency Name", "Polling Stations"]].merge(
                        constituency_allocations, on="Constituency Name", how="left", validate="one_to_one"
                    )
                    preview["Assigned BU vs PS (%)"] = (
                        100 * preview["Base BUs"] / preview["Polling Stations"]
                    ).round(2)
                    preview["Assigned CU vs PS (%)"] = (
                        100 * preview["Base CUs"] / preview["Polling Stations"]
                    ).round(2)
                    st.dataframe(preview.rename(columns={
                        "Base BUs": "BUs to Allocate", "Base CUs": "CUs to Allocate",
                    }), use_container_width=True, hide_index=True)
                    if planned_bu > capacity["available_bus"] or planned_cu > capacity["available_cus"]:
                        st.error(
                            f"Insufficient stock: the plan needs {planned_bu} BUs and {planned_cu} CUs, "
                            f"but only {capacity['available_bus']} BUs and {capacity['available_cus']} CUs "
                            "are available. Reduce the input counts or provide more units."
                        )
                    else:
                        st.success(
                            f"Confirmed totals: {planned_bu} BUs and {planned_cu} CUs. "
                            f"Unallocated stock: {capacity['available_bus'] - planned_bu} BUs "
                            f"and {capacity['available_cus'] - planned_cu} CUs."
                        )
                        if incremental_phase1 and (
                            planned_bu < capacity["available_bus"] or planned_cu < capacity["available_cus"]
                        ):
                            st.warning(
                                "Some newly uploaded units are not included in this additional allocation. "
                                "Increase the constituency quantities if you intend to assign all of them."
                            )
                else:
                    use_all_available = st.checkbox(
                        "Use all available BUs and CUs", value=True,
                        help="Assign all available units in Phase 1; optional reserves are selected in Phase 2.",
                    )
                    if not use_all_available:
                        max_bu_pct = int(capacity["available_bus"] / total_ps * 100)
                        max_cu_pct = int(capacity["available_cus"] / total_ps * 100)
                        if max_bu_pct < 1 or max_cu_pct < 1:
                            st.error("Not enough units for positive BU and CU assignment percentages.")
                        else:
                            col_bu, col_cu = st.columns(2)
                            with col_bu:
                                bu_percentage = st.number_input(
                                    "BU percentage to assign",
                                    min_value=1, max_value=max_bu_pct,
                                    value=max_bu_pct, step=1, key="p1_bu_percentage",
                                )
                            with col_cu:
                                cu_percentage = st.number_input(
                                    "CU percentage to assign",
                                    min_value=1, max_value=max_cu_pct,
                                    value=max_cu_pct, step=1, key="p1_cu_percentage",
                                )

            except Exception as e:
                st.warning(f"Could not prepare Phase 1 allocation: {e}")

        # A prior randomization must never appear to reflect newly edited quotas.
        current_inputs = hashlib.sha256()
        for uploaded in (ps_file, bu_file, cu_file, existing_phase1_file):
            current_inputs.update(uploaded.getvalue() if uploaded is not None else b"<not uploaded>")
            current_inputs.update(b"\0")
        current_inputs.update(repr((
            incremental_phase1, allocation_method, use_all_available,
            bu_percentage, cu_percentage,
            constituency_allocations.to_dict("records") if constituency_allocations is not None else None,
        )).encode("utf-8"))
        current_inputs_digest = current_inputs.hexdigest()
        if (st.session_state.get("phase1_run_input_digest") is not None
                and st.session_state["phase1_run_input_digest"] != current_inputs_digest):
            st.session_state["phase1_result"] = None
            st.session_state["phase1_current_result"] = None
            st.session_state["phase1_finalized"] = False
            st.session_state["phase1_run_input_digest"] = None

        if st.button("Run Phase 1 Randomization", use_container_width=True):
            if not ps_file or not bu_file or not cu_file:
                st.error("Please upload all three files before running.")
            elif incremental_phase1 and not existing_phase1_file:
                st.error("Upload the existing Phase 1 output to append additional units.")
            elif capacity is None:
                st.error("Please correct the uploaded Phase 1 files before randomization.")
            elif allocation_method == "Exact - constituency-wise quantities" and (
                constituency_allocations is None
                or int(constituency_allocations["Base BUs"].sum()) > capacity["available_bus"]
                or int(constituency_allocations["Base CUs"].sum()) > capacity["available_cus"]
            ):
                st.error("Enter valid exact BU/CU quantities for every constituency before running.")
            else:
                with st.spinner("Randomizing EVMs across constituencies..."):
                    try:
                        seed = random.randint(0, 999999)
                        result_df, summary, frames, warnings, remaining_bu_df, remaining_cu_df = run_phase1(
                            ps_file,
                            bu_file,
                            cu_file,
                            seed=seed,
                            use_all_available=use_all_available,
                            bu_percentage=bu_percentage,
                            cu_percentage=cu_percentage,
                            constituency_allocations=constituency_allocations,
                            existing_phase1_file=existing_phase1_file,
                        )
                        st.session_state["phase1_incremental_used"] = incremental_phase1
                        st.session_state["phase1_result"] = result_df
                        st.session_state["phase1_current_result"] = get_phase1_current_assignments(
                            result_df, summary if incremental_phase1 else None
                        )
                        st.session_state["phase1_summary"] = summary
                        st.session_state["phase1_run_input_digest"] = current_inputs_digest
                        st.session_state["phase1_frames"] = frames
                        st.session_state["phase1_warnings"] = warnings
                        st.session_state["phase1_remaining_bu"] = remaining_bu_df
                        st.session_state["phase1_remaining_cu"] = remaining_cu_df
                        st.session_state["phase1_finalized"] = False
                        st.session_state["phase1_randomization_date"] = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %H:%M:%S")
                    except Exception as e:
                        st.error(f"Error during Phase 1 randomization: {e}")

        if st.session_state["phase1_result"] is not None:
            warnings = st.session_state["phase1_warnings"] or []
            for w in warnings:
                st.markdown(f'<div class="warning-box">Warning: {w}</div>', unsafe_allow_html=True)

            frames = st.session_state["phase1_frames"]
            if frames:
                display_phase1_animation(frames)

            incremental_used = st.session_state.get("phase1_incremental_used", False)
            if incremental_used:
                st.markdown(
                    '<div class="section-title">New Units: Constituency-wise Summary</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    '<div class="section-title">Constituency-wise EVM Summary</div>',
                    unsafe_allow_html=True,
                )
            summary = st.session_state["phase1_summary"].copy()
            result_df = st.session_state["phase1_result"]
            current_result_df = st.session_state.get("phase1_current_result")
            if current_result_df is None:
                current_result_df = result_df
            if incremental_used:
                # Screen totals and percentages describe only the current receipt.
                summary["No. of BUs Assigned"] = summary["New BUs Assigned"]
                summary["No. of CUs Assigned"] = summary["New CUs Assigned"]
                summary["Assigned BU vs PS (%)"] = (
                    100 * summary["No. of BUs Assigned"] / summary["No. of Polling Stations"]
                ).round(2)
                summary["Assigned CU vs PS (%)"] = (
                    100 * summary["No. of CUs Assigned"] / summary["No. of Polling Stations"]
                ).round(2)
                summary = summary.drop(columns=["New BUs Assigned", "New CUs Assigned"])
                st.caption(
                    "Only newly allocated BUs/CUs are shown on this page. "
                    "The Excel download contains the combined allocation; "
                    "the PDF shows the new assignments first, followed by the complete list."
                )

            total_bus_assigned = summary["No. of BUs Assigned"].sum()
            total_cus_assigned = summary["No. of CUs Assigned"].sum()
            total_const = len(summary)
            total_ps_count = summary["No. of Polling Stations"].sum()
            
            remaining_bu_records = st.session_state.get("phase1_remaining_bu")
            remaining_cu_records = st.session_state.get("phase1_remaining_cu")
            remaining_bus = len(remaining_bu_records) if remaining_bu_records is not None else 0
            remaining_cus = len(remaining_cu_records) if remaining_cu_records is not None else 0
            
            st.dataframe(summary, use_container_width=True, hide_index=True)
            if total_ps_count:
                st.caption(
                    f"Overall assigned BU vs polling stations: "
                    f"{100 * total_bus_assigned / total_ps_count:.2f}%  |  "
                    f"Assigned CU vs polling stations: "
                    f"{100 * total_cus_assigned / total_ps_count:.2f}%"
                )

            total_bus = summary["No. of BUs Assigned"].sum()
            total_cus = summary["No. of CUs Assigned"].sum()
            total_const = len(summary)
            total_ps_count = summary["No. of Polling Stations"].sum()

            col1, col2, col3, col4, col5, col6 = st.columns(6)

            with col1:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Constituencies</div>'
                    f'<div class="value">{total_const}</div></div>',
                    unsafe_allow_html=True,
                )
            
            with col2:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Polling Stations</div>'
                    f'<div class="value">{total_ps_count}</div></div>',
                    unsafe_allow_html=True,
                )
            
            with col3:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total BU Assigned</div>'
                    f'<div class="value">{total_bus_assigned}</div></div>',
                    unsafe_allow_html=True,
                )
            
            with col4:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total CU Assigned</div>'
                    f'<div class="value">{total_cus_assigned}</div></div>',
                    unsafe_allow_html=True,
                )
            
            with col5:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Remaining BU</div>'
                    f'<div class="value">{remaining_bus}</div></div>',
                    unsafe_allow_html=True,
                )
            
            with col6:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Remaining CU</div>'
                    f'<div class="value">{remaining_cus}</div></div>',
                    unsafe_allow_html=True,
                )

            if incremental_used:
                st.markdown('<div class="section-title">Current Randomization - Additional Units Only</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="section-title">Full Randomization Output</div>', unsafe_allow_html=True)
            st.dataframe(current_result_df, use_container_width=True, hide_index=True)

            remaining_bu_df = st.session_state["phase1_remaining_bu"]
            remaining_cu_df = st.session_state["phase1_remaining_cu"]
            
            st.markdown('<div class="section-title">Remaining BUs and CUs</div>', unsafe_allow_html=True)
            if st.session_state.get("phase1_incremental_used"):
                st.caption("Remaining units from the newly uploaded stock only (not prior allocations).")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Remaining BUs**")
                if remaining_bu_df is not None and not remaining_bu_df.empty:
                    st.dataframe(remaining_bu_df, use_container_width=True, hide_index=True)
                else:
                    st.markdown(
                        '<div class="success-box">No remaining BUs.</div>',
                        unsafe_allow_html=True,
                    )
            
            with col2:
                st.markdown("**Remaining CUs**")
                if remaining_cu_df is not None and not remaining_cu_df.empty:
                    st.dataframe(remaining_cu_df, use_container_width=True, hide_index=True)
                else:
                    st.markdown(
                        '<div class="success-box">No remaining CUs.</div>',
                        unsafe_allow_html=True,
                    )

            if not st.session_state.get("phase1_finalized", False):
                if st.button("Finalize Randomization", key="phase1_finalize_btn", use_container_width=True):
                    st.session_state["phase1_finalized"] = True
                    st.rerun()
            else:
                excel_bytes = get_phase1_excel(
                    result_df,
                    st.session_state.get("phase1_remaining_bu"),
                    st.session_state.get("phase1_remaining_cu"),
                    summary_df=st.session_state["phase1_summary"].drop(
                        columns=["New BUs Assigned", "New CUs Assigned"], errors="ignore"
                    ),
                )
                pdf_bytes = get_phase1_pdf(
                    result_df,
                    st.session_state.get("phase1_remaining_bu"),
                    st.session_state.get("phase1_remaining_cu"),
                    randomization_date=st.session_state.get("phase1_randomization_date", ""),
                    current_result_df=current_result_df if incremental_used else None,
                )

                d1, d2 = st.columns(2)
                with d1:
                    st.download_button(
                        label="Download Phase 1 Excel",
                        data=excel_bytes,
                        file_name="phase1_randomization_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with d2:
                    st.download_button(
                        label="Download Phase 1 PDF",
                        data=pdf_bytes,
                        file_name="phase1_randomization_output.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

    # ------------------------------------------------------------------
    # Phase 2
    # ------------------------------------------------------------------
    else:
        st.markdown("## Phase 2: Polling Station Randomization")
        st.markdown(
            "Upload the Phase 1 output file and the Polling Station List. "
            "The system will randomly assign polling stations to EVMs within each constituency."
        )

        with st.expander("Upload Input Files", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**Phase 1 Output File**")
                p1_out_file = st.file_uploader(
                    "Columns: SlNo, BU Unit ID, CU Unit ID, RO/ARO's Office, Constituency Name",
                    type=["xlsx", "xls"],
                    key="p2_p1out",
                )

            with col2:
                st.markdown("**Polling Station List**")
                ps2_file = st.file_uploader(
                    "Columns: Slno, RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name",
                    type=["xlsx", "xls"],
                    key="p2_ps",
                )

        reserve_per_polling_station = st.checkbox(
            "Reserve one additional BU and CU pair per polling station (constituency-wise)",
            value=False,
            key="p2_extra_reserve",
            help=("When selected, Phase 2 randomly sets aside one complete BU/CU pair for "
                  "each polling station within its own constituency. It uses units already "
                  "assigned in Phase 1; at least two BUs and two CUs per polling station "
                  "must be available in that constituency."),
        )
        if reserve_per_polling_station:
            st.caption(
                "Each polling station gets one working BU/CU pair plus one designated "
                "reserved BU/CU pair. Any remaining surplus also stays reserved. "
                "No extra units are added to Phase 1."
            )

        selected_ro_offices = None
        selected_constituencies = None

        if p1_out_file and ps2_file:
            try:
                p1_preview = pd.read_excel(p1_out_file)
                ps_preview = pd.read_excel(ps2_file)

                p1_preview["RO/ARO's Office"] = p1_preview["RO/ARO's Office"].astype(str).str.strip()
                p1_preview["Constituency Name"] = p1_preview["Constituency Name"].astype(str).str.strip()
                
                ps_preview["RO/ARO's Office"] = ps_preview["RO/ARO's Office"].astype(str).str.strip()
                ps_preview["Constituency Name"] = ps_preview["Constituency Name"].astype(str).str.strip()
                
                common_ro = sorted(
                    set(p1_preview["RO/ARO's Office"].dropna()) &
                    set(ps_preview["RO/ARO's Office"].dropna())
                )
                
                if len(common_ro) > 1:
                    selected_ro_offices = st.multiselect(
                        "Select RO Office(s) to randomize",
                        options=common_ro,
                        default=st.session_state.get("phase2_selected_ro", []),
                        key="phase2_selected_ro",
                    )
                elif len(common_ro) == 1:
                    selected_ro_offices = common_ro
                else:
                    selected_ro_offices = []
                
                if selected_ro_offices:
                    p1_const_filtered = p1_preview[p1_preview["RO/ARO's Office"].isin(selected_ro_offices)]
                    ps_const_filtered = ps_preview[ps_preview["RO/ARO's Office"].isin(selected_ro_offices)]
                else:
                    p1_const_filtered = p1_preview
                    ps_const_filtered = ps_preview
                
                common_const = sorted(
                    set(p1_const_filtered["Constituency Name"].dropna()) &
                    set(ps_const_filtered["Constituency Name"].dropna())
                )
                
                if len(common_const) > 1:
                    selected_constituencies = st.multiselect(
                        "Select Constituency / Constituencies to randomize",
                        options=common_const,
                        default=st.session_state.get("phase2_selected_constituencies", []),
                        key="phase2_selected_constituencies",
                    )
                elif len(common_const) == 1:
                    selected_constituencies = common_const
                else:
                    selected_constituencies = []

            except Exception as e:
                st.warning(f"Could not prepare RO/Constituency selector: {e}")

        if st.button("Run Phase 2 Randomization", use_container_width=True):
            if not p1_out_file or not ps2_file:
                st.error("Please upload both files before running.")
            else:
                with st.spinner("Assigning polling stations to EVMs..."):
                    try:
                        seed = random.randint(0, 999999)
                        p1_out_file.seek(0)
                        ps2_file.seek(0)
                        assigned_df, reserved_df, reserved_summary, anim_frames, warnings = run_phase2(
                            p1_out_file,
                            ps2_file,
                            seed=seed,
                            selected_ro_offices=selected_ro_offices,
                            selected_constituencies=selected_constituencies,
                            reserve_per_polling_station=reserve_per_polling_station,
                        )
                        st.session_state["phase2_reserve_mode_used"] = reserve_per_polling_station
                        st.session_state["phase2_assigned"] = assigned_df
                        st.session_state["phase2_reserved"] = reserved_df
                        st.session_state["phase2_summary"] = reserved_summary
                        st.session_state["phase2_frames"] = anim_frames
                        st.session_state["phase2_warnings"] = warnings
                        st.session_state["phase2_finalized"] = False
                        st.session_state["phase2_randomization_date"] = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %H:%M:%S")
                        st.session_state["phase2_ro_label"] = ", ".join(selected_ro_offices) if selected_ro_offices else "All"
                        st.session_state["phase2_constituency_label"] = ", ".join(selected_constituencies) if selected_constituencies else "All"
                    except Exception as e:
                        st.error(f"Error during Phase 2 randomization: {e}")

        if st.session_state["phase2_assigned"] is not None:
            warnings = st.session_state["phase2_warnings"] or []
            for w in warnings:
                st.markdown(f'<div class="warning-box">Warning: {w}</div>', unsafe_allow_html=True)

            assigned_df = st.session_state["phase2_assigned"]
            reserved_df = st.session_state["phase2_reserved"]
            reserved_summary = st.session_state["phase2_summary"]

            if st.session_state.get("phase2_reserve_mode_used"):
                st.caption(
                    "Polling-station reserve enabled: one designated BU/CU reserve pair "
                    "per station is shown in the Reserved EVMs list, identified by its "
                    "polling station number and a RESERVED label."
                )
            st.markdown('<div class="section-title">Reserved EVMs - Constituency Wise</div>', unsafe_allow_html=True)
            if not reserved_summary.empty:
                st.dataframe(reserved_summary, use_container_width=True, hide_index=True)
            else:
                st.markdown(
                    '<div class="success-box">No reserved EVMs - all EVMs have been assigned to a polling station.</div>',
                    unsafe_allow_html=True,
                )

            total_ps = int(reserved_summary["No. of Polling Stations"].sum()) if not reserved_summary.empty else 0
            total_assigned_bu = int(assigned_df["BU Unit ID"].notna().sum()) if "BU Unit ID" in assigned_df.columns else 0
            total_assigned_cu = int(assigned_df["CU Unit ID"].notna().sum()) if "CU Unit ID" in assigned_df.columns else 0
            total_reserved_bu = int(reserved_summary["Reserved BU"].sum()) if "Reserved BU" in reserved_summary.columns else 0
            total_reserved_cu = int(reserved_summary["Reserved CU"].sum()) if "Reserved CU" in reserved_summary.columns else 0

            col0, col1, col2, col3, col4 = st.columns(5)
            with col0:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Polling Stations</div>'
                    f'<div class="value">{total_ps}</div></div>',
                    unsafe_allow_html=True,
                )
            with col1:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Assigned BU</div>'
                    f'<div class="value">{total_assigned_bu}</div></div>',
                    unsafe_allow_html=True,
                )
            with col2:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Assigned CU</div>'
                    f'<div class="value">{total_assigned_cu}</div></div>',
                    unsafe_allow_html=True,
                )
            with col3:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Reserved BU</div>'
                    f'<div class="value">{total_reserved_bu}</div></div>',
                    unsafe_allow_html=True,
                )
            with col4:
                st.markdown(
                    f'<div class="stat-box"><div class="label">Total Reserved CU</div>'
                    f'<div class="value">{total_reserved_cu}</div></div>',
                    unsafe_allow_html=True,
                )

            anim_frames = st.session_state["phase2_frames"]
            if anim_frames:
                display_phase2_animation(anim_frames)

            st.markdown('<div class="section-title">Assigned EVMs</div>', unsafe_allow_html=True)
            st.dataframe(assigned_df, use_container_width=True, hide_index=True)

            if not reserved_df.empty:
                st.markdown('<div class="section-title">Reserved EVMs</div>', unsafe_allow_html=True)
                st.dataframe(reserved_df, use_container_width=True, hide_index=True)

            if not st.session_state.get("phase2_finalized", False):
                if st.button("Finalize Randomization", key="phase2_finalize_btn", use_container_width=True):
                    st.session_state["phase2_finalized"] = True
                    st.rerun()
            else:
                excel_bytes = get_phase2_excel(assigned_df, reserved_df)
                pdf_bytes = get_phase2_pdf(
                    assigned_df,
                    reserved_df,
                    randomization_date=st.session_state.get("phase2_randomization_date", ""),
                    constituency_label=st.session_state.get("phase2_constituency_label"),
                    ro_office_label=st.session_state.get("phase2_ro_label"),
                )

                d1, d2 = st.columns(2)
                with d1:
                    st.download_button(
                        label="Download Phase 2 Excel",
                        data=excel_bytes,
                        file_name="phase2_randomization_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with d2:
                    st.download_button(
                        label="Download Phase 2 PDF",
                        data=pdf_bytes,
                        file_name="phase2_randomization_output.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if st.session_state["authenticated"] is True:
    main_app()
else:
    captcha_page()
