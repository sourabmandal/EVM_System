import streamlit as st
import pandas as pd
from evm_logic import phase1_randomization, phase2_randomization, to_excel_bytes

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="EVM Randomization System",
    page_icon="🗳️",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main > div { padding-top: 1.5rem; }
    h1 { font-size: 1.6rem !important; font-weight: 600 !important; }
    h3 { font-size: 1rem !important; font-weight: 600 !important;
         color: #1F4E79 !important; margin-top: 1.2rem !important; }
    .stTabs [data-baseweb="tab"] {
        font-size: 0.95rem;
        font-weight: 500;
        padding: 0.6rem 1.4rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F4E79 !important;
        color: white !important;
        border-radius: 6px 6px 0 0;
    }
    .stDownloadButton > button {
        background-color: #1F4E79 !important;
        color: white !important;
        border: none !important;
        width: 100%;
        padding: 0.6rem;
        font-weight: 500;
    }
    .stDownloadButton > button:hover { background-color: #163d61 !important; }
    .stButton > button {
        width: 100%;
        background-color: #1F4E79;
        color: white;
        border: none;
        padding: 0.6rem;
        font-weight: 500;
        border-radius: 6px;
    }
    .stButton > button:hover { background-color: #163d61; }
    div[data-testid="stFileUploader"] label { font-weight: 500; }
    .block-container { max-width: 780px !important; }
    .step-box {
        background: #f0f6ff;
        border-left: 4px solid #1F4E79;
        border-radius: 0 8px 8px 0;
        padding: 0.75rem 1rem;
        margin-bottom: 1rem;
        font-size: 0.92rem;
        color: #1a1a2e;
    }
    .footer {
        margin-top: 3rem;
        text-align: center;
        font-size: 0.78rem;
        color: #888;
        border-top: 1px solid #eee;
        padding-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## 🗳️ EVM Randomization System")
st.markdown("Secure, Transparent & Verifiable")
st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
#tab1, tab2 = st.tabs(["  Phase 1 - Constituency Assignment  ", "  Phase 2 - Polling Station Assignment  "])

tab1 = st.tabs(["  Phase 1 - Constituency Assignment  "])
# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 1
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("""
    <div class="step-box">
    In Phase 1, every EVM is randomly assigned to one of the available constituencies.
    Each run produces a fresh, independent random assignment.
    </div>
    """, unsafe_allow_html=True)

    # Step 1 - Upload
    st.markdown("### Step 1 - Upload Input Files")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Polling Station List**")
        st.caption("Columns: Slno, RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name")
        ps_file = st.file_uploader(
            "Upload Polling Station List",
            type=["xlsx"],
            key="p1_ps",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown("**EVM Details**")
        st.caption("Columns: EVMSlNo, EVMCode")
        evm_file = st.file_uploader(
            "Upload EVM Details",
            type=["xlsx"],
            key="p1_evm",
            label_visibility="collapsed",
        )

    # File validation feedback
    if ps_file:
        st.success(f"✅ Polling station file uploaded: **{ps_file.name}**")
    if evm_file:
        st.success(f"✅ EVM details file uploaded: **{evm_file.name}**")

    # Step 2 — Randomize
    st.markdown("### Step 2 - Run Randomization")

    if st.button("⚙️  Run Phase 1 Randomization", key="run_p1",
                 disabled=not (ps_file and evm_file)):
        try:
            with st.spinner("Randomizing EVMs across constituencies..."):
                polling_df = pd.read_excel(ps_file,  dtype=str).fillna("")
                evm_df     = pd.read_excel(evm_file, dtype=str).fillna("")
                result_df, summary, warnings = phase1_randomization(polling_df, evm_df)

            st.session_state["p1_result"]  = result_df
            st.session_state["p1_summary"] = summary
            st.session_state["p1_done"]    = True

            for w in warnings:
                st.warning(w)

            st.success(
                f"✅ Randomization complete - "
                f"**{summary['total_evms']} EVMs** assigned across "
                f"**{summary['total_const']} constituencies**"
            )

            # Distribution table
            st.markdown("**EVM Distribution:**")
            dist_df = pd.DataFrame(
                list(summary["distribution"].items()),
                columns=["Constituency Name", "EVMs Assigned"]
            )
            st.dataframe(dist_df, use_container_width=True, hide_index=True)
            
            # Preview table
            st.markdown("**Preview (first 10 rows):**")
            st.dataframe(result_df.head(10), use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"❌ Error: {e}")

    elif not (ps_file and evm_file):
        st.info("Please upload both files above to enable randomization.")

    # Restore summary if already run
    elif st.session_state.get("p1_done"):
        summary   = st.session_state["p1_summary"]
        result_df = st.session_state["p1_result"]
        st.success(
            f"✅ Randomization complete — "
            f"**{summary['total_evms']} EVMs** assigned across "
            f"**{summary['total_const']} constituencies**"
        )
        dist_df = pd.DataFrame(
            list(summary["distribution"].items()),
            columns=["Constituency Name", "EVMs Assigned"]
        )
        st.dataframe(dist_df, use_container_width=True, hide_index=True)
        # Preview table
        st.markdown("**Preview (first 10 rows):**")
        st.dataframe(result_df.head(10), use_container_width=True, hide_index=True)

    # Step 3 — Download
    st.markdown("### Step 3 - Download Result")

    if st.session_state.get("p1_done"):
        result_df = st.session_state["p1_result"]
        excel_bytes = to_excel_bytes(
            result_df,
            sheet_name="Phase1_Randomization",
            col_widths=[12, 18, 25, 28],
        )
        st.download_button(
            label="⬇️  Download Phase 1 Output (Excel)",
            data=excel_bytes,
            file_name="Phase1_EVM_Randomization_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("This file will be used as input for Phase 2.")
    else:
        st.button("⬇️  Download Phase 1 Output (Excel)", disabled=True, key="dl_p1_disabled")
        st.caption("Run randomization first to enable download.")


# ═══════════════════════════════════════════════════════════════════════════════
#  PHASE 2
# ═══════════════════════════════════════════════════════════════════════════════
# with tab2:
#     st.markdown("""
#     <div class="step-box">
#     In Phase 2, polling stations are randomly assigned to EVMs <strong>within each constituency</strong>
#     based on the Phase 1 output. Phase 1 must be completed before running Phase 2.
#     </div>
#     """, unsafe_allow_html=True)

#     # Step 1 — Upload
#     st.markdown("### Step 1 - Upload Input Files")
#     col3, col4 = st.columns(2)

#     with col3:
#         st.markdown("**Phase 1 Output File**")
#         st.caption("Columns: EVMSlNo, EVMCode, RO/ARO's Office, Constituency Name")
#         p1_out_file = st.file_uploader(
#             "Upload Phase 1 Output",
#             type=["xlsx"],
#             key="p2_p1out",
#             label_visibility="collapsed",
#         )

#     with col4:
#         st.markdown("**Polling Station List**")
#         st.caption("Columns: Slno, RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name")
#         ps2_file = st.file_uploader(
#             "Upload Polling Station List",
#             type=["xlsx"],
#             key="p2_ps",
#             label_visibility="collapsed",
#         )

#     if p1_out_file:
#         st.success(f"✅ Phase 1 output uploaded: **{p1_out_file.name}**")
#     if ps2_file:
#         st.success(f"✅ Polling station file uploaded: **{ps2_file.name}**")

#     # Step 2 — Randomize
#     st.markdown("### Step 2 - Run Randomization")

#     if st.button("⚙️  Run Phase 2 Randomization", key="run_p2",
#                  disabled=not (p1_out_file and ps2_file)):
#         try:
#             with st.spinner("Assigning polling stations to EVMs..."):
#                 phase1_df  = pd.read_excel(p1_out_file, dtype=str).fillna("")
#                 polling_df = pd.read_excel(ps2_file,    dtype=str).fillna("")
#                 result2_df, summary2, warnings2 = phase2_randomization(phase1_df, polling_df)

#             st.session_state["p2_result"]  = result2_df
#             st.session_state["p2_summary"] = summary2
#             st.session_state["p2_done"]    = True

#             for w in warnings2:
#                 st.warning(f"⚠️ {w}")

#             st.success(
#                 f"✅ Phase 2 complete - "
#                 f"**{summary2['total_evms']} EVMs** assigned to polling stations"
#             )

#             # Preview table
#             st.markdown("**Preview (first 10 rows):**")
#             st.dataframe(result2_df.head(10), use_container_width=True, hide_index=True)

#         except Exception as e:
#             st.error(f"❌ Error: {e}")

#     elif not (p1_out_file and ps2_file):
#         st.info("Please upload both files above to enable randomization.")

#     elif st.session_state.get("p2_done"):
#         result2_df = st.session_state["p2_result"]
#         summary2   = st.session_state["p2_summary"]
#         st.success(
#             f"✅ Phase 2 complete - "
#             f"**{summary2['total_evms']} EVMs** assigned to polling stations"
#         )
#         st.dataframe(result2_df.head(10), use_container_width=True, hide_index=True)

#     # Step 3 — Download
#     st.markdown("### Step 3 - Download Result")

#     if st.session_state.get("p2_done"):
#         result2_df  = st.session_state["p2_result"]
#         excel_bytes2 = to_excel_bytes(
#             result2_df,
#             sheet_name="Phase2_Randomization",
#             col_widths=[12, 18, 25, 28, 20, 30],
#         )
#         st.download_button(
#             label="⬇️  Download Phase 2 Output (Excel)",
#             data=excel_bytes2,
#             file_name="Phase2_EVM_Randomization_Output.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         )
#     else:
#         st.button("⬇️  Download Phase 2 Output (Excel)", disabled=True, key="dl_p2_disabled")
#         st.caption("Run Phase 2 randomization first to enable download.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    EVM Randomization System &nbsp;|&nbsp; All processing is done locally on this machine.
    No data is transmitted externally.
</div>
""", unsafe_allow_html=True)

# ── Session state init ────────────────────────────────────────────────────────
for key in ["p1_done", "p1_result", "p1_summary", "p2_done", "p2_result", "p2_summary"]:
    if key not in st.session_state:
        st.session_state[key] = None
