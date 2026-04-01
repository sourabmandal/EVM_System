import pandas as pd
import random
from io import BytesIO


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def _read_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)


def _normalise_status(df: pd.DataFrame) -> pd.DataFrame:
    """Lower-case strip the Status column so matching is robust."""
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip().str.lower()
    return df


# ---------------------------------------------------------------------------
# Phase 1
# ---------------------------------------------------------------------------

def run_phase1(ps_file, bu_file, cu_file, seed=None):
    """
    Randomly assign available BUs and CUs to constituencies.

    Input columns:
      PS file : RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name
      BU file : Unit ID, Status
      CU file : Unit ID, Status

    Output columns:
      EVMSlNo, BU Unit ID, CU Unit ID, RO/ARO's Office, Constituency Name

    Returns
    -------
    result_df            : final output DataFrame
    constituency_summary : constituency-wise BU / CU counts
    animation_frames     : list of DataFrames for animation preview
    warnings             : list of warning strings
    """
    if seed is not None:
        random.seed(seed)

    warnings = []

    # -- Read inputs ----------------------------------------------------------
    ps_df = _read_excel(ps_file)
    bu_df = _read_excel(bu_file)
    cu_df = _read_excel(cu_file)

    # Validate required columns (Box No. and serial number cols not required)
    required_ps = {"RO/ARO's Office", "Constituency Name",
                   "Polling Station No.", "Polling Station Name"}
    required_bu = {"Unit ID", "Status"}
    required_cu = {"Unit ID", "Status"}

    missing_ps = required_ps - set(ps_df.columns)
    missing_bu = required_bu - set(bu_df.columns)
    missing_cu = required_cu - set(cu_df.columns)

    if missing_ps or missing_bu or missing_cu:
        raise ValueError(
            f"Missing columns -- PS: {missing_ps}, BU: {missing_bu}, CU: {missing_cu}"
        )

    # -- Filter only 'available for use' units --------------------------------
    bu_df = _normalise_status(bu_df)
    cu_df = _normalise_status(cu_df)

    bu_available = bu_df[bu_df["Status"] == "available for use"][["Unit ID"]].copy().reset_index(drop=True)
    cu_available = cu_df[cu_df["Status"] == "available for use"][["Unit ID"]].copy().reset_index(drop=True)

    if bu_available.empty:
        raise ValueError("No BUs with status 'available for use' found.")
    if cu_available.empty:
        raise ValueError("No CUs with status 'available for use' found.")

    # -- Unique constituencies from PS list -----------------------------------
    constituency_info = (
        ps_df[["RO/ARO's Office", "Constituency Name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    constituencies = constituency_info["Constituency Name"].tolist()
    n_const = len(constituencies)

    if n_const == 0:
        raise ValueError("No constituencies found in Polling Station file.")

    # -- Shuffle BU and CU lists independently --------------------------------
    bu_list = bu_available["Unit ID"].tolist()
    cu_list = cu_available["Unit ID"].tolist()
    random.shuffle(bu_list)
    random.shuffle(cu_list)

    # -- Assign BUs to constituencies round-robin -----------------------------
    bu_assignments = []
    for idx, unit_id in enumerate(bu_list):
        const_name = constituencies[idx % n_const]
        ro_office = constituency_info.loc[
            constituency_info["Constituency Name"] == const_name,
            "RO/ARO's Office"
        ].values[0]
        bu_assignments.append({
            "Constituency Name": const_name,
            "RO/ARO's Office": ro_office,
            "BU Unit ID": unit_id,
        })

    # -- Assign CUs to constituencies round-robin -----------------------------
    cu_assignments = []
    for idx, unit_id in enumerate(cu_list):
        const_name = constituencies[idx % n_const]
        cu_assignments.append({
            "Constituency Name": const_name,
            "CU Unit ID": unit_id,
        })

    bu_assign_df = pd.DataFrame(bu_assignments)
    cu_assign_df = pd.DataFrame(cu_assignments)

    # -- Pair BUs and CUs within each constituency and build result -----------
    result_rows = []
    evm_sl = 1

    for const_name in constituencies:
        ro_office = constituency_info.loc[
            constituency_info["Constituency Name"] == const_name,
            "RO/ARO's Office"
        ].values[0]

        bu_const = bu_assign_df[bu_assign_df["Constituency Name"] == const_name].reset_index(drop=True)
        cu_const = cu_assign_df[cu_assign_df["Constituency Name"] == const_name].reset_index(drop=True)

        n_pairs = min(len(bu_const), len(cu_const))

        if len(bu_const) != len(cu_const):
            warnings.append(
                f"Constituency '{const_name}': BU count ({len(bu_const)}) != "
                f"CU count ({len(cu_const)}). Only {n_pairs} pairs created."
            )

        for i in range(n_pairs):
            result_rows.append({
                "EVMSlNo": evm_sl,
                "BU Unit ID": bu_const.at[i, "BU Unit ID"],
                "CU Unit ID": cu_const.at[i, "CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
            })
            evm_sl += 1

    result_df = pd.DataFrame(result_rows, columns=[
        "EVMSlNo", "BU Unit ID", "CU Unit ID",
        "RO/ARO's Office", "Constituency Name"
    ])

    # -- Constituency summary -------------------------------------------------
    summary_rows = []
    for const_name in constituencies:
        sub = result_df[result_df["Constituency Name"] == const_name]
        summary_rows.append({
            "Constituency Name": const_name,
            "RO/ARO's Office": sub["RO/ARO's Office"].iloc[0] if not sub.empty else "",
            "No. of BUs Assigned": len(sub),
            "No. of CUs Assigned": len(sub),
            "Total EVMs": len(sub),
        })
    constituency_summary = pd.DataFrame(summary_rows)

    # -- Animation frames (10 snapshots of progressive shuffle) ---------------
    animation_frames = _build_phase1_frames(bu_list, cu_list, constituencies, n_snaps=10)

    return result_df, constituency_summary, animation_frames, warnings


def _build_phase1_frames(bu_list, cu_list, constituencies, n_snaps=10):
    """Return list of DataFrames showing progressive assignment for animation."""
    frames = []
    total = min(len(bu_list), len(cu_list))
    step = max(1, total // n_snaps)
    n_const = len(constituencies)

    for snap in range(1, n_snaps + 1):
        count = min(snap * step, total)
        rows = []
        for i in range(count):
            rows.append({
                "BU Unit ID": bu_list[i],
                "CU Unit ID": cu_list[i],
                "Constituency": constituencies[i % n_const],
            })
        frames.append(pd.DataFrame(rows))

    return frames


def get_phase1_excel(result_df: pd.DataFrame) -> bytes:
    return _to_excel_bytes(result_df)


# ---------------------------------------------------------------------------
# Phase 2
# ---------------------------------------------------------------------------

def run_phase2(phase1_file, ps_file, seed=None):
    """
    Phase 2 randomization:
      - Within each constituency, re-shuffle BUs independently and CUs independently.
      - Also shuffle polling stations within each constituency.
      - Pair: one polling station gets one BU and one CU.
      - Surplus EVM pairs (no polling station available) are marked RESERVED.
      - BUs and CUs are NEVER moved out of their Phase 1 assigned constituency.

    Input columns (Phase 1 output):
      EVMSlNo, BU Unit ID, CU Unit ID, RO/ARO's Office, Constituency Name

    Input columns (PS file):
      RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name

    Output columns:
      EVMSlNo, BU Unit ID, CU Unit ID,
      RO/ARO's Office, Constituency Name,
      Polling Station No., Polling Station Name

    Returns
    -------
    assigned_df      : rows where a polling station is assigned
    reserved_df      : surplus BU+CU rows marked RESERVED
    reserved_summary : constituency-wise reserved count
    animation_frames : list of dicts for animation preview
    warnings         : list of warning strings
    """
    if seed is not None:
        random.seed(seed)

    warnings = []

    # -- Read inputs ----------------------------------------------------------
    phase1_df = _read_excel(phase1_file)
    ps_df = _read_excel(ps_file)

    required_p1 = {"EVMSlNo", "BU Unit ID", "CU Unit ID",
                   "RO/ARO's Office", "Constituency Name"}
    required_ps = {"RO/ARO's Office", "Constituency Name",
                   "Polling Station No.", "Polling Station Name"}

    missing_p1 = required_p1 - set(phase1_df.columns)
    missing_ps = required_ps - set(ps_df.columns)

    if missing_p1 or missing_ps:
        raise ValueError(
            f"Missing columns -- Phase1 file: {missing_p1}, PS file: {missing_ps}"
        )

    constituencies = phase1_df["Constituency Name"].unique().tolist()

    assigned_rows = []
    reserved_rows = []
    animation_frames = []
    new_evm_sl = 1

    for const_name in constituencies:
        evm_sub = phase1_df[phase1_df["Constituency Name"] == const_name].copy().reset_index(drop=True)
        ps_sub = ps_df[ps_df["Constituency Name"] == const_name].copy().reset_index(drop=True)

        ro_office = evm_sub["RO/ARO's Office"].iloc[0]

        # Re-shuffle BUs and CUs independently within this constituency
        bu_list = evm_sub["BU Unit ID"].tolist()
        cu_list = evm_sub["CU Unit ID"].tolist()
        random.shuffle(bu_list)
        random.shuffle(cu_list)

        # Shuffle polling stations within this constituency
        ps_list = ps_sub[["Polling Station No.", "Polling Station Name"]].values.tolist()
        random.shuffle(ps_list)

        n_bu = len(bu_list)
        n_cu = len(cu_list)
        n_ps = len(ps_list)

        # Total EVM pairs available (limited by whichever unit has fewer)
        n_total_evms = min(n_bu, n_cu)
        # Assigned = how many polling stations can be served
        n_assign = min(n_total_evms, n_ps)

        if n_ps > n_total_evms:
            warnings.append(
                f"Constituency '{const_name}': More polling stations ({n_ps}) "
                f"than EVM pairs ({n_total_evms}). Some polling stations will not get an EVM."
            )
        if n_total_evms > n_ps:
            warnings.append(
                f"Constituency '{const_name}': More EVM pairs ({n_total_evms}) "
                f"than polling stations ({n_ps}). {n_total_evms - n_ps} EVMs will be reserved."
            )

        # Animation snapshot for this constituency
        animation_frames.append({
            "constituency": const_name,
            "evms": [
                {"BU Unit ID": bu_list[i], "CU Unit ID": cu_list[i]}
                for i in range(min(n_assign, 5))
            ],
            "polling_stations": [
                {"Polling Station No.": ps_list[i][0], "Polling Station Name": ps_list[i][1]}
                for i in range(min(n_assign, 5))
            ],
        })

        # Assigned rows: each polling station gets one freshly shuffled BU and CU
        for i in range(n_assign):
            assigned_rows.append({
                "EVMSlNo": new_evm_sl,
                "BU Unit ID": bu_list[i],
                "CU Unit ID": cu_list[i],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": ps_list[i][0],
                "Polling Station Name": ps_list[i][1],
            })
            new_evm_sl += 1

        # Reserved rows: surplus EVM pairs with no polling station
        for i in range(n_assign, n_total_evms):
            reserved_rows.append({
                "EVMSlNo": new_evm_sl,
                "BU Unit ID": bu_list[i],
                "CU Unit ID": cu_list[i],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": "NA",
                "Polling Station Name": "RESERVED",
            })
            new_evm_sl += 1

    output_cols = [
        "EVMSlNo", "BU Unit ID", "CU Unit ID",
        "RO/ARO's Office", "Constituency Name",
        "Polling Station No.", "Polling Station Name"
    ]

    assigned_df = (
        pd.DataFrame(assigned_rows, columns=output_cols)
        if assigned_rows else pd.DataFrame(columns=output_cols)
    )
    reserved_df = (
        pd.DataFrame(reserved_rows, columns=output_cols)
        if reserved_rows else pd.DataFrame(columns=output_cols)
    )

    # -- Reserved summary -----------------------------------------------------
    if not reserved_df.empty:
        reserved_summary = (
            reserved_df.groupby(["Constituency Name", "RO/ARO's Office"])
            .size()
            .reset_index(name="Reserved EVMs")
        )
    else:
        reserved_summary = pd.DataFrame(
            columns=["Constituency Name", "RO/ARO's Office", "Reserved EVMs"]
        )

    return assigned_df, reserved_df, reserved_summary, animation_frames, warnings


def get_phase2_excel(assigned_df: pd.DataFrame, reserved_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        assigned_df.to_excel(writer, sheet_name="Assigned", index=False)
        reserved_df.to_excel(writer, sheet_name="Reserved", index=False)
    return buf.getvalue()
