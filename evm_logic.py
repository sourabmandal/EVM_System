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

def run_phase1(ps_file, bu_file, cu_file, seed: int | None = None):
    """
    Randomise BUs and CUs across constituencies.

    Returns
    -------
    result_df        : final output DataFrame
    constituency_summary : DataFrame with constituency-wise BU / CU counts
    animation_frames : list of DataFrames (snapshots for animation preview)
    warnings         : list of warning strings
    """
    if seed is not None:
        random.seed(seed)

    warnings = []

    # -- Read inputs ----------------------------------------------------------
    ps_df = _read_excel(ps_file)
    bu_df = _read_excel(bu_file)
    cu_df = _read_excel(cu_file)

    # Validate required columns
    required_ps = {"Slno", "RO/ARO's Office", "Constituency Name",
                   "Polling Station No.", "Polling Station Name"}
    required_bu = {"S. No.", "Unit ID", "Box No.", "Status"}
    required_cu = {"S. No.", "Unit ID", "Box No.", "Status"}

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

    bu_available = bu_df[bu_df["Status"] == "available for use"].copy().reset_index(drop=True)
    cu_available = cu_df[cu_df["Status"] == "available for use"].copy().reset_index(drop=True)

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
    bu_indices = list(range(len(bu_available)))
    cu_indices = list(range(len(cu_available)))
    random.shuffle(bu_indices)
    random.shuffle(cu_indices)

    bu_shuffled = bu_available.iloc[bu_indices].reset_index(drop=True)
    cu_shuffled = cu_available.iloc[cu_indices].reset_index(drop=True)

    # -- Assign BUs to constituencies round-robin -----------------------------
    bu_assignments = []
    for idx, row in bu_shuffled.iterrows():
        const_name = constituencies[idx % n_const]
        ro_office = constituency_info.loc[
            constituency_info["Constituency Name"] == const_name,
            "RO/ARO's Office"
        ].values[0]
        bu_assignments.append({
            "Constituency Name": const_name,
            "RO/ARO's Office": ro_office,
            "BU Unit ID": row["Unit ID"],
            "BU Box No.": row["Box No."],
        })

    # -- Assign CUs to constituencies round-robin -----------------------------
    cu_assignments = []
    for idx, row in cu_shuffled.iterrows():
        const_name = constituencies[idx % n_const]
        cu_assignments.append({
            "Constituency Name": const_name,
            "CU Unit ID": row["Unit ID"],
            "CU Box No.": row["Box No."],
        })

    bu_assign_df = pd.DataFrame(bu_assignments)
    cu_assign_df = pd.DataFrame(cu_assignments)

    # Group by constituency and pair BU + CU entries
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
                "BU Box No.": bu_const.at[i, "BU Box No."],
                "CU Unit ID": cu_const.at[i, "CU Unit ID"],
                "CU Box No.": cu_const.at[i, "CU Box No."],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
            })
            evm_sl += 1

    result_df = pd.DataFrame(result_rows, columns=[
        "EVMSlNo", "BU Unit ID", "BU Box No.",
        "CU Unit ID", "CU Box No.",
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
    animation_frames = _build_phase1_frames(bu_available, cu_available, constituencies, n_const=10)

    return result_df, constituency_summary, animation_frames, warnings


def _build_phase1_frames(bu_df, cu_df, constituencies, n_const=10):
    """Return list of DataFrames showing progressive assignment for animation."""
    frames = []
    total = min(len(bu_df), len(cu_df))
    step = max(1, total // n_const)

    for snap in range(1, n_const + 1):
        count = min(snap * step, total)
        rows = []
        for i in range(count):
            rows.append({
                "BU Unit ID": bu_df.at[i, "Unit ID"],
                "CU Unit ID": cu_df.at[i, "Unit ID"],
                "Constituency": constituencies[i % len(constituencies)],
            })
        frames.append(pd.DataFrame(rows))

    return frames


def get_phase1_excel(result_df: pd.DataFrame) -> bytes:
    return _to_excel_bytes(result_df)


# ---------------------------------------------------------------------------
# Phase 2
# ---------------------------------------------------------------------------

def run_phase2(phase1_file, ps_file, seed: int | None = None):
    """
    Randomly assign polling stations to EVMs within each constituency.

    Returns
    -------
    assigned_df      : EVMs that got a polling station
    reserved_df      : EVMs that are reserved (no polling station)
    reserved_summary : constituency-wise reserved count
    animation_frames : list of dicts for animation
    warnings         : list of warning strings
    """
    if seed is not None:
        random.seed(seed)

    warnings = []

    # -- Read inputs ----------------------------------------------------------
    phase1_df = _read_excel(phase1_file)
    ps_df = _read_excel(ps_file)

    required_p1 = {"EVMSlNo", "BU Unit ID", "BU Box No.", "CU Unit ID",
                   "CU Box No.", "RO/ARO's Office", "Constituency Name"}
    required_ps = {"Slno", "RO/ARO's Office", "Constituency Name",
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

    for const_name in constituencies:
        evm_sub = phase1_df[phase1_df["Constituency Name"] == const_name].copy().reset_index(drop=True)
        ps_sub = ps_df[ps_df["Constituency Name"] == const_name].copy().reset_index(drop=True)

        n_evms = len(evm_sub)
        n_ps = len(ps_sub)

        # Shuffle polling stations
        ps_indices = list(range(n_ps))
        random.shuffle(ps_indices)
        ps_shuffled = ps_sub.iloc[ps_indices].reset_index(drop=True)

        # Shuffle EVMs
        evm_indices = list(range(n_evms))
        random.shuffle(evm_indices)
        evm_shuffled = evm_sub.iloc[evm_indices].reset_index(drop=True)

        if n_ps > n_evms:
            warnings.append(
                f"Constituency '{const_name}': More polling stations ({n_ps}) "
                f"than EVMs ({n_evms}). Some polling stations will not get an EVM."
            )
        if n_evms > n_ps:
            warnings.append(
                f"Constituency '{const_name}': More EVMs ({n_evms}) "
                f"than polling stations ({n_ps}). {n_evms - n_ps} EVMs will be reserved."
            )

        n_assign = min(n_evms, n_ps)

        # Animation snapshot for this constituency
        animation_frames.append({
            "constituency": const_name,
            "evms": evm_shuffled[["EVMSlNo", "BU Unit ID", "CU Unit ID"]].head(n_assign).to_dict("records"),
            "polling_stations": ps_shuffled[["Polling Station No.", "Polling Station Name"]].head(n_assign).to_dict("records"),
        })

        # Assigned EVMs
        for i in range(n_assign):
            row = evm_shuffled.iloc[i].to_dict()
            row["Polling Station No."] = ps_shuffled.at[i, "Polling Station No."]
            row["Polling Station Name"] = ps_shuffled.at[i, "Polling Station Name"]
            assigned_rows.append(row)

        # Reserved EVMs
        for i in range(n_assign, n_evms):
            row = evm_shuffled.iloc[i].to_dict()
            row["Polling Station No."] = ""
            row["Polling Station Name"] = "RESERVED"
            reserved_rows.append(row)

    output_cols = [
        "EVMSlNo", "BU Unit ID", "BU Box No.",
        "CU Unit ID", "CU Box No.",
        "RO/ARO's Office", "Constituency Name",
        "Polling Station No.", "Polling Station Name"
    ]

    assigned_df = pd.DataFrame(assigned_rows, columns=output_cols) if assigned_rows else pd.DataFrame(columns=output_cols)
    reserved_df = pd.DataFrame(reserved_rows, columns=output_cols) if reserved_rows else pd.DataFrame(columns=output_cols)

    # -- Reserved summary -----------------------------------------------------
    if not reserved_df.empty:
        reserved_summary = (
            reserved_df.groupby(["Constituency Name", "RO/ARO's Office"])
            .size()
            .reset_index(name="Reserved EVMs")
        )
    else:
        reserved_summary = pd.DataFrame(columns=["Constituency Name", "RO/ARO's Office", "Reserved EVMs"])

    return assigned_df, reserved_df, reserved_summary, animation_frames, warnings


def get_phase2_excel(assigned_df: pd.DataFrame, reserved_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        assigned_df.to_excel(writer, sheet_name="Assigned", index=False)
        reserved_df.to_excel(writer, sheet_name="Reserved", index=False)
    return buf.getvalue()
