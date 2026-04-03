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


def _find_box_column(df: pd.DataFrame):
    """
    Return the actual Box No. column name if present, else None.
    Accept flexible variants like:
      Box No., Box No, Box Number, BoxNo, BOX NO, Box_No, etc.
    Also supports BU/CU-prefixed box headers.
    """
    for col in df.columns:
        norm = str(col).strip().lower()
        norm = norm.replace("_", " ").replace("-", " ").replace(".", " ")
        norm = " ".join(norm.split())

        accepted = {
            "box no",
            "box number",
            "boxno",
            "bu box no",
            "bu box number",
            "cu box no",
            "cu box number",
        }

        if norm in accepted:
            return col

    return None


def _allocate_by_ratio(items, ps_counts, constituencies, total_ps):
    """
    Distribute items across constituencies proportional to polling station
    counts using the largest remainder method to handle rounding.

    items can be a list of dicts / records, not just IDs.
    Returns a dict: constituency_name -> list of items
    """
    n_items = len(items)
    raw = {c: (ps_counts[c] / total_ps) * n_items for c in constituencies}
    floored = {c: int(raw[c]) for c in constituencies}
    remainders = {c: raw[c] - floored[c] for c in constituencies}
    shortfall = n_items - sum(floored.values())

    sorted_by_remainder = sorted(constituencies, key=lambda c: remainders[c], reverse=True)
    for i in range(shortfall):
        floored[sorted_by_remainder[i]] += 1

    allocation = {}
    idx = 0
    for c in constituencies:
        count = floored[c]
        allocation[c] = items[idx: idx + count]
        idx += count

    return allocation

def get_phase1_capacity(ps_file, bu_file, cu_file):
    """
    Return Phase 1 capacity details so UI can show max assignable percentage.
    """
    ps_df = _read_excel(ps_file)
    bu_df = _read_excel(bu_file)
    cu_df = _read_excel(cu_file)

    required_ps = {"RO/ARO's Office", "Constituency Name", "Polling Station No.", "Polling Station Name"}
    required_bu = {"Unit ID", "Status"}
    required_cu = {"Unit ID", "Status"}

    missing_ps = required_ps - set(ps_df.columns)
    missing_bu = required_bu - set(bu_df.columns)
    missing_cu = required_cu - set(cu_df.columns)

    if missing_ps or missing_bu or missing_cu:
        raise ValueError(
            f"Missing columns -- PS: {missing_ps}, BU: {missing_bu}, CU: {missing_cu}"
        )

    bu_df = _normalise_status(bu_df)
    cu_df = _normalise_status(cu_df)

    # Explicitly remove rejected / training entries if you have added that helper
    if "_remove_rejected_training" in globals():
        bu_df = _remove_rejected_training(bu_df)
        cu_df = _remove_rejected_training(cu_df)

    bu_available_count = (bu_df["Status"] == "available for use").sum()
    cu_available_count = (cu_df["Status"] == "available for use").sum()
    total_ps = len(ps_df)

    if total_ps == 0:
        raise ValueError("Total polling station count is zero.")

    max_pairs = min(bu_available_count, cu_available_count)
    max_percentage = round((max_pairs / total_ps) * 100, 2)

    return {
        "total_polling_stations": int(total_ps),
        "available_bus": int(bu_available_count),
        "available_cus": int(cu_available_count),
        "max_pairs": int(max_pairs),
        "max_percentage": float(max_percentage),
    }


def _cap_records_to_percentage(records, total_ps, percentage):
    """
    Keep only the number of records needed for the requested percentage
    of polling stations.
    """
    if percentage is None:
        return records

    required_pairs = int((percentage / 100) * total_ps)

    # if percentage > 0 but rounding gives 0, allocate at least 1 pair
    if percentage > 0 and required_pairs == 0:
        required_pairs = 1

    required_pairs = min(required_pairs, len(records))
    return records[:required_pairs]


# ---------------------------------------------------------------------------
# Phase 1
# ---------------------------------------------------------------------------

def run_phase1(ps_file, bu_file, cu_file, seed=None, use_all_available=True, evm_percentage=None):
    """
    Randomly assign available BUs and CUs to constituencies.

    Input columns:
      PS file : RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name
      BU file : Unit ID, Status, [optional Box No.]
      CU file : Unit ID, Status, [optional Box No.]

    Output columns:
      SlNo, BU Unit ID, [BU Box No.], CU Unit ID, [CU Box No.],
      RO/ARO's Office, Constituency Name

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

    # Validate required columns
    required_ps = {"RO/ARO's Office", "Constituency Name", "Polling Station No.", "Polling Station Name"}
    required_bu = {"Unit ID", "Status"}
    required_cu = {"Unit ID", "Status"}

    missing_ps = required_ps - set(ps_df.columns)
    missing_bu = required_bu - set(bu_df.columns)
    missing_cu = required_cu - set(cu_df.columns)

    if missing_ps or missing_bu or missing_cu:
        raise ValueError(
            f"Missing columns -- PS: {missing_ps}, BU: {missing_bu}, CU: {missing_cu}"
        )

    # Detect optional box columns
    bu_box_col = _find_box_column(bu_df)
    cu_box_col = _find_box_column(cu_df)

    # -- Filter only 'available for use' units --------------------------------
    bu_df = _normalise_status(bu_df)
    cu_df = _normalise_status(cu_df)

    bu_keep_cols = ["Unit ID"]
    if bu_box_col:
        bu_keep_cols.append(bu_box_col)

    cu_keep_cols = ["Unit ID"]
    if cu_box_col:
        cu_keep_cols.append(cu_box_col)

    bu_available = (
        bu_df[bu_df["Status"] == "available for use"][bu_keep_cols]
        .copy()
        .reset_index(drop=True)
    )
    cu_available = (
        cu_df[cu_df["Status"] == "available for use"][cu_keep_cols]
        .copy()
        .reset_index(drop=True)
    )

    if bu_available.empty:
        raise ValueError("No BUs with status 'available for use' found.")
    if cu_available.empty:
        raise ValueError("No CUs with status 'available for use' found.")

    # Rename optional box columns to standard output names
    if bu_box_col:
        bu_available = bu_available.rename(columns={bu_box_col: "BU Box No."})
    if cu_box_col:
        cu_available = cu_available.rename(columns={cu_box_col: "CU Box No."})

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

    # -- Shuffle BU and CU records independently ------------------------------
    bu_records = bu_available.to_dict(orient="records")
    cu_records = cu_available.to_dict(orient="records")
    random.shuffle(bu_records)
    random.shuffle(cu_records)

        # -- Compute polling station counts per constituency ----------------------
    ps_counts = (
        ps_df.groupby("Constituency Name")
        .size()
        .reindex(constituencies, fill_value=0)
    )
    total_ps = ps_counts.sum()

    if total_ps == 0:
        raise ValueError("Total polling station count is zero. Cannot compute ratios.")

    # -- Limit usable pool if percentage mode is selected ---------------------
    max_pairs_possible = min(len(bu_records), len(cu_records))
    max_percentage_possible = round((max_pairs_possible / total_ps) * 100, 2)
    
    if not use_all_available:
        if evm_percentage is None:
            raise ValueError("EVM percentage is required when 'Use all available' is not selected.")
    
        evm_percentage = round(float(evm_percentage), 2)
    
        if evm_percentage <= 0:
            raise ValueError("EVM percentage must be greater than 0.")
    
        if evm_percentage > max_percentage_possible:
            raise ValueError(
                f"Requested percentage ({evm_percentage}%) exceeds maximum possible "
                f"percentage ({max_percentage_possible}%)."
            )

        bu_records = _cap_records_to_percentage(bu_records, total_ps, evm_percentage)
        cu_records = _cap_records_to_percentage(cu_records, total_ps, evm_percentage)

        warnings.append(
            f"Percentage mode selected: only enough units for {evm_percentage}% of polling stations "
            f"have been included in Phase 1 randomization."
        )

    # -- Allocate BUs and CUs by polling station ratio ------------------------
    bu_allocation = _allocate_by_ratio(bu_records, ps_counts, constituencies, total_ps)
    cu_allocation = _allocate_by_ratio(cu_records, ps_counts, constituencies, total_ps)

    # -- Build assignment DataFrames from allocations -------------------------
    bu_assignments = []
    for const_name in constituencies:
        ro_office = constituency_info.loc[
            constituency_info["Constituency Name"] == const_name,
            "RO/ARO's Office"
        ].values[0]

        for rec in bu_allocation[const_name]:
            row = {
                "Constituency Name": const_name,
                "RO/ARO's Office": ro_office,
                "BU Unit ID": rec["Unit ID"],
            }
            if "BU Box No." in rec:
                row["BU Box No."] = rec["BU Box No."]
            bu_assignments.append(row)

    cu_assignments = []
    for const_name in constituencies:
        for rec in cu_allocation[const_name]:
            row = {
                "Constituency Name": const_name,
                "CU Unit ID": rec["Unit ID"],
            }
            if "CU Box No." in rec:
                row["CU Box No."] = rec["CU Box No."]
            cu_assignments.append(row)

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
            row = {
                "SlNo": evm_sl,
                "BU Unit ID": bu_const.at[i, "BU Unit ID"],
                "CU Unit ID": cu_const.at[i, "CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
            }

            if "BU Box No." in bu_const.columns:
                row["BU Box No."] = bu_const.at[i, "BU Box No."]
            if "CU Box No." in cu_const.columns:
                row["CU Box No."] = cu_const.at[i, "CU Box No."]

            result_rows.append(row)
            evm_sl += 1

    # Build output columns dynamically
    output_cols = ["SlNo", "BU Unit ID"]
    if bu_box_col:
        output_cols.append("BU Box No.")
    output_cols.append("CU Unit ID")
    if cu_box_col:
        output_cols.append("CU Box No.")
    output_cols.extend(["RO/ARO's Office", "Constituency Name"])

    result_df = pd.DataFrame(result_rows)
    if result_df.empty:
        result_df = pd.DataFrame(columns=output_cols)
    else:
        result_df = result_df[output_cols]

    # -- Constituency summary -------------------------------------------------
    summary_rows = []
    for const_name in constituencies:
        sub = result_df[result_df["Constituency Name"] == const_name]
        ps_count = int(ps_counts[const_name])
        assigned_count = len(sub)
        assigned_pct = round((assigned_count / ps_count) * 100, 2) if ps_count else 0

        summary_rows.append({
            "Constituency Name": const_name,
            "RO/ARO's Office": sub["RO/ARO's Office"].iloc[0] if not sub.empty else "",
            "No. of Polling Stations": ps_count,
            "No. of BUs Assigned": assigned_count,
            "No. of CUs Assigned": assigned_count,
            "Assigned EVM vs PS (%)": assigned_pct,
        })
    constituency_summary = pd.DataFrame(summary_rows)

    # -- Animation frames (preview only) -------------------------------------
    animation_frames = _build_phase1_frames(bu_records, cu_records, constituencies, n_snaps=10)

    # -- Remaining unused BUs / CUs -----------------------------------------
    assigned_bu_ids = set(result_df["BU Unit ID"]) if not result_df.empty else set()
    assigned_cu_ids = set(result_df["CU Unit ID"]) if not result_df.empty else set()

    remaining_bu_df = bu_available[~bu_available["Unit ID"].isin(assigned_bu_ids)].copy()
    remaining_cu_df = cu_available[~cu_available["Unit ID"].isin(assigned_cu_ids)].copy()

    return result_df, constituency_summary, animation_frames, warnings, remaining_bu_df, remaining_cu_df


def _build_phase1_frames(bu_records, cu_records, constituencies, n_snaps=10):
    """Return list of DataFrames showing progressive assignment for animation."""
    frames = []
    total = min(len(bu_records), len(cu_records))
    step = max(1, total // n_snaps)
    n_const = len(constituencies)

    for snap in range(1, n_snaps + 1):
        count = min(snap * step, total)
        rows = []
        for i in range(count):
            rows.append({
                "BU Unit ID": bu_records[i]["Unit ID"],
                "CU Unit ID": cu_records[i]["Unit ID"],
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
      - Also keep BU Box No. / CU Box No. mapped to the same BU / CU if present.
      - Pair: one polling station gets one BU and one CU.
      - Surplus EVM pairs (no polling station available) are marked RESERVED.
      - BUs and CUs are NEVER moved out of their Phase 1 assigned constituency.

    Input columns (Phase 1 output):
      EVMSlNo, BU Unit ID, [BU Box No.], CU Unit ID, [CU Box No.],
      RO/ARO's Office, Constituency Name

    Input columns (PS file):
      RO/ARO's Office, Constituency Name, Polling Station No., Polling Station Name

    Output columns:
      EVMSlNo, BU Unit ID, [BU Box No.], CU Unit ID, [CU Box No.],
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

    required_p1 = {"SlNo", "BU Unit ID", "CU Unit ID", "RO/ARO's Office", "Constituency Name"}
    required_ps = {"RO/ARO's Office", "Constituency Name", "Polling Station No.", "Polling Station Name"}

    missing_p1 = required_p1 - set(phase1_df.columns)
    missing_ps = required_ps - set(ps_df.columns)

    if missing_p1 or missing_ps:
        raise ValueError(
            f"Missing columns -- Phase1 file: {missing_p1}, PS file: {missing_ps}"
        )

    has_bu_box = "BU Box No." in phase1_df.columns
    has_cu_box = "CU Box No." in phase1_df.columns

    constituencies = phase1_df["Constituency Name"].unique().tolist()

    assigned_rows = []
    reserved_rows = []
    animation_frames = []
    new_evm_sl = 1

    for const_name in constituencies:
        evm_sub = phase1_df[phase1_df["Constituency Name"] == const_name].copy().reset_index(drop=True)
        ps_sub = ps_df[ps_df["Constituency Name"] == const_name].copy().reset_index(drop=True)

        ro_office = evm_sub["RO/ARO's Office"].iloc[0]

        # Re-shuffle BUs and CUs independently within this constituency,
        # while keeping box numbers attached to the same unit.
        bu_cols = ["BU Unit ID"] + (["BU Box No."] if has_bu_box else [])
        cu_cols = ["CU Unit ID"] + (["CU Box No."] if has_cu_box else [])

        bu_records = evm_sub[bu_cols].to_dict(orient="records")
        cu_records = evm_sub[cu_cols].to_dict(orient="records")

        random.shuffle(bu_records)
        random.shuffle(cu_records)

        # Keep polling stations in original file order
        ps_list = ps_sub[["Polling Station No.", "Polling Station Name"]].values.tolist()

        n_bu = len(bu_records)
        n_cu = len(cu_records)
        n_ps = len(ps_list)

        n_total_evms = min(n_bu, n_cu)
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

        animation_frames.append({
            "constituency": const_name,
            "evms": [
                {
                    "BU Unit ID": bu_records[i]["BU Unit ID"],
                    "CU Unit ID": cu_records[i]["CU Unit ID"],
                }
                for i in range(min(n_assign, 5))
            ],
            "polling_stations": [
                {
                    "Polling Station No.": ps_list[i][0],
                    "Polling Station Name": ps_list[i][1],
                }
                for i in range(min(n_assign, 5))
            ],
        })

        # Assigned rows
        for i in range(n_assign):
            row = {
                "SlNo": new_evm_sl,
                "BU Unit ID": bu_records[i]["BU Unit ID"],
                "CU Unit ID": cu_records[i]["CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": ps_list[i][0],
                "Polling Station Name": ps_list[i][1],
            }

            if has_bu_box:
                row["BU Box No."] = bu_records[i].get("BU Box No.")
            if has_cu_box:
                row["CU Box No."] = cu_records[i].get("CU Box No.")

            assigned_rows.append(row)
            new_evm_sl += 1

        # Reserved rows
        for i in range(n_assign, n_total_evms):
            row = {
                "EVMSlNo": new_evm_sl,
                "BU Unit ID": bu_records[i]["BU Unit ID"],
                "CU Unit ID": cu_records[i]["CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": "NA",
                "Polling Station Name": "RESERVED",
            }

            if has_bu_box:
                row["BU Box No."] = bu_records[i].get("BU Box No.")
            if has_cu_box:
                row["CU Box No."] = cu_records[i].get("CU Box No.")

            reserved_rows.append(row)
            new_evm_sl += 1

    output_cols = ["SlNo", "BU Unit ID"]
    if has_bu_box:
        output_cols.append("BU Box No.")
    output_cols.append("CU Unit ID")
    if has_cu_box:
        output_cols.append("CU Box No.")
    output_cols.extend([
        "RO/ARO's Office",
        "Constituency Name",
        "Polling Station No.",
        "Polling Station Name"
    ])

    assigned_df = (
        pd.DataFrame(assigned_rows, columns=output_cols)
        if assigned_rows else pd.DataFrame(columns=output_cols)
    )
    reserved_df = (
        pd.DataFrame(reserved_rows, columns=output_cols)
        if reserved_rows else pd.DataFrame(columns=output_cols)
    )

    # -- Reserved summary -----------------------------------------------------
    ps_counts_p2 = (
        ps_df.groupby("Constituency Name")
        .size()
        .reset_index(name="No. of Polling Stations")
    )

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

    reserved_summary = reserved_summary.merge(ps_counts_p2, on="Constituency Name", how="left")

    assigned_summary = (
        assigned_df.groupby(["Constituency Name", "RO/ARO's Office"])
        .size()
        .reset_index(name="Assigned EVMs")
    )

    reserved_summary = reserved_summary.merge(
        assigned_summary,
        on=["Constituency Name", "RO/ARO's Office"],
        how="left"
    )

    reserved_summary["Assigned EVMs"] = reserved_summary["Assigned EVMs"].fillna(0).astype(int)
    reserved_summary["Reserved EVMs"] = reserved_summary["Reserved EVMs"].fillna(0).astype(int)

    reserved_summary = reserved_summary[[
        "Constituency Name",
        "RO/ARO's Office",
        "No. of Polling Stations",
        "Assigned EVMs",
        "Reserved EVMs"
    ]]

    return assigned_df, reserved_df, reserved_summary, animation_frames, warnings


def get_phase2_excel(assigned_df: pd.DataFrame, reserved_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        assigned_df.to_excel(writer, sheet_name="Assigned", index=False)
        reserved_df.to_excel(writer, sheet_name="Reserved", index=False)
    return buf.getvalue()
