import pandas as pd
import random
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak


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


def _drop_box_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove any Box No. style columns from output/display DataFrames."""
    if df is None or df.empty:
        return df

    drop_cols = []
    for col in df.columns:
        norm = str(col).strip().lower()
        norm = norm.replace("_", " ").replace("-", " ").replace(".", " ")
        norm = " ".join(norm.split())
        if "box" in norm and ("no" in norm or "number" in norm):
            drop_cols.append(col)

    return df.drop(columns=drop_cols, errors="ignore")

def _get_pdf_meta(df: pd.DataFrame):
    if df is None or df.empty:
        return "N/A", "N/A"

    const_values = [str(x) for x in df.get("Constituency Name", pd.Series(dtype=str)).dropna().unique() if str(x).strip()]
    ro_values = [str(x) for x in df.get("RO/ARO's Office", pd.Series(dtype=str)).dropna().unique() if str(x).strip()]

    constituency = const_values[0] if len(const_values) == 1 else "Multiple"
    ro_office = ro_values[0] if len(ro_values) == 1 else "Multiple"

    return constituency, ro_office


def _df_to_pdf_table(
    df: pd.DataFrame,
    title: str,
    randomization_date: str,
    constituency: str = None,
    ro_office: str = None,
    deo_name: str = None,
):
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<b>Date of Randomization:</b> {randomization_date}", styles["Normal"]))

    if deo_name is not None:
        story.append(Paragraph(f"<b>DEO:</b> {deo_name}", styles["Normal"]))
    else:
        story.append(Paragraph(f"<b>Concerned Constituency:</b> {constituency}", styles["Normal"]))
        story.append(Paragraph(f"<b>RO Office:</b> {ro_office}", styles["Normal"]))

    story.append(Spacer(1, 10))

    if df is None or df.empty:
        story.append(Paragraph("No data available.", styles["Normal"]))
        return story

    df = df.fillna("")
    data = [list(df.columns)] + df.astype(str).values.tolist()

    page_width = landscape(A4)[0] - 40

    weights = []
    for col in df.columns:
        col_norm = str(col).strip().lower()

        if col_norm in ["slno", "sl no", "s. no.", "s.no.", "polling station no."]:
            weights.append(0.7)
        elif col_norm in ["bu unit id", "cu unit id"]:
            weights.append(1.0)
        elif col_norm in ["ro/aro's office", "constituency name", "polling station name"]:
            weights.append(2.2)
        else:
            weights.append(1.2)

    total_weight = sum(weights)
    col_widths = [(page_width * w / total_weight) for w in weights]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("LEADING", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
    ]))

    story.append(table)
    return story


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

    bu_available_count = (bu_df["Status"] == "available for use").sum()
    cu_available_count = (cu_df["Status"] == "available for use").sum()
    total_ps = len(ps_df)

    if total_ps == 0:
        raise ValueError("Total polling station count is zero.")

    max_units = max(bu_available_count, cu_available_count)
    max_percentage = round((max_units / total_ps) * 100, 2)

    return {
        "total_polling_stations": int(total_ps),
        "available_bus": int(bu_available_count),
        "available_cus": int(cu_available_count),
        "max_units": int(max_units),
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

    if percentage > 0 and required_pairs == 0:
        required_pairs = 1

    required_pairs = min(required_pairs, len(records))
    return records[:required_pairs]


# ---------------------------------------------------------------------------
# Phase 1
# ---------------------------------------------------------------------------

def run_phase1(ps_file, bu_file, cu_file, seed=None, use_all_available=True, evm_percentage=None):
    """Randomly assign available BUs and CUs to constituencies."""
    if seed is not None:
        random.seed(seed)

    warnings = []

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

    bu_box_col = _find_box_column(bu_df)
    cu_box_col = _find_box_column(cu_df)

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

    if bu_box_col:
        bu_available = bu_available.rename(columns={bu_box_col: "BU Box No."})
    if cu_box_col:
        cu_available = cu_available.rename(columns={cu_box_col: "CU Box No."})

    constituency_info = (
        ps_df[["RO/ARO's Office", "Constituency Name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    constituencies = constituency_info["Constituency Name"].tolist()
    n_const = len(constituencies)

    if n_const == 0:
        raise ValueError("No constituencies found in Polling Station file.")

    bu_records = bu_available.to_dict(orient="records")
    cu_records = cu_available.to_dict(orient="records")
    random.shuffle(bu_records)
    random.shuffle(cu_records)

    ps_counts = (
        ps_df.groupby("Constituency Name")
        .size()
        .reindex(constituencies, fill_value=0)
    )
    total_ps = ps_counts.sum()

    if total_ps == 0:
        raise ValueError("Total polling station count is zero. Cannot compute ratios.")

    max_units_possible = max(len(bu_records), len(cu_records))
    max_percentage_possible = int((max_units_possible / total_ps) * 100)

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

    bu_allocation = _allocate_by_ratio(bu_records, ps_counts, constituencies, total_ps)
    cu_allocation = _allocate_by_ratio(cu_records, ps_counts, constituencies, total_ps)

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

    result_rows = []
    evm_sl = 1

    for const_name in constituencies:
        ro_office = constituency_info.loc[
            constituency_info["Constituency Name"] == const_name,
            "RO/ARO's Office"
        ].values[0]

        bu_const = bu_assign_df[bu_assign_df["Constituency Name"] == const_name].reset_index(drop=True)
        cu_const = cu_assign_df[cu_assign_df["Constituency Name"] == const_name].reset_index(drop=True)

        if len(bu_const) != len(cu_const):
            warnings.append(
                f"Constituency '{const_name}': BU count ({len(bu_const)}) != CU count ({len(cu_const)})."
            )

        n_rows = max(len(bu_const), len(cu_const))

        for i in range(n_rows):
            row = {
                "SlNo": evm_sl,
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "BU Unit ID": bu_const.at[i, "BU Unit ID"] if i < len(bu_const) else None,
                "CU Unit ID": cu_const.at[i, "CU Unit ID"] if i < len(cu_const) else None,
            }

            if "BU Box No." in bu_const.columns:
                row["BU Box No."] = bu_const.at[i, "BU Box No."] if i < len(bu_const) else None

            if "CU Box No." in cu_const.columns:
                row["CU Box No."] = cu_const.at[i, "CU Box No."] if i < len(cu_const) else None

            result_rows.append(row)
            evm_sl += 1

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

    summary_rows = []
    for const_name in constituencies:
        sub = result_df[result_df["Constituency Name"] == const_name]
        ps_count = int(ps_counts[const_name])

        bu_assigned = sub["BU Unit ID"].notna().sum() if not sub.empty else 0
        cu_assigned = sub["CU Unit ID"].notna().sum() if not sub.empty else 0
        assigned_pct = round((max(bu_assigned, cu_assigned) / ps_count) * 100, 2) if ps_count else 0

        summary_rows.append({
            "Constituency Name": const_name,
            "RO/ARO's Office": sub["RO/ARO's Office"].iloc[0] if not sub.empty else "",
            "No. of Polling Stations": ps_count,
            "No. of BUs Assigned": int(bu_assigned),
            "No. of CUs Assigned": int(cu_assigned),
            "Assigned EVM vs PS (%)": assigned_pct,
        })

    constituency_summary = pd.DataFrame(summary_rows)
    animation_frames = _build_phase1_frames(bu_records, cu_records, constituencies, n_snaps=10)

    assigned_bu_ids = set(result_df["BU Unit ID"].dropna()) if not result_df.empty else set()
    assigned_cu_ids = set(result_df["CU Unit ID"].dropna()) if not result_df.empty else set()

    remaining_bu_df = bu_available[~bu_available["Unit ID"].isin(assigned_bu_ids)].copy()
    remaining_cu_df = cu_available[~cu_available["Unit ID"].isin(assigned_cu_ids)].copy()

    result_df = _drop_box_columns(result_df)
    remaining_bu_df = _drop_box_columns(remaining_bu_df)
    remaining_cu_df = _drop_box_columns(remaining_cu_df)

    return result_df, constituency_summary, animation_frames, warnings, remaining_bu_df, remaining_cu_df


def _build_phase1_frames(bu_records, cu_records, constituencies, n_snaps=10):
    """Return list of DataFrames showing progressive assignment for animation."""
    frames = []
    total = min(len(bu_records), len(cu_records))
    step = max(1, total // n_snaps) if total else 1
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


def get_phase1_excel(result_df, remaining_bu_df=None, remaining_cu_df=None):
    output = BytesIO()

    result_df = _drop_box_columns(result_df)
    remaining_bu_df = _drop_box_columns(remaining_bu_df)
    remaining_cu_df = _drop_box_columns(remaining_cu_df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Assigned EVMs")

        if remaining_bu_df is not None and not remaining_bu_df.empty:
            remaining_bu_df.to_excel(writer, index=False, sheet_name="Remaining BUs")

        if remaining_cu_df is not None and not remaining_cu_df.empty:
            remaining_cu_df.to_excel(writer, index=False, sheet_name="Remaining CUs")

    return output.getvalue()

def get_phase1_pdf(result_df, remaining_bu_df=None, remaining_cu_df=None, randomization_date=""):
    buf = BytesIO()

    deo_name = "District Election Officer"

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    story = []
    story.extend(_df_to_pdf_table(
        result_df,
        "Phase 1 - Assigned EVMs",
        randomization_date,
        deo_name=deo_name,
    ))

    if remaining_bu_df is not None:
        story.append(PageBreak())
        story.extend(_df_to_pdf_table(
            remaining_bu_df,
            "Phase 1 - Remaining BUs",
            randomization_date,
            deo_name=deo_name,
        ))

    if remaining_cu_df is not None:
        story.append(PageBreak())
        story.extend(_df_to_pdf_table(
            remaining_cu_df,
            "Phase 1 - Remaining CUs",
            randomization_date,
            deo_name=deo_name,
        ))

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Phase 2
# ---------------------------------------------------------------------------

def run_phase2(phase1_file, ps_file, seed=None):
    """Run Phase 2 randomization with separate BU/CU reserved counts."""
    if seed is not None:
        random.seed(seed)

    warnings = []

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

    constituencies = phase1_df["Constituency Name"].dropna().unique().tolist()

    assigned_rows = []
    reserved_rows = []
    animation_frames = []
    assigned_sl = 1
    reserved_sl = 1

    for const_name in constituencies:
        evm_sub = phase1_df[phase1_df["Constituency Name"] == const_name].copy().reset_index(drop=True)
        ps_sub = ps_df[ps_df["Constituency Name"] == const_name].copy().reset_index(drop=True)

        if evm_sub.empty:
            continue

        ro_office = evm_sub["RO/ARO's Office"].iloc[0]

        bu_cols = ["BU Unit ID"] + (["BU Box No."] if has_bu_box else [])
        cu_cols = ["CU Unit ID"] + (["CU Box No."] if has_cu_box else [])

        valid_bu_mask = evm_sub["BU Unit ID"].notna() & (evm_sub["BU Unit ID"].astype(str).str.strip() != "")
        valid_cu_mask = evm_sub["CU Unit ID"].notna() & (evm_sub["CU Unit ID"].astype(str).str.strip() != "")

        bu_records = evm_sub.loc[valid_bu_mask, bu_cols].to_dict(orient="records")
        cu_records = evm_sub.loc[valid_cu_mask, cu_cols].to_dict(orient="records")

        random.shuffle(bu_records)
        random.shuffle(cu_records)

        ps_list = ps_sub[["Polling Station No.", "Polling Station Name"]].values.tolist()

        n_bu = len(bu_records)
        n_cu = len(cu_records)
        n_ps = len(ps_list)
        n_assign = min(n_bu, n_cu, n_ps)

        if n_ps > n_assign:
            warnings.append(
                f"Constituency '{const_name}': More polling stations ({n_ps}) than complete BU-CU sets ({n_assign}). "
                f"Some polling stations will not get an EVM."
            )

        reserved_bu_count = max(0, n_bu - n_assign)
        reserved_cu_count = max(0, n_cu - n_assign)
        if reserved_bu_count or reserved_cu_count:
            warnings.append(
                f"Constituency '{const_name}': Reserved BU={reserved_bu_count}, Reserved CU={reserved_cu_count}."
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

        for i in range(n_assign):
            row = {
                "SlNo": assigned_sl,
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
            assigned_sl += 1

        for i in range(n_assign, n_bu):
            row = {
                "SlNo": reserved_sl,
                "BU Unit ID": bu_records[i]["BU Unit ID"],
                "CU Unit ID": None,
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": "NA",
                "Polling Station Name": "RESERVED",
            }
            if has_bu_box:
                row["BU Box No."] = bu_records[i].get("BU Box No.")
            if has_cu_box:
                row["CU Box No."] = None
            reserved_rows.append(row)
            reserved_sl += 1

        for i in range(n_assign, n_cu):
            row = {
                "SlNo": reserved_sl,
                "BU Unit ID": None,
                "CU Unit ID": cu_records[i]["CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": "NA",
                "Polling Station Name": "RESERVED",
            }
            if has_bu_box:
                row["BU Box No."] = None
            if has_cu_box:
                row["CU Box No."] = cu_records[i].get("CU Box No.")
            reserved_rows.append(row)
            reserved_sl += 1

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
        "Polling Station Name",
    ])

    assigned_df = pd.DataFrame(assigned_rows, columns=output_cols) if assigned_rows else pd.DataFrame(columns=output_cols)
    reserved_df = pd.DataFrame(reserved_rows, columns=output_cols) if reserved_rows else pd.DataFrame(columns=output_cols)

    ps_counts_p2 = (
        ps_df.groupby(["Constituency Name", "RO/ARO's Office"])
        .size()
        .reset_index(name="No. of Polling Stations")
    )

    summary_base = (
        phase1_df[["Constituency Name", "RO/ARO's Office"]]
        .drop_duplicates()
        .merge(ps_counts_p2, on=["Constituency Name", "RO/ARO's Office"], how="left")
    )
    summary_base["No. of Polling Stations"] = summary_base["No. of Polling Stations"].fillna(0).astype(int)

    assigned_summary = (
        assigned_df.groupby(["Constituency Name", "RO/ARO's Office"]).agg(
            **{
                "Assigned BU": ("BU Unit ID", lambda s: s.notna().sum()),
                "Assigned CU": ("CU Unit ID", lambda s: s.notna().sum()),
            }
        ).reset_index()
        if not assigned_df.empty else pd.DataFrame(columns=["Constituency Name", "RO/ARO's Office", "Assigned BU", "Assigned CU"])
    )

    reserved_summary_counts = (
        reserved_df.groupby(["Constituency Name", "RO/ARO's Office"]).agg(
            **{
                "Reserved BU": ("BU Unit ID", lambda s: s.notna().sum()),
                "Reserved CU": ("CU Unit ID", lambda s: s.notna().sum()),
            }
        ).reset_index()
        if not reserved_df.empty else pd.DataFrame(columns=["Constituency Name", "RO/ARO's Office", "Reserved BU", "Reserved CU"])
    )

    reserved_summary = summary_base.merge(
        assigned_summary,
        on=["Constituency Name", "RO/ARO's Office"],
        how="left",
    ).merge(
        reserved_summary_counts,
        on=["Constituency Name", "RO/ARO's Office"],
        how="left",
    )

    for col in ["Assigned BU", "Assigned CU", "Reserved BU", "Reserved CU"]:
        reserved_summary[col] = reserved_summary[col].fillna(0).astype(int)

    reserved_summary = reserved_summary[[
        "Constituency Name",
        "RO/ARO's Office",
        "No. of Polling Stations",
        "Assigned BU",
        "Assigned CU",
        "Reserved BU",
        "Reserved CU",
    ]]

    assigned_df = _drop_box_columns(assigned_df)
    reserved_df = _drop_box_columns(reserved_df)

    return assigned_df, reserved_df, reserved_summary, animation_frames, warnings


def get_phase2_excel(assigned_df: pd.DataFrame, reserved_df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    assigned_df = _drop_box_columns(assigned_df)
    reserved_df = _drop_box_columns(reserved_df)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        assigned_df.to_excel(writer, sheet_name="Assigned", index=False)
        reserved_df.to_excel(writer, sheet_name="Reserved", index=False)
    return buf.getvalue()

def get_phase2_pdf(assigned_df: pd.DataFrame, reserved_df: pd.DataFrame, randomization_date="") -> bytes:
    buf = BytesIO()

    constituency, ro_office = _get_pdf_meta(assigned_df if assigned_df is not None and not assigned_df.empty else reserved_df)

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    story = []
    story.extend(_df_to_pdf_table(
        assigned_df,
        "Phase 2 - Assigned EVMs",
        randomization_date,
        constituency,
        ro_office,
    ))

    story.append(PageBreak())
    story.extend(_df_to_pdf_table(
        reserved_df,
        "Phase 2 - Reserved EVMs",
        randomization_date,
        constituency,
        ro_office,
    ))

    doc.build(story)
    return buf.getvalue()
