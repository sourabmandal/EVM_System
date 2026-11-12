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
    """
    Intelligently detect status columns (e.g. 'Movememt Status'), ignore empty/NaN 
    placeholder columns, and map operational statuses like 'FLC OK' to 'available for use'.
    """
    # Drop empty placeholder 'Status' column if present
    if "Status" in df.columns and df["Status"].dropna().empty:
        df = df.drop(columns=["Status"])

    # Find any column containing 'status'
    status_col = None
    for col in df.columns:
        norm = str(col).strip().lower().replace("_", " ").replace("-", " ")
        if "status" in norm:
            status_col = col
            break

    if status_col and status_col != "Status":
        df = df.rename(columns={status_col: "Status"})

    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip().str.lower()
        # Intelligent mapping for real-world EVM status terms
        status_mapping = {
            "flc ok": "available for use",
            "ok": "available for use",
            "available": "available for use",
            "ready": "available for use",
        }
        df["Status"] = df["Status"].replace(status_mapping)

    return df


def _find_box_column(df: pd.DataFrame):
    """
    Intelligently detect box columns across various naming conventions 
    such as 'Box ID', 'Box No.', 'BoxNumber', etc.
    """
    for col in df.columns:
        norm = str(col).strip().lower()
        norm = norm.replace("_", " ").replace("-", " ").replace(".", " ")
        norm = " ".join(norm.split())

        accepted = {
            "box no",
            "box number",
            "boxno",
            "box id",
            "boxid",
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
    Uses intelligent column normalization to support 'Movememt Status' and other variants.
    """
    ps_df = _read_excel(ps_file)
    bu_df = _read_excel(bu_file)
    cu_df = _read_excel(cu_file)

    required_ps = {"RO/ARO's Office", "Constituency Name", "Polling Station No.", "Polling Station Name"}
    
    missing_ps = required_ps - set(ps_df.columns)
    if missing_ps:
        raise ValueError(f"Missing columns in Polling Station file: {missing_ps}")

    # Intelligently normalize status columns and values before checking capacity
    bu_df = _normalise_status(bu_df)
    cu_df = _normalise_status(cu_df)

    if "Unit ID" not in bu_df.columns or "Status" not in bu_df.columns:
        raise ValueError("BU file must contain 'Unit ID' and a status column.")
    if "Unit ID" not in cu_df.columns or "Status" not in cu_df.columns:
        raise ValueError("CU file must contain 'Unit ID' and a status column.")

    bu_available_count = (bu_df["Status"] == "available for use").sum()
    cu_available_count = (cu_df["Status"] == "available for use").sum()
    total_ps = len(ps_df)

    if total_ps == 0:
        raise ValueError("Total polling station count is zero.")

    return {
        "total_polling_stations": int(total_ps),
        "available_bus": int(bu_available_count),
        "available_cus": int(cu_available_count),
        "max_bu_percentage": round((bu_available_count / total_ps) * 100, 2),
        "max_cu_percentage": round((cu_available_count / total_ps) * 100, 2),
        # Backward-compatible keys for callers outside this application.
        "max_units": int(max(bu_available_count, cu_available_count)),
        "max_percentage": round((max(bu_available_count, cu_available_count) / total_ps) * 100, 2),
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

def _allocate_by_exact_counts(items, counts, constituencies):
    """Allocate randomized records according to explicit constituency quotas."""
    allocations = {}
    offset = 0
    for constituency in constituencies:
        count = counts[constituency]
        allocations[constituency] = items[offset:offset + count]
        offset += count
    if offset != len(items):
        raise ValueError("The requested constituency counts do not match the selected units.")
    return allocations


def _validate_constituency_allocations(allocations, constituencies):
    """Read integer BU/CU base quotas from the editable Phase 1 allocation table."""
    table = pd.DataFrame(allocations)
    required_columns = {"Constituency Name", "Base BUs", "Base CUs"}
    if not required_columns.issubset(table.columns):
        raise ValueError("Exact allocation requires Constituency Name, Base BUs and Base CUs columns.")

    names = table["Constituency Name"].astype(str).str.strip()
    expected = set(constituencies)
    if len(table) != len(constituencies) or len(set(names)) != len(names) or set(names) != expected:
        raise ValueError("Provide exactly one allocation row for every constituency in the polling-station file.")

    counts = {"BU": {}, "CU": {}}
    for _, row in table.iterrows():
        constituency = str(row["Constituency Name"]).strip()
        for column, key in (("Base BUs", "BU"), ("Base CUs", "CU")):
            value = row[column]
            if pd.isna(value):
                raise ValueError(f"{column} cannot be empty for '{constituency}'.")
            try:
                number = float(value)
            except (TypeError, ValueError):
                raise ValueError(f"{column} must be a whole number for '{constituency}'.") from None
            if not number.is_integer() or number < 0:
                raise ValueError(f"{column} must be a non-negative whole number for '{constituency}'.")
            counts[key][constituency] = int(number)
    return counts["BU"], counts["CU"]


def run_phase1(
    ps_file, bu_file, cu_file, seed=None, use_all_available=True, evm_percentage=None,
    bu_percentage=None, cu_percentage=None, constituency_allocations=None,
    existing_phase1_file=None,
):
    """
    Randomly allocate actual BUs and CUs in Phase 1.

    Automatic mode distributes separate BU/CU totals by polling-station ratio.
    Exact mode takes independent base BU/CU quotas for each constituency.
    Phase 1 never creates or designates reserved units; that is an option in
    Phase 2. The Phase 1 export shape stays compatible with Phase 2.
    An earlier Phase 1 file may be merged without re-randomizing its Unit IDs.
    """
    rng = random.Random(seed)
    warnings = []

    ps_df = _read_excel(ps_file)
    bu_df = _normalise_status(_read_excel(bu_file))
    cu_df = _normalise_status(_read_excel(cu_file))

    required_ps = {"RO/ARO's Office", "Constituency Name", "Polling Station No.", "Polling Station Name"}
    missing_ps = required_ps - set(ps_df.columns)
    if missing_ps:
        raise ValueError(f"Missing columns in Polling Station file: {missing_ps}")
    if "Unit ID" not in bu_df.columns or "Status" not in bu_df.columns:
        raise ValueError("BU file must contain 'Unit ID' and a status column.")
    if "Unit ID" not in cu_df.columns or "Status" not in cu_df.columns:
        raise ValueError("CU file must contain 'Unit ID' and a status column.")

    bu_box_col = _find_box_column(bu_df)
    cu_box_col = _find_box_column(cu_df)
    bu_keep = ["Unit ID"] + ([bu_box_col] if bu_box_col else [])
    cu_keep = ["Unit ID"] + ([cu_box_col] if cu_box_col else [])
    bu_available = bu_df.loc[bu_df["Status"] == "available for use", bu_keep].copy().reset_index(drop=True)
    cu_available = cu_df.loc[cu_df["Status"] == "available for use", cu_keep].copy().reset_index(drop=True)

    if bu_available.empty:
        raise ValueError("No BUs with status 'available for use' found.")
    if cu_available.empty:
        raise ValueError("No CUs with status 'available for use' found.")
    if bu_box_col:
        bu_available = bu_available.rename(columns={bu_box_col: "BU Box No."})
    if cu_box_col:
        cu_available = cu_available.rename(columns={cu_box_col: "CU Box No."})

    for label, data in (("BU", bu_available), ("CU", cu_available)):
        if data["Unit ID"].isna().any() or (data["Unit ID"].astype(str).str.strip() == "").any():
            raise ValueError(f"Available {label} records must have non-empty Unit IDs.")
        if data["Unit ID"].astype(str).str.strip().duplicated().any():
            raise ValueError(f"Duplicate available {label} Unit IDs were found. Please correct the input file.")

    constituency_info = ps_df[["RO/ARO's Office", "Constituency Name"]].drop_duplicates().reset_index(drop=True)
    if constituency_info["Constituency Name"].duplicated().any():
        raise ValueError("Each constituency must map to one RO/ARO office in the Polling Station file.")
    constituencies = constituency_info["Constituency Name"].tolist()
    if not constituencies:
        raise ValueError("No constituencies found in Polling Station file.")
    ro_map = dict(zip(constituency_info["Constituency Name"], constituency_info["RO/ARO's Office"]))
    ps_counts = ps_df.groupby("Constituency Name").size().reindex(constituencies, fill_value=0)
    total_ps = int(ps_counts.sum())
    if not total_ps:
        raise ValueError("Total polling station count is zero. Cannot compute ratios.")

    existing_df = None
    if existing_phase1_file is not None:
        # Read the assigned sheet from an earlier Phase 1 workbook. Previous
        # assignments and SlNo values are preserved when the new units are appended.
        existing_df = _drop_box_columns(_read_excel(existing_phase1_file)).copy()
        required_existing = {"SlNo", "BU Unit ID", "CU Unit ID", "RO/ARO's Office", "Constituency Name"}
        missing_existing = required_existing - set(existing_df.columns)
        if missing_existing:
            raise ValueError(f"Existing Phase 1 output is missing columns: {missing_existing}")
        if existing_df.empty:
            raise ValueError("Existing Phase 1 output contains no assignments.")
        if existing_df["Constituency Name"].isna().any():
            raise ValueError("Existing Phase 1 output has blank constituency names.")
        prior_names = set(existing_df["Constituency Name"])
        if not prior_names.issubset(set(constituencies)):
            raise ValueError(
                "All constituencies in the existing Phase 1 output must also be present "
                "in the uploaded polling-station list."
            )
        for row in existing_df[["Constituency Name", "RO/ARO's Office"]].drop_duplicates().itertuples(index=False, name=None):
            name, ro = row
            if str(ro).strip() != str(ro_map[name]).strip():
                raise ValueError(f"RO/ARO office differs for existing constituency '{name}'.")
        slnos = pd.to_numeric(existing_df["SlNo"], errors="coerce")
        if slnos.isna().any() or (slnos < 1).any() or (slnos % 1 != 0).any() or slnos.duplicated().any():
            raise ValueError("Existing Phase 1 output must have unique, positive whole-number SlNo values.")
        for kind, old_column, new_data in (
            ("BU", "BU Unit ID", bu_available), ("CU", "CU Unit ID", cu_available)
        ):
            old_ids = existing_df[old_column].dropna().astype(str).str.strip()
            old_ids = old_ids[old_ids != ""]
            if old_ids.duplicated().any():
                raise ValueError(f"Existing Phase 1 output has duplicate {kind} Unit IDs.")
            clashes = set(old_ids) & set(new_data["Unit ID"].astype(str).str.strip())
            if clashes:
                raise ValueError(
                    f"The new {kind} file contains Unit IDs already assigned in the "
                    "existing Phase 1 output. Upload newly received units only."
                )

    bu_records = bu_available.to_dict(orient="records")
    cu_records = cu_available.to_dict(orient="records")
    rng.shuffle(bu_records)
    rng.shuffle(cu_records)

    base_bu_capacity = len(bu_records)
    base_cu_capacity = len(cu_records)

    if constituency_allocations is not None:
        bu_base_counts, cu_base_counts = _validate_constituency_allocations(constituency_allocations, constituencies)
        bu_base_total = sum(bu_base_counts.values())
        cu_base_total = sum(cu_base_counts.values())
    else:
        if use_all_available:
            # In Phase 1, all available units are allocated to constituencies.
            bu_base_total, cu_base_total = base_bu_capacity, base_cu_capacity
        else:
            if bu_percentage is None:
                bu_percentage = evm_percentage  # older API compatibility
            if cu_percentage is None:
                cu_percentage = evm_percentage
            if bu_percentage is None or cu_percentage is None:
                raise ValueError("Enter separate BU and CU assignment percentages.")
            bu_percentage, cu_percentage = float(bu_percentage), float(cu_percentage)
            if not (bu_percentage > 0 and cu_percentage > 0):
                raise ValueError("BU and CU assignment percentages must each be greater than zero.")
            # Keep the pre-existing floor rule when converting percentages to units.
            bu_base_total = len(_cap_records_to_percentage(bu_records, total_ps, bu_percentage))
            cu_base_total = len(_cap_records_to_percentage(cu_records, total_ps, cu_percentage))
            if int((bu_percentage / 100) * total_ps) > base_bu_capacity or int((cu_percentage / 100) * total_ps) > base_cu_capacity:
                raise ValueError(
                    "Requested BU/CU percentages exceed available inventory."
                )
            warnings.append(
                f"Percentage mode: base BUs={bu_base_total} ({bu_percentage:g}%), "
                f"base CUs={cu_base_total} ({cu_percentage:g}%)."
            )

        # Determine the base counts separately for BUs and CUs.
        def ratio_counts(n):
            return {name: len(group) for name, group in _allocate_by_ratio(
                [None] * n, ps_counts, constituencies, total_ps
            ).items()}
        bu_base_counts = ratio_counts(bu_base_total)
        cu_base_counts = ratio_counts(cu_base_total)

    if bu_base_total > base_bu_capacity or cu_base_total > base_cu_capacity:
        raise ValueError(
            f"Requested allocation needs {bu_base_total} BUs and {cu_base_total} CUs, "
            f"but only {len(bu_records)} BUs and {len(cu_records)} CUs are available."
        )
    if bu_base_total <= 0 or cu_base_total <= 0:
        raise ValueError("At least one BU and one CU must be allocated.")

    bu_allocation = _allocate_by_exact_counts(bu_records[:bu_base_total], bu_base_counts, constituencies)
    cu_allocation = _allocate_by_exact_counts(cu_records[:cu_base_total], cu_base_counts, constituencies)

    result_rows = []
    for name in constituencies:
        bu_const = bu_allocation[name]
        cu_const = cu_allocation[name]
        if len(bu_const) != len(cu_const):
            warnings.append(
                f"Constituency '{name}': BU count ({len(bu_const)}) != CU count ({len(cu_const)})."
            )
        for i in range(max(len(bu_const), len(cu_const))):
            bu = bu_const[i] if i < len(bu_const) else {}
            cu = cu_const[i] if i < len(cu_const) else {}
            row = {
                "SlNo": len(result_rows) + 1,
                "BU Unit ID": bu.get("Unit ID"),
                "CU Unit ID": cu.get("Unit ID"),
                "RO/ARO's Office": ro_map[name],
                "Constituency Name": name,
            }
            if bu_box_col:
                row["BU Box No."] = bu.get("BU Box No.")
            if cu_box_col:
                row["CU Box No."] = cu.get("CU Box No.")
            result_rows.append(row)

    output_cols = ["SlNo", "BU Unit ID"]
    if bu_box_col:
        output_cols.append("BU Box No.")
    output_cols.append("CU Unit ID")
    if cu_box_col:
        output_cols.append("CU Box No.")
    output_cols.extend(["RO/ARO's Office", "Constituency Name"])
    new_result_df = pd.DataFrame(result_rows, columns=output_cols)
    if existing_df is not None:
        # Don't change any previous allocation. Only append new records with new serials.
        existing_df = existing_df[[col for col in output_cols if col in existing_df.columns]].copy()
        added = new_result_df.copy()
        added["SlNo"] = range(int(existing_df["SlNo"].max()) + 1,
                                int(existing_df["SlNo"].max()) + 1 + len(added))
        result_df = pd.concat([existing_df, added], ignore_index=True)
        result_df = result_df.reindex(columns=output_cols)
        warnings.append(
            f"Previous Phase 1 allocations retained; {bu_base_total} new BU(s) "
            f"and {cu_base_total} new CU(s) appended."
        )
    else:
        result_df = new_result_df

    summary_rows = []
    for name in constituencies:
        ps_count = int(ps_counts[name])
        new_bu_assigned = len(bu_allocation[name])
        new_cu_assigned = len(cu_allocation[name])
        if existing_df is not None:
            prior_const = existing_df.loc[existing_df["Constituency Name"] == name]
            bu_assigned = new_bu_assigned + int(prior_const["BU Unit ID"].notna().sum())
            cu_assigned = new_cu_assigned + int(prior_const["CU Unit ID"].notna().sum())
        else:
            bu_assigned, cu_assigned = new_bu_assigned, new_cu_assigned
        summary_row = {
            "Constituency Name": name,
            "RO/ARO's Office": ro_map[name],
            "No. of Polling Stations": ps_count,
            "No. of BUs Assigned": bu_assigned,
            "No. of CUs Assigned": cu_assigned,
            "Assigned BU vs PS (%)": round(100 * bu_assigned / ps_count, 2),
            "Assigned CU vs PS (%)": round(100 * cu_assigned / ps_count, 2),
        }
        if existing_df is not None:
            summary_row["New BUs Assigned"] = new_bu_assigned
            summary_row["New CUs Assigned"] = new_cu_assigned
        summary_rows.append(summary_row)
    constituency_summary = pd.DataFrame(summary_rows)
    animation_frames = _build_phase1_frames(new_result_df, n_snaps=10)

    assigned_bu_ids = set(result_df["BU Unit ID"].dropna())
    assigned_cu_ids = set(result_df["CU Unit ID"].dropna())
    remaining_bu_df = bu_available.loc[~bu_available["Unit ID"].isin(assigned_bu_ids)].copy()
    remaining_cu_df = cu_available.loc[~cu_available["Unit ID"].isin(assigned_cu_ids)].copy()

    return (
        _drop_box_columns(result_df), constituency_summary, animation_frames, warnings,
        _drop_box_columns(remaining_bu_df), _drop_box_columns(remaining_cu_df),
    )


def _build_phase1_frames(result_df, n_snaps=10):
    """Preview real randomized constituency assignments, including unmatched units."""
    if result_df.empty:
        return []
    preview = result_df[["BU Unit ID", "CU Unit ID", "Constituency Name"]].rename(
        columns={"Constituency Name": "Constituency"}
    )
    n_frames = min(n_snaps, len(preview))
    return [preview.head((len(preview) * i + n_frames - 1) // n_frames).copy()
            for i in range(1, n_frames + 1)]


def get_phase1_excel(result_df, remaining_bu_df=None, remaining_cu_df=None, summary_df=None):
    output = BytesIO()

    result_df = _drop_box_columns(result_df)
    remaining_bu_df = _drop_box_columns(remaining_bu_df)
    remaining_cu_df = _drop_box_columns(remaining_cu_df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Assigned EVMs")
        if summary_df is not None:
            summary_df.to_excel(writer, index=False, sheet_name="Constituency Summary")

        if remaining_bu_df is not None and not remaining_bu_df.empty:
            remaining_bu_df.to_excel(writer, index=False, sheet_name="Remaining BUs")

        if remaining_cu_df is not None and not remaining_cu_df.empty:
            remaining_cu_df.to_excel(writer, index=False, sheet_name="Remaining CUs")

    return output.getvalue()

def get_phase1_pdf(result_df, remaining_bu_df=None, remaining_cu_df=None, randomization_date="", summary_df=None):
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

    if summary_df is not None:
        story.append(PageBreak())
        story.extend(_df_to_pdf_table(
            summary_df,
            "Phase 1 - Constituency Summary",
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

def run_phase2(
    phase1_file, ps_file, seed=None, selected_ro_offices=None,
    selected_constituencies=None, reserve_per_polling_station=False,
):
    """Assign EVMs to stations; optionally designate a reserved pair per station.

    Default (option off) retains Phase 2's original random assignment and surplus
    reservation. The option uses Phase 1 constituency stock, never new units.
    """
    if seed is not None:
        random.seed(seed)

    warnings = []

    phase1_df = _read_excel(phase1_file)
    ps_df = _read_excel(ps_file)

    if selected_ro_offices:
        phase1_df = phase1_df[phase1_df["RO/ARO's Office"].isin(selected_ro_offices)].copy()
        ps_df = ps_df[ps_df["RO/ARO's Office"].isin(selected_ro_offices)].copy()

    if selected_constituencies:
        phase1_df = phase1_df[phase1_df["Constituency Name"].isin(selected_constituencies)].copy()
        ps_df = ps_df[ps_df["Constituency Name"].isin(selected_constituencies)].copy()

    if phase1_df.empty or ps_df.empty:
        raise ValueError("No data found for the selected RO Office / Constituency.")

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

    if reserve_per_polling_station:
        missing_constituencies = set(ps_df["Constituency Name"].dropna()) - set(constituencies)
        if missing_constituencies:
            raise ValueError(
                "Phase 1 output is missing BU/CU allocations for polling-station "
                f"constituencies: {', '.join(sorted(map(str, missing_constituencies)))}."
            )
        # Validate *all* selected constituencies before making any assignment,
        # rather than silently generating an incomplete reserve allocation.
        for name in constituencies:
            const_units = phase1_df[phase1_df["Constituency Name"] == name]
            const_stations = ps_df[ps_df["Constituency Name"] == name]
            station_count = len(const_stations)
            bu_count = int((const_units["BU Unit ID"].notna() &
                            (const_units["BU Unit ID"].astype(str).str.strip() != "")).sum())
            cu_count = int((const_units["CU Unit ID"].notna() &
                            (const_units["CU Unit ID"].astype(str).str.strip() != "")).sum())
            needed = 2 * station_count
            if station_count == 0 or bu_count < needed or cu_count < needed:
                raise ValueError(
                    f"Constituency '{name}' needs at least {needed} BUs and {needed} CUs "
                    f"for {station_count} polling station(s): one working pair and one "
                    f"reserved pair per station. Found {bu_count} BUs and {cu_count} CUs."
                )

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

        # Optional designation: one extra randomly chosen *paired* reserve per PS.
        # The normal polling-station assignment above remains untouched.
        reserve_pair_count = n_ps if reserve_per_polling_station else 0
        for i in range(reserve_pair_count):
            reserve_bu = bu_records[n_assign + i]
            reserve_cu = cu_records[n_assign + i]
            row = {
                "SlNo": reserved_sl,
                "BU Unit ID": reserve_bu["BU Unit ID"],
                "CU Unit ID": reserve_cu["CU Unit ID"],
                "RO/ARO's Office": ro_office,
                "Constituency Name": const_name,
                "Polling Station No.": ps_list[i][0],
                "Polling Station Name": f"RESERVED for {ps_list[i][1]}",
            }
            if has_bu_box:
                row["BU Box No."] = reserve_bu.get("BU Box No.")
            if has_cu_box:
                row["CU Box No."] = reserve_cu.get("CU Box No.")
            reserved_rows.append(row)
            reserved_sl += 1

        for i in range(n_assign + reserve_pair_count, n_bu):
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

        for i in range(n_assign + reserve_pair_count, n_cu):
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

def get_phase2_pdf(
    assigned_df: pd.DataFrame,
    reserved_df: pd.DataFrame,
    randomization_date="",
    constituency_label=None,
    ro_office_label=None,
) -> bytes:
    buf = BytesIO()

    constituency, ro_office = _get_pdf_meta(
        assigned_df if assigned_df is not None and not assigned_df.empty else reserved_df
    )

    if constituency_label is not None:
        constituency = constituency_label
    if ro_office_label is not None:
        ro_office = ro_office_label

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
