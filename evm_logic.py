import pandas as pd
import random
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
CELL_FONT    = Font(name="Arial", size=10)
ALT_FILL     = PatternFill("solid", start_color="D6E4F0")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

def style_worksheet(ws, col_widths):
    for col_idx, cell in enumerate(ws[1], start=1):
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border    = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths[col_idx - 1]
    ws.row_dimensions[1].height = 28
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
        for cell in row:
            cell.font      = CELL_FONT
            cell.fill      = fill
            cell.border    = THIN_BORDER
            cell.alignment = CENTER_ALIGN
    ws.freeze_panes = "A2"

def to_excel_bytes(df, sheet_name, col_widths):
    """Return a styled Excel file as bytes (for Streamlit download)."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name=sheet_name)
    buf.seek(0)
    wb = load_workbook(buf)
    style_worksheet(wb.active, col_widths)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

def phase1_randomization(polling_df: pd.DataFrame, evm_df: pd.DataFrame):
    """
    Returns (result_df, summary_dict, warnings_list)
    """
    warnings = []

    required_ps  = {"Slno", "RO/ARO's Office", "Constituency Name",
                    "Polling Station No.", "Polling Station Name"}
    required_evm = {"EVMSlNo", "EVMCode"}

    missing_ps  = required_ps  - set(polling_df.columns)
    missing_evm = required_evm - set(evm_df.columns)
    if missing_ps:
        raise ValueError(f"Polling station file is missing columns: {missing_ps}")
    if missing_evm:
        raise ValueError(f"EVM file is missing columns: {missing_evm}")

    constituencies = (
        polling_df[["RO/ARO's Office", "Constituency Name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    num_const = len(constituencies)
    num_evms  = len(evm_df)

    if num_const == 0:
        raise ValueError("No constituencies found in the polling station file.")
    if num_evms == 0:
        raise ValueError("No EVMs found in the EVM file.")

    evm_indices   = list(range(num_evms))
    random.shuffle(evm_indices)
    const_indices = [i % num_const for i in range(num_evms)]
    random.shuffle(const_indices)

    records = []
    for pos, evm_idx in enumerate(evm_indices):
        const_row = constituencies.iloc[const_indices[pos]]
        records.append({
            "EVMSlNo":           evm_df.iloc[evm_idx]["EVMSlNo"],
            "EVMCode":           evm_df.iloc[evm_idx]["EVMCode"],
            "RO/ARO's Office":   const_row["RO/ARO's Office"],
            "Constituency Name": const_row["Constituency Name"],
        })

    result_df = pd.DataFrame(records)

    summary = {
        "total_evms":       num_evms,
        "total_const":      num_const,
        "distribution":     result_df.groupby("Constituency Name").size().to_dict(),
    }

    return result_df, summary, warnings


def phase2_randomization(phase1_df: pd.DataFrame, polling_df: pd.DataFrame):
    """
    Returns (result_df, summary_dict, warnings_list)
    """
    warnings = []

    required_evm = {"EVMSlNo", "EVMCode", "RO/ARO's Office", "Constituency Name"}
    required_ps  = {"Slno", "RO/ARO's Office", "Constituency Name",
                    "Polling Station No.", "Polling Station Name"}

    missing_evm = required_evm - set(phase1_df.columns)
    missing_ps  = required_ps  - set(polling_df.columns)
    if missing_evm:
        raise ValueError(f"Phase 1 output is missing columns: {missing_evm}")
    if missing_ps:
        raise ValueError(f"Polling station file is missing columns: {missing_ps}")

    all_records = []

    for constituency, evm_group in phase1_df.groupby("Constituency Name"):
        ps_group = polling_df[polling_df["Constituency Name"] == constituency].copy()

        if ps_group.empty:
            warnings.append(f"No polling stations found for '{constituency}'. "
                            "Polling station fields left blank.")
            for _, evm_row in evm_group.iterrows():
                all_records.append({
                    "EVMSlNo":              evm_row["EVMSlNo"],
                    "EVMCode":              evm_row["EVMCode"],
                    "RO/ARO's Office":      evm_row["RO/ARO's Office"],
                    "Constituency Name":    evm_row["Constituency Name"],
                    "Polling Station No.":  "",
                    "Polling Station Name": "",
                })
            continue

        ps_shuffled = ps_group.sample(frac=1).reset_index(drop=True)
        evm_list    = evm_group.reset_index(drop=True)
        num_ps      = len(ps_shuffled)
        num_evms    = len(evm_list)

        if num_evms > num_ps:
            warnings.append(
                f"'{constituency}': {num_evms} EVMs but only {num_ps} polling stations. "
                "Some stations assigned to multiple EVMs."
            )

        for i, evm_row in evm_list.iterrows():
            ps_row = ps_shuffled.iloc[i % num_ps]
            all_records.append({
                "EVMSlNo":              evm_row["EVMSlNo"],
                "EVMCode":              evm_row["EVMCode"],
                "RO/ARO's Office":      evm_row["RO/ARO's Office"],
                "Constituency Name":    evm_row["Constituency Name"],
                "Polling Station No.":  ps_row["Polling Station No."],
                "Polling Station Name": ps_row["Polling Station Name"],
            })

    result_df = pd.DataFrame(all_records)
    summary = {
        "total_evms":    len(result_df),
        "distribution":  result_df.groupby("Constituency Name")[["Constituency Name"]].size().to_dict(),
    }

    return result_df, summary, warnings
