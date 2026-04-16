
import io
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Asset Register Automation v3", layout="wide")

APP_VERSION = "v3.0.0 - 2026-04-16"

PREFERRED_SHEETS = ["FY-25-26-New ", "FY-25-26", "FY-24-25-New", "FY-24-25"]

COLUMN_ALIASES = {
    "asset_id": ["Assets ID", "ASTNO", "Asset ID"],
    "description": ["Assets Discreption", "DESC", "Description"],
    "asset_class": ["Assets Class", "CATEGORY", "Asset Class"],
    "nature": ["Nature of Assets", "COSTCENT", "Nature"],
    "location": ["Assets Location", "LOCATION", "Location"],
    "purchase_date": ["Purchase Date", "ACQDATE", "Acquisition Date"],
    "gross_block_closing": ["Gross Block Closing Value", "BKVALUE", "Purchase cost", "Gross Block"],
    "salvage_value": ["Salvage Value", "BKSALVAL", "Salvage"],
    "method": ["Depreciation Method", "BKMETHOD", "Method"],
    "life": ["Assets Life", "BKLIFE", "Life"],
    "opening_gross": ["Opening Gross Block"],
    "addition_during_year": ["Addition During the Year"],
    "disposed_during_year": ["Disposed During the Year", "Sales During the Year"],
    "closing_gross": ["Closing Gross Block"],
    "opening_acc_dep": ["Opening Accumlated Depreciation", "Opening Accumulated Depreciation"],
    "status": ["Disposed /Inuse", "Status of Assets", "Status"],
    "disposal_date": ["Disposal Date"],
    "depr_rate": ["Depreciation Rate", "BKDPRATE", "Rate"],
    "asset_qty": ["Assets Qty", "UNITS", "Qty"],
    "serial_no": ["Serial Number", "SERNO"],
    "vendor": ["VENDOR", "Vendor"],
    "invoice_no": ["Invoices Number", "APINVCNO", "Invoice Number"],
    "aqucode": ["AQUCODE"],
}

FY_START = pd.Timestamp("2026-04-01")
FY_END = pd.Timestamp("2027-03-31")


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for std_col, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in df.columns:
                rename_map[alias] = std_col
                break
    out = df.rename(columns=rename_map).copy()
    return out


def normalize_dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def pick_best_sheet(xls: pd.ExcelFile) -> str:
    for s in PREFERRED_SHEETS:
        if s in xls.sheet_names:
            return s
    fy_sheets = [s for s in xls.sheet_names if "FY-" in s]
    return fy_sheets[-1] if fy_sheets else xls.sheet_names[0]


def load_register(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    selected_sheet = pick_best_sheet(xls)
    raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=2)
    raw = raw.dropna(how="all")
    raw = raw.loc[:, ~raw.columns.astype(str).str.contains("^Unnamed")]
    raw = standardize_columns(raw)

    for col in ["purchase_date", "disposal_date"]:
        if col in raw.columns:
            raw[col] = normalize_dates(raw[col])

    for col in ["gross_block_closing", "salvage_value", "life", "opening_acc_dep", "depr_rate",
                "opening_gross", "addition_during_year", "disposed_during_year", "closing_gross"]:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce")

    if "status" in raw.columns:
        raw["status"] = raw["status"].astype(str).str.strip().replace({"nan": ""})

    raw = raw[raw.get("asset_id", pd.Series(index=raw.index, dtype=str)).notna()]
    raw = raw[raw["asset_id"].astype(str).str.strip() != ""]
    return raw, selected_sheet, xls.sheet_names


def default_life_map(df: pd.DataFrame):
    tmp = df[["asset_class", "life"]].copy() if {"asset_class", "life"}.issubset(df.columns) else pd.DataFrame(columns=["asset_class", "life"])
    tmp = tmp.dropna()
    if tmp.empty:
        return {}
    med = tmp.groupby("asset_class", dropna=True)["life"].median().to_dict()
    return med


def annual_slm_depr(cost, salvage, life_years):
    if pd.isna(cost) or pd.isna(life_years) or life_years in (0, None):
        return 0.0
    salvage = 0.0 if pd.isna(salvage) else salvage
    return max((cost - salvage) / life_years, 0.0)


def annual_wdv_rate(row):
    if not pd.isna(row.get("depr_rate")) and row.get("depr_rate") > 0:
        return float(row["depr_rate"]) / 100.0
    life = row.get("life")
    salvage = row.get("salvage_value")
    cost = row.get("gross_block_closing")
    if pd.isna(life) or life <= 0 or pd.isna(cost) or cost <= 0:
        return 0.0
    salvage_pct = 0 if pd.isna(salvage) else salvage / cost
    salvage_pct = min(max(salvage_pct, 0), 0.99)
    return 1 - (salvage_pct ** (1 / life))


def days_in_service_for_fy(purchase_date, disposal_date):
    start = max(FY_START, purchase_date) if pd.notna(purchase_date) else FY_START
    end = min(FY_END, disposal_date) if pd.notna(disposal_date) else FY_END
    if end < start:
        return 0
    return int((end - start).days) + 1


def calculate_fy_2026_27(df: pd.DataFrame, class_life_overrides: dict):
    df = df.copy()

    if "life" not in df.columns:
        df["life"] = np.nan
    if "salvage_value" not in df.columns:
        df["salvage_value"] = 0.0
    if "method" not in df.columns:
        df["method"] = "SLM"
    if "status" not in df.columns:
        df["status"] = "In use"

    # use closing gross as opening FY26-27 cost base
    df["opening_cost_fy26_27"] = np.where(
        df.get("closing_gross", pd.Series(index=df.index)).notna(),
        df.get("closing_gross"),
        df.get("gross_block_closing")
    )

    # fill life using class map
    if "asset_class" in df.columns:
        df["life"] = df.apply(
            lambda r: class_life_overrides.get(r.get("asset_class"), r.get("life")) if pd.isna(r.get("life")) or r.get("life") == 0 else r.get("life"),
            axis=1
        )

    df["purchase_date"] = normalize_dates(df.get("purchase_date"))
    df["disposal_date"] = normalize_dates(df.get("disposal_date"))
    df["status_norm"] = df.get("status", pd.Series([""] * len(df), index=df.index)).fillna("").astype(str).str.lower().str.strip()

    def _effective_disposal_date(row):
        status_val = row.get("status_norm", "")
        status_text = status_val if isinstance(status_val, str) else str(status_val or "")
        return row.get("disposal_date") if "disposed" in status_text else pd.NaT

    df["days_in_service_fy26_27"] = df.apply(
        lambda r: days_in_service_for_fy(r.get("purchase_date"), _effective_disposal_date(r)),
        axis=1,
    )

    def depr(row):
        cost = row.get("opening_cost_fy26_27")
        opening_acc = row.get("opening_acc_dep", 0.0)
        nbv_open = max((cost or 0) - (opening_acc or 0), 0.0)
        salvage = row.get("salvage_value", 0.0)
        method = str(row.get("method", "SLM")).upper().strip()
        days = row.get("days_in_service_fy26_27", 0)
        proportion = max(min(days / 365.0, 1), 0)

        if method == "WDV":
            rate = annual_wdv_rate(row)
            charge = nbv_open * rate * proportion
            max_allowed = max(nbv_open - salvage, 0.0)
            return min(charge, max_allowed)
        annual = annual_slm_depr(cost or 0, salvage or 0, row.get("life"))
        charge = annual * proportion
        max_allowed = max(nbv_open - salvage, 0.0)
        return min(charge, max_allowed)

    df["depreciation_fy26_27"] = df.apply(depr, axis=1).round(2)
    df["closing_acc_dep_fy26_27"] = (df.get("opening_acc_dep", 0).fillna(0) + df["depreciation_fy26_27"]).round(2)
    df["closing_nbv_31_03_2027"] = (df["opening_cost_fy26_27"].fillna(0) - df["closing_acc_dep_fy26_27"]).clip(lower=0).round(2)
    return df


def summary_table(df: pd.DataFrame):
    group_col = "asset_class" if "asset_class" in df.columns else None
    if group_col is None:
        return pd.DataFrame([{
            "Opening Cost FY26-27": df["opening_cost_fy26_27"].sum(),
            "Opening Acc Dep": df.get("opening_acc_dep", pd.Series(dtype=float)).fillna(0).sum(),
            "Depreciation FY26-27": df["depreciation_fy26_27"].sum(),
            "Closing NBV 31-03-2027": df["closing_nbv_31_03_2027"].sum(),
        }])

    out = (
        df.groupby(group_col, dropna=False)
        .agg(
            assets=("asset_id", "count"),
            opening_cost_fy26_27=("opening_cost_fy26_27", "sum"),
            opening_acc_dep=("opening_acc_dep", "sum"),
            depreciation_fy26_27=("depreciation_fy26_27", "sum"),
            closing_nbv_31_03_2027=("closing_nbv_31_03_2027", "sum"),
        )
        .reset_index()
        .sort_values("asset_class", na_position="last")
    )
    return out


def to_excel_bytes(detail_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="FY 2026-27 Computation")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
    bio.seek(0)
    return bio.read()


st.title("Asset Register Automation Portal")
st.caption("Upload your asset register, compute FY 2026–27 depreciation, review the schedule, and download the output workbook.")

with st.sidebar:
    st.header("Configuration")
    st.write("Financial Year: **2026–27**")
    fy_note = st.info("This app assumes opening balances are taken from the uploaded register and computes depreciation for 01-Apr-2026 to 31-Mar-2027.")

uploaded = st.file_uploader("Upload Asset Register Excel", type=["xlsx", "xlsm"])

if not uploaded:
    st.markdown(
        """
        ### What this software does
        - Reads your current asset register workbook
        - Picks the latest FY sheet automatically
        - Rolls values forward into FY 2026–27
        - Computes depreciation class-wise and asset-wise
        - Handles disposed assets with part-year depreciation
        - Exports the updated asset register
        """
    )
    st.stop()

try:
    df, selected_sheet, all_sheets = load_register(uploaded)
except Exception as e:
    st.error(f"Could not read the workbook: {e}")
    st.stop()

st.success(f"Loaded sheet: **{selected_sheet}**")
with st.expander("Detected workbook sheets"):
    st.write(all_sheets)

life_map = default_life_map(df)

st.subheader("Life assumptions")
if life_map:
    editable = pd.DataFrame({
        "asset_class": list(life_map.keys()),
        "life_years": list(life_map.values())
    }).sort_values("asset_class")
else:
    editable = pd.DataFrame({"asset_class": [], "life_years": []})

edited = st.data_editor(editable, num_rows="dynamic", use_container_width=True, key="life_editor")
override_map = dict(zip(edited["asset_class"], pd.to_numeric(edited["life_years"], errors="coerce")))

result = calculate_fy_2026_27(df, override_map)
summary = summary_table(result)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Assets", f"{len(result):,}")
k2.metric("Opening Cost", f"{result['opening_cost_fy26_27'].fillna(0).sum():,.2f}")
k3.metric("Depreciation FY26-27", f"{result['depreciation_fy26_27'].sum():,.2f}")
k4.metric("Closing NBV", f"{result['closing_nbv_31_03_2027'].sum():,.2f}")

tab1, tab2, tab3 = st.tabs(["Computation", "Summary", "Checks"])

with tab1:
    show_cols = [c for c in [
        "asset_id", "description", "asset_class", "location", "purchase_date",
        "method", "life", "opening_cost_fy26_27", "opening_acc_dep",
        "depreciation_fy26_27", "closing_acc_dep_fy26_27", "closing_nbv_31_03_2027",
        "status", "disposal_date", "days_in_service_fy26_27"
    ] if c in result.columns]
    st.dataframe(result[show_cols], use_container_width=True, height=500)

with tab2:
    st.dataframe(summary, use_container_width=True, height=450)

with tab3:
    checks = pd.DataFrame({
        "Check": [
            "Missing purchase dates",
            "Missing useful life",
            "Disposed assets without disposal date",
            "Negative closing NBV",
        ],
        "Count": [
            int(result.get("purchase_date", pd.Series(dtype='datetime64[ns]')).isna().sum()),
            int(result.get("life", pd.Series(dtype=float)).isna().sum()),
            int(((result.get("status", pd.Series(dtype=str)).astype(str).str.contains("disposed", case=False, na=False)) &
                 (result.get("disposal_date", pd.Series(dtype='datetime64[ns]')).isna())).sum()),
            int((result["closing_nbv_31_03_2027"] < 0).sum()),
        ]
    })
    st.dataframe(checks, use_container_width=True)

excel_bytes = to_excel_bytes(result, summary)
st.download_button(
    "Download FY 2026-27 Output Workbook",
    data=excel_bytes,
    file_name="Asset_Register_FY2026_27_Output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# Deployment note: if Streamlit traceback still shows the old inline status check, the app is still running an older GitHub commit.
