
import io
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st


APP_VERSION = "v1.0.0 - 2026-04-17"
FY_START = pd.Timestamp("2026-04-01")
FY_END = pd.Timestamp("2027-03-31")


def normalize_col(x):
    if x is None:
        return ""
    return (
        str(x)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("  ", " ")
    )


def first_present(df, candidates, default=None):
    norm = {normalize_col(c): c for c in df.columns}
    for c in candidates:
        if normalize_col(c) in norm:
            return norm[normalize_col(c)]
    return default


def coerce_numeric(series, fill=0.0):
    return pd.to_numeric(series, errors="coerce").fillna(fill)


def coerce_date(series):
    return pd.to_datetime(series, errors="coerce")


def month_ends_for_fy(start=FY_START):
    months = pd.date_range(start, FY_END, freq="M")
    return list(months)


def month_start(ts):
    return pd.Timestamp(year=ts.year, month=ts.month, day=1)


def days_overlap(start_a, end_a, start_b, end_b):
    if pd.isna(start_a) or pd.isna(end_a):
        return 0
    s = max(pd.Timestamp(start_a), pd.Timestamp(start_b))
    e = min(pd.Timestamp(end_a), pd.Timestamp(end_b))
    if s > e:
        return 0
    return (e - s).days + 1


def safe_status_text(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return str(x).strip().lower()


def ensure_series(df, colname, fill=0.0):
    if colname in df.columns:
        return pd.to_numeric(df[colname], errors="coerce").fillna(fill)
    return pd.Series([fill] * len(df), index=df.index, dtype="float64")


def choose_header_row(preview_df):
    best_row = 0
    best_score = -1
    expected = {"assets id", "purchase date", "gross block closing value", "assets class", "status"}
    for idx in range(min(15, len(preview_df))):
        vals = [normalize_col(v) for v in preview_df.iloc[idx].tolist()]
        score = len(expected.intersection(set(vals)))
        if score > best_score:
            best_score = score
            best_row = idx
    return best_row


def read_workbook(uploaded_file, chosen_sheet=None):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = chosen_sheet or xls.sheet_names[0]
    preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, nrows=15)
    header_row = choose_header_row(preview)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    return df, sheet_name, header_row


def map_columns(df):
    cols = {}
    cols["asset_id"] = first_present(df, ["Assets ID", "Asset ID", "ASTNO"])
    cols["description"] = first_present(df, ["Assets Discreption", "Asset Description", "DESC"])
    cols["serial_number"] = first_present(df, ["Serial Number", "SERNO"])
    cols["aqucode"] = first_present(df, ["AQUCODE"])
    cols["invoice_no"] = first_present(df, ["Invoices Number", "Invoice Number", "APINVCNO"])
    cols["vendor"] = first_present(df, ["VENDOR"])
    cols["asset_class"] = first_present(df, ["Assets Class", "Asset Class", "CATEGORY"])
    cols["nature"] = first_present(df, ["Nature of Assets", "COSTCENT"])
    cols["location"] = first_present(df, ["Assets Location", "Asset Location", "LOCATION"])
    cols["purchase_date"] = first_present(df, ["Purchase Date", "ACQDATE"])
    cols["gross_closing"] = first_present(df, ["Gross Block Closing Value", "BKVALUE"])
    cols["salvage"] = first_present(df, ["Salvage Value", "BKSALVAL"])
    cols["dep_method"] = first_present(df, ["Depreciation Method"])
    cols["opening_asset_date"] = first_present(df, ["Opening Assets Dates"])
    cols["asset_method"] = first_present(df, ["Assets Method"])
    cols["qty"] = first_present(df, ["Assets Qty"])
    cols["dep_rate"] = first_present(df, ["Depreciation Rate"])
    cols["asset_life"] = first_present(df, ["Assets Life"])
    cols["opening_gross"] = first_present(df, ["Opening Gross Block as on 1st Apr-25"])
    cols["addition"] = first_present(df, ["Addition During the Year"])
    cols["deletion"] = first_present(df, ["Deletion During the Year"])
    cols["closing_gross"] = first_present(df, ["Closing Gross Block"])
    cols["expiry_date"] = first_present(df, ["Exipiry Date", "Expiry Date"])
    cols["life_days"] = first_present(df, ["Life of Asstess in Days", "Life of Assets in Days"])
    cols["depreciable_value"] = first_present(df, ["Depreciable Value"])
    cols["dep_per_day"] = first_present(df, ["Depreciation per Day"])
    cols["opening_acc_dep"] = first_present(df, ["Opening Accumlated Dep as on 1st Apr-25"])
    cols["total_depreciation"] = first_present(df, ["Total Depreciation"])
    cols["dep_on_disposal"] = first_present(df, ["Depreciation on Disposal"])
    cols["closing_acc_dep"] = first_present(df, ["Closing  Accummlated Depreciation", "Closing Accummlated Depreciation"])
    cols["opening_net"] = first_present(df, ["Opening Net Block"])
    cols["closing_net"] = first_present(df, ["Closing Net Block"])
    cols["status"] = first_present(df, ["Status"])
    cols["sub_classification"] = first_present(df, ["Sub- Classification"])
    cols["remark"] = first_present(df, ["Remark"])
    return cols


def prepare_base_dataframe(df, cols):
    out = pd.DataFrame(index=df.index)
    for key, col in cols.items():
        if col and col in df.columns:
            out[key] = df[col]
    out["asset_id"] = out.get("asset_id", pd.Series(index=df.index, dtype="object")).astype(str).str.strip()
    out["description"] = out.get("description", "")
    out["asset_class"] = out.get("asset_class", "")
    out["nature"] = out.get("nature", "")
    out["location"] = out.get("location", "")
    out["vendor"] = out.get("vendor", "")
    out["purchase_date"] = coerce_date(out.get("purchase_date", pd.NaT))
    out["status"] = out.get("status", "").apply(safe_status_text) if "status" in out.columns else ""
    out["gross_closing"] = coerce_numeric(out.get("gross_closing", 0))
    out["closing_gross"] = coerce_numeric(out.get("closing_gross", out["gross_closing"]))
    out["salvage"] = coerce_numeric(out.get("salvage", 0))
    out["asset_life"] = coerce_numeric(out.get("asset_life", 0))
    out["life_days"] = coerce_numeric(out.get("life_days", 0))
    out["opening_acc_dep"] = coerce_numeric(out.get("closing_acc_dep", out.get("opening_acc_dep", 0)))
    out["opening_net"] = coerce_numeric(out.get("closing_net", out.get("opening_net", 0)))
    out["sub_classification"] = out.get("sub_classification", "")
    out["remark"] = out.get("remark", "")
    return out


def build_fy_2026_27(df, disposals_df=None):
    cols = map_columns(df)
    asset_df = prepare_base_dataframe(df, cols)
    asset_df = asset_df[asset_df["asset_id"].ne("")].copy()

    if disposals_df is not None and not disposals_df.empty:
        disp = disposals_df.copy()
        disp.columns = [c.strip() for c in disp.columns]
        if "asset_id" in disp.columns:
            disp["asset_id"] = disp["asset_id"].astype(str).str.strip()
            disp["disposal_date"] = pd.to_datetime(disp.get("disposal_date"), errors="coerce")
            disp["status_override"] = disp.get("status_override", "Disposed").fillna("Disposed")
            asset_df = asset_df.merge(
                disp[["asset_id", "disposal_date", "status_override"]],
                how="left",
                on="asset_id",
            )
        else:
            asset_df["disposal_date"] = pd.NaT
            asset_df["status_override"] = None
    else:
        asset_df["disposal_date"] = pd.NaT
        asset_df["status_override"] = None

    asset_df["fy26_opening_gross"] = asset_df["closing_gross"].where(asset_df["closing_gross"].gt(0), asset_df["gross_closing"])
    asset_df["fy26_addition"] = 0.0
    asset_df["fy26_deletion"] = 0.0
    asset_df["fy26_closing_gross"] = asset_df["fy26_opening_gross"] + asset_df["fy26_addition"] - asset_df["fy26_deletion"]

    # derive asset life in days
    derived_life_days = np.where(
        asset_df["life_days"].gt(0),
        asset_df["life_days"],
        np.where(asset_df["asset_life"].gt(0), asset_df["asset_life"] * 365, np.nan),
    )
    asset_df["life_days_fy26"] = pd.Series(derived_life_days, index=asset_df.index).fillna(0).round(0)

    # depreciation base
    asset_df["salvage"] = asset_df["salvage"].clip(lower=0)
    asset_df["depreciable_value_fy26"] = (asset_df["fy26_closing_gross"] - asset_df["salvage"]).clip(lower=0)
    asset_df["dep_per_day_fy26"] = np.where(
        asset_df["life_days_fy26"] > 0,
        asset_df["depreciable_value_fy26"] / asset_df["life_days_fy26"],
        0.0,
    )

    asset_df["current_status"] = asset_df["status_override"].fillna(asset_df["status"]).apply(safe_status_text)
    asset_df["disposal_date"] = pd.to_datetime(asset_df["disposal_date"], errors="coerce")

    month_ends = month_ends_for_fy()
    month_cols = []
    for me in month_ends:
        ms = month_start(me)
        col = me.strftime("%b-%y")
        month_cols.append(col)
        days = []
        for _, r in asset_df.iterrows():
            service_start = max(FY_START, r["purchase_date"]) if pd.notna(r["purchase_date"]) else FY_START
            service_end = FY_END
            if ("disposed" in safe_status_text(r["current_status"])) and pd.notna(r["disposal_date"]):
                service_end = min(FY_END, r["disposal_date"])
            overlap = days_overlap(service_start, service_end, ms, me)
            remaining_dep = max(r["depreciable_value_fy26"] - r["opening_acc_dep"], 0)
            amount = min(overlap * r["dep_per_day_fy26"], remaining_dep)
            days.append(round(float(amount), 2))
        asset_df[col] = days

    asset_df["total_depreciation_fy26_27"] = asset_df[month_cols].sum(axis=1).round(2)
    asset_df["depreciation_on_disposal_fy26_27"] = 0.0
    asset_df["closing_acc_dep_fy26_27"] = (asset_df["opening_acc_dep"] + asset_df["total_depreciation_fy26_27"]).round(2)
    asset_df["opening_net_block_fy26_27"] = (asset_df["fy26_opening_gross"] - asset_df["opening_acc_dep"]).round(2)
    asset_df["closing_net_block_fy26_27"] = (asset_df["fy26_closing_gross"] - asset_df["closing_acc_dep_fy26_27"]).round(2)
    asset_df["closing_net_block_fy26_27"] = asset_df["closing_net_block_fy26_27"].clip(lower=asset_df["salvage"])
    asset_df["expiry_date_fy26"] = np.where(
        (asset_df["purchase_date"].notna()) & (asset_df["asset_life"].gt(0)),
        asset_df["purchase_date"] + pd.to_timedelta((asset_df["asset_life"] * 365).round(0), unit="D"),
        pd.NaT,
    )

    result = pd.DataFrame({
        "SR No.": range(1, len(asset_df) + 1),
        "Assets ID": asset_df["asset_id"],
        "Assets Description": asset_df["description"],
        "Asset Class": asset_df["asset_class"],
        "Nature of Assets": asset_df["nature"],
        "Assets Location": asset_df["location"],
        "Vendor": asset_df["vendor"],
        "Purchase Date": asset_df["purchase_date"],
        "Opening Gross Block as on 1st Apr-26": asset_df["fy26_opening_gross"],
        "Addition During FY 26-27": asset_df["fy26_addition"],
        "Deletion During FY 26-27": asset_df["fy26_deletion"],
        "Closing Gross Block": asset_df["fy26_closing_gross"],
        "Salvage Value": asset_df["salvage"],
        "Assets Life (Years)": asset_df["asset_life"],
        "Life of Assets in Days": asset_df["life_days_fy26"],
        "Depreciable Value": asset_df["depreciable_value_fy26"],
        "Depreciation per Day": asset_df["dep_per_day_fy26"],
        "Opening Accumulated Dep as on 1st Apr-26": asset_df["opening_acc_dep"],
    })
    for col in month_cols:
        result[col] = asset_df[col]
    result["Total Depreciation"] = asset_df["total_depreciation_fy26_27"]
    result["Depreciation on Disposal"] = asset_df["depreciation_on_disposal_fy26_27"]
    result["Closing Accumulated Depreciation"] = asset_df["closing_acc_dep_fy26_27"]
    result["Opening Net Block"] = asset_df["opening_net_block_fy26_27"]
    result["Closing Net Block"] = asset_df["closing_net_block_fy26_27"]
    result["Status"] = asset_df["current_status"].str.title()
    result["Disposal Date"] = asset_df["disposal_date"]
    result["Sub-Classification"] = asset_df["sub_classification"]
    result["Remark"] = asset_df["remark"]

    summary = (
        result.groupby("Asset Class", dropna=False)
        .agg(
            asset_count=("Assets ID", "count"),
            opening_gross=("Opening Gross Block as on 1st Apr-26", "sum"),
            total_depreciation=("Total Depreciation", "sum"),
            closing_gross=("Closing Gross Block", "sum"),
            closing_net_block=("Closing Net Block", "sum"),
        )
        .reset_index()
        .sort_values("Asset Class", na_position="last")
    )
    summary.columns = [
        "Asset Class",
        "Asset Count",
        "Opening Gross Block",
        "Total Depreciation FY26-27",
        "Closing Gross Block",
        "Closing Net Block",
    ]
    return result, summary


def to_excel_bytes(register_df, summary_df, source_df, source_sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        register_df.to_excel(writer, index=False, sheet_name="FY-26-27-Auto")
        summary_df.to_excel(writer, index=False, sheet_name="Summary FY26-27")
        source_df.to_excel(writer, index=False, sheet_name="Source Copy")
        info = pd.DataFrame({
            "Item": [
                "App version",
                "Source sheet",
                "Financial year",
                "Note 1",
                "Note 2",
            ],
            "Value": [
                APP_VERSION,
                source_sheet_name,
                "FY 2026-27",
                "Opening balances are rolled from the selected source sheet.",
                "Additions for FY26-27 can be incorporated by editing the exported workbook.",
            ]
        })
        info.to_excel(writer, index=False, sheet_name="Utility Notes")

        wb = writer.book
        for name in ["FY-26-27-Auto", "Summary FY26-27", "Source Copy", "Utility Notes"]:
            ws = wb[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                length = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(length + 2, 12), 28)
    return output.getvalue()


st.set_page_config(page_title="Asset Register App", layout="wide")
st.title("Asset Register Automation App")
st.caption(f"App version: {APP_VERSION}")

st.markdown(
    """
Upload the existing asset register workbook, select the source sheet, and generate a FY 2026-27 register,
summary, and export file.
"""
)

uploaded_file = st.file_uploader("Upload asset register workbook (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    chosen_sheet = st.selectbox("Select asset register sheet", xls.sheet_names, index=min(1, len(xls.sheet_names)-1))
    uploaded_file.seek(0)
    raw_df, source_sheet_name, header_row = read_workbook(uploaded_file, chosen_sheet)

    st.success(f"Loaded sheet: {source_sheet_name} | detected header row: {header_row + 1}")

    with st.expander("Preview source data"):
        st.dataframe(raw_df.head(20), use_container_width=True)

    st.subheader("Optional disposals for FY 2026-27")
    disp_template = pd.DataFrame({
        "asset_id": [""],
        "disposal_date": [pd.NaT],
        "status_override": ["Disposed"],
    })
    disposals = st.data_editor(disp_template, num_rows="dynamic", use_container_width=True)

    if st.button("Generate FY 2026-27 Register", type="primary"):
        try:
            register_df, summary_df = build_fy_2026_27(raw_df, disposals)
            st.subheader("FY 2026-27 Register Preview")
            st.dataframe(register_df.head(50), use_container_width=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("Assets", int(register_df["Assets ID"].count()))
            c2.metric("Opening Gross Block", f"{register_df['Opening Gross Block as on 1st Apr-26'].sum():,.2f}")
            c3.metric("Total Depreciation", f"{register_df['Total Depreciation'].sum():,.2f}")

            st.subheader("Summary by Asset Class")
            st.dataframe(summary_df, use_container_width=True)

            excel_bytes = to_excel_bytes(register_df, summary_df, raw_df, source_sheet_name)
            st.download_button(
                "Download FY 2026-27 Excel Output",
                data=excel_bytes,
                file_name="Asset_Register_FY2026-27_App_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Processing error: {e}")
            st.exception(e)
else:
    st.info("Upload the asset register workbook to begin.")
