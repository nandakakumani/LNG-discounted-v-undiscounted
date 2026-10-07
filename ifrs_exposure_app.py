from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "LNG IFRS Exposure Builder"
DEFAULT_YEARS = list(range(2026, 2032))
REQUIRED_COLUMNS = {"PnL Year", "Delta", "Final Index", "Group", "ViewGroup"}

# Delta is assumed to be in MMBtu for gas/HH and barrels for oil.
# Dividing by 1,000,000 gives TBTU for gas/HH and MM bbl for oil.
DEFAULT_DIVISORS = {
    "NET GAS (TBTU)": 1_000_000.0,
    "NET OIL (MM BBL)": 1_000_000.0,
    "HH (TBTU)": 1_000_000.0,
}

OIL_INDEXES = {
    "BRENT", "BRENT BULLET", "BRENT FUTURES", "BRENT SWAPS",
    "DATED BRENT", "DUBAI", "JCC", "JCC SWAPS", "WTI"
}
GAS_INDEXES = {
    "TTF", "NBP", "PEG", "PEG - DNK", "PVB", "THE", "TVB",
    "ZTP", "ZTP - DNK", "JKM"
}


def _clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).upper()


def parse_pnl_year(series: pd.Series) -> pd.Series:
    """Return nullable integer years from dates, strings, Excel serials, or direct years."""
    out = pd.Series(pd.NA, index=series.index, dtype="Int64")
    numeric = pd.to_numeric(series, errors="coerce")

    direct_year = numeric.between(1900, 2200, inclusive="both")
    out.loc[direct_year] = numeric.loc[direct_year].round().astype("Int64")

    serial = numeric.between(20000, 100000, inclusive="both") & ~direct_year
    if serial.any():
        serial_dates = pd.to_datetime(
            numeric.loc[serial], unit="D", origin="1899-12-30", errors="coerce"
        )
        out.loc[serial] = serial_dates.dt.year.astype("Int64")

    remaining = out.isna()
    if remaining.any():
        parsed = pd.to_datetime(series.loc[remaining], errors="coerce", dayfirst=False)
        out.loc[remaining] = parsed.dt.year.astype("Int64")
    return out


def classify_exposure(row: pd.Series) -> str:
    group = _clean_text(row.get("Group"))
    index = _clean_text(row.get("Final Index"))

    # HH is deliberately kept separate from general gas.
    if index == "HH" or index.startswith("HENRY HUB"):
        return "HH (TBTU)"

    # Exclude emissions and all freight, including LNG freight.
    if "EUA" in group or "EUA" in index or "FREIGHT" in group or "FREIGHT" in index:
        return "EXCLUDED"

    if group == "OIL" or index in OIL_INDEXES or any(token in index for token in ("BRENT", "JCC", "DUBAI")):
        return "NET OIL (MM BBL)"

    gas_group = (
        group == "NATURAL GAS & LNG"
        or group == "UNSOLD HOUSE CURVES"
        or "NATURAL GAS" in group
        or "LNG" in group
        or "GAS" in group
    )
    gas_index = index in GAS_INDEXES or "LNG" in index
    if gas_group or gas_index:
        return "NET GAS (TBTU)"

    return "UNCLASSIFIED"


def load_raw_data(file_obj) -> pd.DataFrame:
    xls = pd.ExcelFile(file_obj, engine="openpyxl")
    matching = [s for s in xls.sheet_names if s.strip().casefold() == "raw data"]
    if not matching:
        raise ValueError('The workbook must contain a sheet named "Raw Data".')

    df = pd.read_excel(xls, sheet_name=matching[0])
    df.columns = [str(c).strip() for c in df.columns]
    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError("Raw Data is missing required columns: " + ", ".join(missing))

    df = df.copy()
    df["Delta"] = pd.to_numeric(df["Delta"], errors="coerce").fillna(0.0)
    df["IFRS Year"] = parse_pnl_year(df["PnL Year"])
    df["Exposure Category"] = df.apply(classify_exposure, axis=1)
    df["ViewGroup Clean"] = df["ViewGroup"].map(_clean_text)
    return df


def _view_rows(df: pd.DataFrame, view_name: str) -> pd.DataFrame:
    wanted = view_name.upper()
    # 'Both' records belong in both base-case views.
    return df[df["ViewGroup Clean"].isin([wanted, "BOTH"])]


def build_table(
    df: pd.DataFrame,
    view_name: str,
    years: list[int],
    divisors: dict[str, float],
) -> pd.DataFrame:
    categories = ["NET GAS (TBTU)", "NET OIL (MM BBL)", "HH (TBTU)"]
    view = _view_rows(df, view_name)
    rows = []
    for category in categories:
        row = {"LNG PORTFOLIO EXPOSURES": category}
        divisor = divisors[category]
        subset = view[view["Exposure Category"].eq(category)]
        for year in years:
            row[f"IFRS {year}"] = subset.loc[subset["IFRS Year"].eq(year), "Delta"].sum() / divisor
        row["NET"] = sum(row[f"IFRS {year}"] for year in years)
        rows.append(row)
    return pd.DataFrame(rows).set_index("LNG PORTFOLIO EXPOSURES")


def build_outputs(
    df: pd.DataFrame,
    years: list[int],
    divisors: dict[str, float],
    portfolios: list[str] | None = None,
    books: list[str] | None = None,
):
    filtered = df.copy()
    if portfolios is not None and "Portfolio" in filtered.columns:
        filtered = filtered[filtered["Portfolio"].astype(str).isin(portfolios)]
    if books is not None and "Book" in filtered.columns:
        filtered = filtered[filtered["Book"].astype(str).isin(books)]

    included = filtered[filtered["Exposure Category"].isin(DEFAULT_DIVISORS)].copy()
    undiscounted = build_table(included, "Undiscounted", years, divisors)
    discounted = build_table(included, "Discounted", years, divisors)
    impact = discounted - undiscounted

    detail_parts = []
    for view in ("Undiscounted", "Discounted"):
        d = _view_rows(included, view).copy()
        if d.empty:
            continue
        d["IFRS Year"] = d["IFRS Year"].astype("Int64")
        grouped = (
            d[d["IFRS Year"].isin(years)]
            .groupby(["Exposure Category", "Final Index", "IFRS Year"], dropna=False)["Delta"]
            .sum()
            .reset_index()
        )
        grouped["Exposure"] = grouped.apply(
            lambda r: r["Delta"] / divisors[r["Exposure Category"]], axis=1
        )
        grouped.insert(0, "View", view)
        detail_parts.append(grouped.drop(columns="Delta"))
    detail = pd.concat(detail_parts, ignore_index=True) if detail_parts else pd.DataFrame()

    audit = (
        filtered.groupby(["Exposure Category", "Group", "Final Index"], dropna=False)
        .agg(Rows=("Delta", "size"), Raw_Delta=("Delta", "sum"))
        .reset_index()
        .sort_values(["Exposure Category", "Group", "Final Index"], na_position="last")
    )
    return filtered, undiscounted, discounted, impact, detail, audit


def _write_summary_section(ws, start_row: int, title: str, table: pd.DataFrame, years: list[int], impact_refs=None):
    final_col = len(years) + 2
    dark_fill = PatternFill("solid", fgColor="D9EAF7")
    header_fill = PatternFill("solid", fgColor="EDEDED")
    thin_gray = Side(style="thin", color="808080")
    title_border = Border(top=thin_gray, bottom=thin_gray)

    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=final_col)
    c = ws.cell(start_row, 1, title)
    c.font = Font(bold=True, color="1F1F1F")
    c.alignment = Alignment(horizontal="center")
    c.fill = dark_fill
    c.border = title_border

    header_row = start_row + 2
    headers = ["LNG PORTFOLIO EXPOSURES"] + [f"IFRS {y}" for y in years] + ["NET"]
    for col, value in enumerate(headers, 1):
        cell = ws.cell(header_row, col, value)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=thin_gray)

    data_start = header_row + 1
    for r_offset, label in enumerate(table.index):
        row = data_start + r_offset
        ws.cell(row, 1, label).font = Font(bold=True)
        for y_offset, year in enumerate(years, 2):
            cell = ws.cell(row, y_offset)
            if impact_refs:
                disc_row, undisc_row = impact_refs[label]
                cell.value = f"={get_column_letter(y_offset)}{disc_row}-{get_column_letter(y_offset)}{undisc_row}"
            else:
                cell.value = float(table.loc[label, f"IFRS {year}"])
            cell.number_format = '0.00;[Red]-0.00;-'
        net_cell = ws.cell(row, final_col)
        net_cell.value = f"=SUM(B{row}:{get_column_letter(final_col-1)}{row})"
        net_cell.font = Font(bold=True)
        net_cell.number_format = '0.00;[Red]-0.00;-'
        net_cell.border = Border(left=thin_gray)
    return {label: data_start + i for i, label in enumerate(table.index)}, data_start + len(table.index) - 1


def make_excel(
    source_df: pd.DataFrame,
    undiscounted: pd.DataFrame,
    discounted: pd.DataFrame,
    impact: pd.DataFrame,
    detail: pd.DataFrame,
    audit: pd.DataFrame,
    years: list[int],
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    undisc_rows, end1 = _write_summary_section(
        ws, 1, "OVERALL PORTFOLIO BALANCE BASE CASE (UNDISCOUNTED)", undiscounted, years
    )
    disc_start = end1 + 3
    disc_rows, end2 = _write_summary_section(
        ws, disc_start, "OVERALL PORTFOLIO BALANCE BASE CASE (DISCOUNTED)", discounted, years
    )
    impact_start = end2 + 3
    refs = {k: (disc_rows[k], undisc_rows[k]) for k in undiscounted.index}
    _write_summary_section(
        ws, impact_start, "DISCOUNTING IMPACT ON HEDGES", impact, years, impact_refs=refs
    )

    ws.column_dimensions["A"].width = 27
    for col in range(2, len(years) + 3):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.freeze_panes = "B4"

    def add_dataframe_sheet(name: str, frame: pd.DataFrame):
        sh = wb.create_sheet(name)
        sh.sheet_view.showGridLines = False
        if frame.empty:
            sh["A1"] = "No rows"
            return
        export = frame.copy()
        for col_idx, col_name in enumerate(export.columns, 1):
            cell = sh.cell(1, col_idx, str(col_name))
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")
        for row_idx, values in enumerate(export.itertuples(index=False, name=None), 2):
            for col_idx, value in enumerate(values, 1):
                if pd.isna(value):
                    value = None
                elif isinstance(value, np.generic):
                    value = value.item()
                sh.cell(row_idx, col_idx, value)
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions
        for col_idx, col_name in enumerate(export.columns, 1):
            width = max(len(str(col_name)) + 2, 12)
            for value in export.iloc[:500, col_idx - 1]:
                width = max(width, min(len(str(value)) + 2, 35))
            sh.column_dimensions[get_column_letter(col_idx)].width = width
        for col_idx, col_name in enumerate(export.columns, 1):
            if str(col_name) in {"Exposure", "Raw_Delta", "Delta"}:
                for row_idx in range(2, sh.max_row + 1):
                    sh.cell(row_idx, col_idx).number_format = '0.00;[Red]-0.00;-'

    add_dataframe_sheet("Product Detail", detail)
    add_dataframe_sheet("Classification Audit", audit)

    raw_export = source_df.drop(columns=["ViewGroup Clean"], errors="ignore").copy()
    add_dataframe_sheet("Filtered Raw Data", raw_export)

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def style_for_streamlit(df: pd.DataFrame):
    return df.style.format("{:.2f}").map(lambda x: "color: red" if isinstance(x, (int, float)) and x < 0 else "")


def run_app():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption('Upload the exposure workbook. The app reads "Raw Data", uses PnL Year for IFRS classification, and exports a formatted Excel pack.')

    uploaded = st.file_uploader("Upload XLSX workbook", type=["xlsx"])
    if uploaded is None:
        st.info("Upload the LNG exposure workbook to begin.")
        return

    try:
        df = load_raw_data(uploaded)
    except Exception as exc:
        st.error(str(exc))
        return

    with st.sidebar:
        st.header("Controls")
        years = st.multiselect("IFRS years", options=list(range(2020, 2041)), default=DEFAULT_YEARS)
        years = sorted(years)

        portfolios = None
        if "Portfolio" in df.columns:
            portfolio_options = sorted(df["Portfolio"].dropna().astype(str).unique())
            portfolios = st.multiselect("Portfolios", portfolio_options, default=portfolio_options)

        books = None
        if "Book" in df.columns:
            book_options = sorted(df["Book"].dropna().astype(str).unique())
            books = st.multiselect("Books", book_options, default=book_options)

        with st.expander("Unit divisors", expanded=False):
            st.caption("Defaults assume gas/HH Delta is MMBtu and oil Delta is barrels.")
            gas_divisor = st.number_input("Gas divisor", min_value=1.0, value=1_000_000.0, step=100_000.0)
            oil_divisor = st.number_input("Oil divisor", min_value=1.0, value=1_000_000.0, step=100_000.0)
            hh_divisor = st.number_input("HH divisor", min_value=1.0, value=1_000_000.0, step=100_000.0)

    if not years:
        st.warning("Select at least one IFRS year.")
        return

    divisors = {
        "NET GAS (TBTU)": gas_divisor,
        "NET OIL (MM BBL)": oil_divisor,
        "HH (TBTU)": hh_divisor,
    }

    filtered, undiscounted, discounted, impact, detail, audit = build_outputs(
        df, years, divisors, portfolios=portfolios, books=books
    )

    unclassified = filtered[filtered["Exposure Category"].eq("UNCLASSIFIED")]
    if not unclassified.empty:
        st.warning(f"{len(unclassified):,} rows are unclassified and excluded. Review the Classification Audit tab in the Excel output.")

    tab1, tab2, tab3, tab4 = st.tabs(["Undiscounted", "Discounted", "Discounting impact", "Classification audit"])
    with tab1:
        st.subheader("Overall portfolio balance base case (Undiscounted)")
        st.dataframe(style_for_streamlit(undiscounted), use_container_width=True)
    with tab2:
        st.subheader("Overall portfolio balance base case (Discounted)")
        st.dataframe(style_for_streamlit(discounted), use_container_width=True)
    with tab3:
        st.subheader("Discounting impact on hedges")
        st.caption("Calculated as Discounted minus Undiscounted.")
        st.dataframe(style_for_streamlit(impact), use_container_width=True)
    with tab4:
        st.dataframe(audit, use_container_width=True, hide_index=True)

    excel_bytes = make_excel(filtered, undiscounted, discounted, impact, detail, audit, years)
    cob = ""
    if "COB" in filtered.columns:
        parsed_cob = pd.to_datetime(filtered["COB"], errors="coerce").max()
        if pd.notna(parsed_cob):
            cob = f"_{parsed_cob:%Y%m%d}"
    st.download_button(
        "Download Excel exposure pack",
        data=excel_bytes,
        file_name=f"IFRS_Exposure_Output{cob}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    with st.expander("Classification logic"):
        st.markdown(
            """
- **HH**: `Final Index = HH`, kept separate from Net Gas.
- **Net Oil**: `Group = Oil` or a recognised oil index.
- **Net Gas**: natural gas/LNG groups, including `Natural Gas & LNG` and `Unsold House Curves`, plus recognised gas/LNG indices.
- **Excluded**: EUA and any freight group/index.
- **ViewGroup**: `Both` rows are included in both the Discounted and Undiscounted tables.
- **Discounting impact**: Discounted minus Undiscounted.
            """
        )


if __name__ == "__main__":
    run_app()
