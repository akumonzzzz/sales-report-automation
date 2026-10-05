"""Sales report automation.

Merges messy CSV exports from several sources into one clean Excel report
with summary tables and charts.

Usage:
    python sales_report.py --input data --output report.xlsx
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

log = logging.getLogger("sales_report")

# Every source names its columns differently. Map them all to one schema.
COLUMN_ALIASES = {
    "date": ["date", "order_date", "order date", "sold_on", "transaction date"],
    "product": ["product", "item", "product_name", "product name", "sku_name"],
    "quantity": ["quantity", "qty", "units", "units_sold"],
    "unit_price": ["unit_price", "unit price", "price", "price_each"],
    "region": ["region", "store", "location", "state"],
}

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%b %d, %Y", "%Y/%m/%d"]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY_FORMAT = '"$"#,##0.00'


def parse_date(value) -> pd.Timestamp:
    """Parse a date written in any of the known formats. Returns NaT if none match."""
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    return pd.NaT


def parse_number(series: pd.Series) -> pd.Series:
    """Turn strings like '$1,299.50' or ' 12 ' into numbers. Bad values become NaN."""
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d.\-]", "", regex=True)
        .replace("", pd.NA)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def tidy_product_names(products: pd.Series) -> pd.Series:
    """Give 'USB-C HUB', 'usb-c hub' and 'USB-C Hub' one spelling.

    Names are matched ignoring case and extra spaces. The spelling shown in the
    report is the most common mixed-case one, so 'USB-C Hub' wins over the
    all-caps and all-lowercase versions instead of becoming 'Usb-C Hub'.
    """
    stripped = products.str.strip().str.replace(r"\s+", " ", regex=True)
    key = stripped.str.lower()

    def best_spelling(group: pd.Series) -> str:
        counts = group.value_counts()
        mixed = [name for name in counts.index if name != name.upper() and name != name.lower()]
        return mixed[0] if mixed else counts.index[0].title()

    display = stripped.dropna().groupby(key.dropna()).agg(best_spelling)
    return key.map(display)


def normalise_columns(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """Rename a file's columns to the standard schema, or raise if one is missing."""
    lookup = {col.strip().lower(): col for col in df.columns}
    rename = {}
    for standard, aliases in COLUMN_ALIASES.items():
        match = next((lookup[a] for a in aliases if a in lookup), None)
        if match is None:
            raise ValueError(f"{source}: no column found for '{standard}'")
        rename[match] = standard
    return df.rename(columns=rename)[list(COLUMN_ALIASES)]


def load_files(input_dir: Path) -> pd.DataFrame:
    """Read every CSV in the folder and stack them into one table."""
    files = sorted(input_dir.glob("*.csv"))
    if not files:
        raise FileNotFoundError(f"No CSV files found in {input_dir}")

    frames = []
    for path in files:
        try:
            raw = pd.read_csv(path, dtype=str)
            frame = normalise_columns(raw, path.name)
        except Exception as exc:  # one broken file should not stop the whole run
            log.warning("Skipped %s: %s", path.name, exc)
            continue
        frame["source_file"] = path.name
        frames.append(frame)
        log.info("Loaded %s (%d rows)", path.name, len(frame))

    if not frames:
        raise ValueError("No file could be read")
    return pd.concat(frames, ignore_index=True)


def clean(raw: pd.DataFrame, keep_duplicates: bool = False) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Clean the merged data.

    Returns (good rows, rejected rows with a reason, duplicate rows removed).
    Nothing is dropped silently: every row ends up in exactly one of the three.
    """
    df = raw.copy()
    df["date"] = df["date"].map(parse_date)
    df["quantity"] = parse_number(df["quantity"])
    df["unit_price"] = parse_number(df["unit_price"])
    df["product"] = tidy_product_names(df["product"])
    df["region"] = df["region"].str.strip().str.upper()

    reason = pd.Series("", index=df.index)
    reason[df["date"].isna()] = "invalid or missing date"
    reason[df["product"].isna() | (df["product"] == "")] = "missing product"
    reason[df["quantity"].isna() | (df["quantity"] <= 0) | (df["quantity"] % 1 != 0)] = "invalid quantity"
    reason[df["unit_price"].isna() | (df["unit_price"] < 0)] = "invalid price"

    rejected = raw[reason != ""].copy()
    rejected["reason"] = reason[reason != ""]

    good = df[reason == ""].copy()
    duplicates = good.iloc[0:0]
    if not keep_duplicates:
        # Without an order ID, two identical rows cannot be told apart from a
        # genuine repeat sale, so removed rows are listed in their own sheet.
        is_dup = good.duplicated(subset=["date", "product", "quantity", "unit_price", "region"])
        duplicates = good[is_dup].sort_values("date")
        good = good[~is_dup]
    log.info("Removed %d duplicate rows", len(duplicates))

    good["quantity"] = good["quantity"].astype(int)
    good["revenue"] = (good["quantity"] * good["unit_price"]).round(2)
    good = good.sort_values("date").reset_index(drop=True)
    return good, rejected, duplicates


def build_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Aggregate the clean data into the tables shown in the report."""
    by_month = (
        df.assign(month=df["date"].dt.strftime("%Y-%m"))
        .groupby("month", as_index=False)
        .agg(orders=("revenue", "size"), units=("quantity", "sum"), revenue=("revenue", "sum"))
    )
    by_product = (
        df.groupby("product", as_index=False)
        .agg(units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )
    by_region = (
        df.groupby("region", as_index=False)
        .agg(orders=("revenue", "size"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )
    summary = pd.DataFrame(
        {
            "metric": [
                "Total revenue",
                "Total orders",
                "Units sold",
                "Average order value",
                "Best-selling product",
                "Top region",
                "Period",
            ],
            "value": [
                round(df["revenue"].sum(), 2),
                len(df),
                int(df["quantity"].sum()),
                round(df["revenue"].mean(), 2),
                by_product.iloc[0]["product"],
                by_region.iloc[0]["region"],
                f"{df['date'].min():%d %b %Y} to {df['date'].max():%d %b %Y}",
            ],
        }
    )
    return {
        "Summary": summary,
        "By Month": by_month,
        "By Product": by_product,
        "By Region": by_region,
    }


def style_sheet(ws) -> None:
    """Bold header row, sensible column widths, currency formatting."""
    for cell in ws[1]:
        cell.value = str(cell.value).replace("_", " ").title()
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToHeight = 0

    for idx, column in enumerate(ws.columns, start=1):
        header = str(column[0].value or "").lower().replace(" ", "_")
        width = max(len(str(c.value)) if c.value is not None else 0 for c in column) + 3
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width, 12), 40)
        if header in {"revenue", "unit_price"}:
            for cell in column[1:]:
                cell.number_format = MONEY_FORMAT
        if header == "date":
            for cell in column[1:]:
                cell.number_format = "DD/MM/YYYY"


def add_charts(wb) -> None:
    """Put a revenue trend line and a revenue-by-product bar chart on the Summary sheet."""
    summary = wb["Summary"]

    ws = wb["By Month"]
    line = LineChart()
    line.title = "Revenue by month"
    line.y_axis.title = "Revenue ($)"
    line.y_axis.number_format = '"$"#,##0'
    line.legend = None
    line.height, line.width = 7.5, 16
    line.add_data(Reference(ws, min_col=4, min_row=1, max_row=ws.max_row), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    line.series[0].smooth = False  # straight segments: a smoothed curve invents values
    summary.add_chart(line, "D2")

    ws = wb["By Product"]
    bar = BarChart()
    bar.title = "Revenue by product"
    bar.y_axis.title = "Revenue ($)"
    bar.y_axis.number_format = '"$"#,##0'
    bar.legend = None
    bar.height, bar.width = 7.5, 16
    bar.add_data(Reference(ws, min_col=3, min_row=1, max_row=ws.max_row), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    summary.add_chart(bar, "D18")


def write_report(
    tables: dict, clean_df: pd.DataFrame, rejected: pd.DataFrame, duplicates: pd.DataFrame, output: Path
) -> None:
    """Write all sheets to Excel, then format them and add charts."""
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name, index=False)
        clean_df.to_excel(writer, sheet_name="Clean Data", index=False)
        if not rejected.empty:
            rejected.to_excel(writer, sheet_name="Rejected Rows", index=False)
        if not duplicates.empty:
            duplicates.to_excel(writer, sheet_name="Duplicates Removed", index=False)

    wb = load_workbook(output)
    for ws in wb.worksheets:
        style_sheet(ws)
    for row in wb["Summary"].iter_rows(min_row=2, max_col=2):
        if row[0].value in {"Total revenue", "Average order value"}:
            row[1].number_format = MONEY_FORMAT
        row[1].alignment = Alignment(horizontal="right")
    add_charts(wb)
    wb.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description="Merge sales CSVs into one Excel report.")
    parser.add_argument("--input", type=Path, default=Path("data"), help="folder with CSV files")
    parser.add_argument("--output", type=Path, default=Path("report.xlsx"), help="Excel file to create")
    parser.add_argument(
        "--keep-duplicates", action="store_true",
        help="keep identical rows (use when repeat sales on the same day are normal)",
    )
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")

    try:
        raw = load_files(args.input)
        clean_df, rejected, duplicates = clean(raw, args.keep_duplicates)
        if clean_df.empty:
            raise ValueError("No valid rows left after cleaning")
        tables = build_tables(clean_df)
        write_report(tables, clean_df, rejected, duplicates, args.output)
    except Exception as exc:
        log.error("%s", exc)
        return 1

    log.info(
        "Done: %d rows read, %d kept, %d rejected, %d duplicates removed -> %s",
        len(raw), len(clean_df), len(rejected), len(duplicates), args.output,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
