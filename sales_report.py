#!/usr/bin/env python3
"""Sales report automation.

Merges messy sales exports (CSV and Excel) from several sources into one clean
Excel report with summary figures, breakdown tables and charts.

Usage:
    python sales_report.py --input data --output report.xlsx
    python sales_report.py --input data --config config.json

Every input row ends up in exactly one place: the clean data, the rejected
rows (with the reason), or the removed duplicates. Nothing is dropped silently.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

__version__ = "1.1.0"

log = logging.getLogger("sales_report")

REQUIRED_COLUMNS = ["date", "product", "quantity", "unit_price", "region"]
OPTIONAL_COLUMNS = ["order_id"]
INPUT_SUFFIXES = {".csv", ".xlsx", ".xlsm"}
DELIMITERS = [",", ";", "\t", "|"]
NO_REGION = "(NOT SPECIFIED)"

# Built-in settings. A config file can extend the aliases and replace the rest.
DEFAULT_CONFIG = {
    # Every source names its columns differently. Map them all to one schema.
    # Matching ignores case, underscores, hyphens and extra spaces.
    "column_aliases": {
        "order_id": ["order id", "order no", "order number", "invoice no", "invoice number", "transaction id"],
        "date": ["date", "order date", "invoice date", "sold on", "transaction date"],
        "product": ["product", "item", "product name", "description", "sku name"],
        "quantity": ["quantity", "qty", "units", "units sold"],
        "unit_price": ["unit price", "price", "price each"],
        "region": ["region", "store", "location", "state"],
    },
    # Tried in order, so day-first comes before month-first: 03/04/2026 is 3 April.
    "date_formats": [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d/%m/%y",
        "%d %b %Y",
        "%d %B %Y",
        "%b %d, %Y",
        "%B %d, %Y",
        "%Y/%m/%d",
    ],
    "decimal_separator": ".",
    "currency_symbol": "$",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
CHART_COLOUR = "1F4E78"
INT_FORMAT = "#,##0"
PERCENT_FORMAT = "0.0%"
GROWTH_FORMAT = "+0.0%;[Red]-0.0%"
DATE_FORMAT = "DD/MM/YYYY"
# Column headings that need more than 'snake_case' -> 'Title Case'.
HEADER_LABELS = {
    "order_id": "Order ID",
    "share_of_revenue": "Share of Revenue",
    "growth_vs_prev_month": "Growth vs Prev Month",
    "duplicate_of": "Duplicate of",
}


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

def load_config(path: Path | None) -> dict:
    """Return the settings to use: the defaults, updated by an optional JSON file."""
    config = {
        "column_aliases": {name: list(aliases) for name, aliases in DEFAULT_CONFIG["column_aliases"].items()},
        "date_formats": list(DEFAULT_CONFIG["date_formats"]),
        "decimal_separator": DEFAULT_CONFIG["decimal_separator"],
        "currency_symbol": DEFAULT_CONFIG["currency_symbol"],
    }
    if path is None:
        return config

    try:
        user = json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"{path} is not valid JSON: {exc}") from exc

    unknown = set(user) - set(config)
    if unknown:
        raise ValueError(f"{path}: unknown setting(s) {sorted(unknown)}. Allowed: {sorted(config)}")

    for name, aliases in user.get("column_aliases", {}).items():
        if name not in config["column_aliases"]:
            raise ValueError(f"{path}: unknown column '{name}'. Allowed: {sorted(config['column_aliases'])}")
        # New aliases are added in front of the built-in ones.
        config["column_aliases"][name] = list(dict.fromkeys(list(aliases) + config["column_aliases"][name]))

    if "date_formats" in user:
        if not user["date_formats"]:
            raise ValueError(f"{path}: date_formats must not be empty")
        config["date_formats"] = list(user["date_formats"])

    if user.get("decimal_separator", ".") not in {".", ","}:
        raise ValueError(f"{path}: decimal_separator must be '.' or ','")
    config["decimal_separator"] = user.get("decimal_separator", config["decimal_separator"])
    config["currency_symbol"] = user.get("currency_symbol", config["currency_symbol"])
    return config


# --------------------------------------------------------------------------- #
# Parsing helpers
# --------------------------------------------------------------------------- #

def is_blank(values: pd.Series) -> pd.Series:
    """True where a value is missing or only whitespace."""
    text = values.astype("string").str.strip()
    return (text.isna() | (text == "")).astype(bool)


def parse_dates(values: pd.Series, formats: list[str]) -> pd.Series:
    """Parse dates written in any of the given formats. Unparseable values become NaT.

    Formats are tried in order and the first one that fits a value wins. A date
    that does not exist, such as 30 February, fits none and becomes NaT.
    """
    text = values.astype("string").str.strip()
    result = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")
    for fmt in formats:
        todo = (result.isna() & text.notna()).astype(bool)
        if not todo.any():
            break
        parsed = pd.to_datetime(text[todo].astype(object), format=fmt, errors="coerce")
        result.loc[todo] = parsed.astype("datetime64[ns]")
    return result


def parse_number(values: pd.Series, decimal_separator: str = ".") -> pd.Series:
    """Turn strings like '$1,299.50', ' 12 ' or '(45.00)' into numbers.

    Currency symbols, letters and spaces are ignored and brackets mean negative.
    Anything that is not clearly a number becomes NaN, so a value such as
    '1.299,50' read with the wrong decimal separator is rejected and never
    silently turned into a different amount.
    """
    text = values.astype("string").str.strip()
    bracketed = text.str.fullmatch(r"\(.*\)").fillna(False).astype(bool)
    cleaned = text.str.replace(r"[^\d.,\-]", "", regex=True)
    if decimal_separator == ",":
        cleaned = cleaned.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        valid = cleaned.str.fullmatch(r"-?\d+(\.\d+)?")
    else:
        valid = cleaned.str.fullmatch(r"-?(\d+|\d{1,3}(,\d{3})+)(\.\d+)?")
        cleaned = cleaned.str.replace(",", "", regex=False)
    valid = valid.fillna(False).astype(bool)
    numbers = pd.to_numeric(cleaned.where(valid).astype(object), errors="coerce").astype("float64")
    return numbers.where(~bracketed, -numbers)


def tidy_product_names(products: pd.Series) -> pd.Series:
    """Give 'USB-C HUB', 'usb-c hub' and 'USB-C Hub' one spelling.

    Names are matched ignoring case and extra spaces. The spelling shown in the
    report is the most common mixed-case one, so 'USB-C Hub' wins over the
    all-caps and all-lowercase versions instead of becoming 'Usb-C Hub'.
    """
    stripped = products.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
    stripped = stripped.where(stripped != "")
    key = stripped.str.lower()

    def best_spelling(group: pd.Series) -> str:
        counts = group.value_counts()
        ranked = sorted(counts.index, key=lambda name: (-counts[name], name))
        mixed = [name for name in ranked if name != name.upper() and name != name.lower()]
        return mixed[0] if mixed else ranked[0].title()

    known = stripped.notna()
    display = stripped[known].astype(object).groupby(key[known].astype(object)).agg(best_spelling)
    return key.astype(object).map(display)


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #

def _header_key(name) -> str:
    """Normalise a column name so 'Units_Sold', 'units sold' and 'UNITS-SOLD' match."""
    return re.sub(r"[\s_\-]+", " ", str(name).strip().lower())


def normalise_columns(df: pd.DataFrame, source: str, aliases: dict[str, list[str]]) -> pd.DataFrame:
    """Rename a file's columns to the standard schema, or raise if a required one is missing."""
    lookup: dict[str, str] = {}
    for column in df.columns:
        lookup.setdefault(_header_key(column), column)

    rename, missing = {}, []
    for standard in OPTIONAL_COLUMNS + REQUIRED_COLUMNS:
        candidates = [_header_key(standard)] + [_header_key(alias) for alias in aliases.get(standard, [])]
        match = next((lookup[name] for name in candidates if name in lookup), None)
        if match is not None:
            rename[match] = standard
        elif standard in REQUIRED_COLUMNS:
            missing.append(standard)

    if missing:
        raise ValueError(
            f"{source}: no column found for {missing}. Columns in the file: {list(df.columns)}. "
            "Add the file's column names to column_aliases in the config file."
        )

    out = df.rename(columns=rename)
    for column in OPTIONAL_COLUMNS:
        if column not in out.columns:
            out[column] = pd.NA
    return out[OPTIONAL_COLUMNS + REQUIRED_COLUMNS]


def detect_delimiter(path: Path, encoding: str) -> str:
    """Pick the delimiter that appears most often in the header line."""
    with open(path, encoding=encoding, newline="") as handle:
        for line in handle:
            if line.strip():
                return max(DELIMITERS, key=line.count)
    raise ValueError("file is empty")


def read_source(path: Path) -> pd.DataFrame:
    """Read one CSV or Excel file as text, keeping every row so row numbers match the file."""
    options = dict(dtype=str, keep_default_na=False, na_values=[""])
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, **options)

    for encoding in ("utf-8-sig", "cp1252"):
        try:
            delimiter = detect_delimiter(path, encoding)
            return pd.read_csv(
                path, sep=delimiter, encoding=encoding, skip_blank_lines=False, skipinitialspace=True, **options
            )
        except UnicodeDecodeError:
            continue
    raise ValueError("could not read the file as UTF-8 or Windows-1252 text")


def find_input_files(input_dir: Path, output: Path | None = None) -> list[Path]:
    """List the CSV and Excel files in a folder, skipping temp files and the report itself."""
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")
    files = [
        path
        for path in sorted(input_dir.iterdir())
        if path.is_file()
        and path.suffix.lower() in INPUT_SUFFIXES
        and not path.name.startswith(("~$", "."))
        and (output is None or path.resolve() != Path(output).resolve())
    ]
    if not files:
        raise FileNotFoundError(f"No CSV or Excel files found in {input_dir}")
    return files


def load_files(files: list[Path], config: dict) -> tuple[pd.DataFrame, list[tuple[str, str]]]:
    """Read every file and stack the rows into one table.

    Returns (rows, skipped files). A file that cannot be read is skipped with a
    warning so one broken export does not stop the whole run.
    """
    frames, skipped = [], []
    for path in files:
        try:
            frame = normalise_columns(read_source(path), path.name, config["column_aliases"])
        except Exception as exc:
            log.warning("Skipped %s: %s", path.name, exc)
            skipped.append((path.name, str(exc)))
            continue

        frame["source_file"] = path.name
        frame["source_row"] = range(2, len(frame) + 2)  # row 1 is the header
        blank = pd.concat([is_blank(frame[column]) for column in REQUIRED_COLUMNS], axis=1).all(axis=1)
        frame = frame[~blank]
        frames.append(frame)
        log.info("Loaded %s (%d rows)", path.name, len(frame))

    if not frames:
        raise ValueError("None of the input files could be read")
    return pd.concat(frames, ignore_index=True), skipped


# --------------------------------------------------------------------------- #
# Cleaning
# --------------------------------------------------------------------------- #

def clean(
    raw: pd.DataFrame, config: dict | None = None, keep_duplicates: bool = False
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Clean the merged rows.

    Returns (good rows, rejected rows with the reason, duplicate rows removed).
    """
    config = config or DEFAULT_CONFIG
    df = raw.copy()
    order_id = raw["order_id"].astype("string").str.strip()
    df["order_id"] = order_id.astype(object).where(~is_blank(raw["order_id"]), None)
    df["date"] = parse_dates(raw["date"], config["date_formats"])
    df["quantity"] = parse_number(raw["quantity"], config["decimal_separator"])
    df["unit_price"] = parse_number(raw["unit_price"], config["decimal_separator"])
    df["product"] = tidy_product_names(raw["product"])
    region = raw["region"].astype("string").str.strip().str.upper()
    df["region"] = region.astype(object).where(~is_blank(raw["region"]), NO_REGION)

    checks = {
        "missing date": is_blank(raw["date"]),
        "unrecognised date": ~is_blank(raw["date"]) & df["date"].isna(),
        "missing product": is_blank(raw["product"]),
        "missing quantity": is_blank(raw["quantity"]),
        "quantity is not a number": ~is_blank(raw["quantity"]) & df["quantity"].isna(),
        "quantity is not a whole number": df["quantity"].notna() & (df["quantity"] % 1 != 0),
        "quantity is zero or negative": df["quantity"].notna() & (df["quantity"] <= 0),
        "missing price": is_blank(raw["unit_price"]),
        "price is not a number": ~is_blank(raw["unit_price"]) & df["unit_price"].isna(),
        "negative price": df["unit_price"].notna() & (df["unit_price"] < 0),
    }
    reason = pd.Series("", index=df.index, dtype=object)
    for label, failed in checks.items():
        failed = failed.astype(bool)
        reason[failed] = reason[failed] + label + "; "
    reason = reason.str.rstrip("; ")
    bad = reason != ""

    rejected = raw[bad].copy()
    rejected["reason"] = reason[bad]
    rejected = rejected[["source_file", "source_row", "reason"] + OPTIONAL_COLUMNS + REQUIRED_COLUMNS]

    good = df[~bad].copy()
    duplicates = good.iloc[0:0].assign(duplicate_of=pd.Series(dtype=object))
    if not keep_duplicates and not good.empty:
        # Rows count as duplicates when every field matches, order ID included.
        # With an order ID, two identical sales on different orders are both
        # kept. Without one they cannot be told apart, so the removed rows are
        # listed in their own sheet for review.
        key = OPTIONAL_COLUMNS + REQUIRED_COLUMNS
        origin = good["source_file"] + " row " + good["source_row"].astype(str)
        first_seen = origin.groupby([good[column] for column in key], dropna=False).transform("first")
        is_duplicate = good.duplicated(subset=key)
        duplicates = good[is_duplicate].assign(duplicate_of=first_seen[is_duplicate])
        good = good[~is_duplicate]
    log.info("Rejected %d invalid rows, removed %d duplicate rows", len(rejected), len(duplicates))

    good = good.assign(
        quantity=good["quantity"].astype(int),
        revenue=(good["quantity"] * good["unit_price"]).round(2),
    )
    columns = OPTIONAL_COLUMNS + ["date", "product", "quantity", "unit_price", "revenue", "region"]
    good = good.sort_values(["date", "source_file", "source_row"]).reset_index(drop=True)
    good = good[columns + ["source_file", "source_row"]]
    duplicates = duplicates.sort_values(["source_file", "source_row"]).reset_index(drop=True)
    duplicates = duplicates[OPTIONAL_COLUMNS + REQUIRED_COLUMNS + ["source_file", "source_row", "duplicate_of"]]
    return good, rejected.reset_index(drop=True), duplicates


# --------------------------------------------------------------------------- #
# Aggregation
# --------------------------------------------------------------------------- #

def build_quality_table(
    raw: pd.DataFrame, good: pd.DataFrame, rejected: pd.DataFrame, duplicates: pd.DataFrame
) -> pd.DataFrame:
    """Count, per source file, where every row went. Read = kept + rejected + duplicates."""
    table = pd.DataFrame({"rows_read": raw.groupby("source_file").size()})
    table["rows_kept"] = good.groupby("source_file").size()
    table["rows_rejected"] = rejected.groupby("source_file").size()
    table["duplicates_removed"] = duplicates.groupby("source_file").size()
    table = table.fillna(0).astype(int)
    table["first_date"] = good.groupby("source_file")["date"].min()
    table["last_date"] = good.groupby("source_file")["date"].max()
    table = table.reset_index()

    counts = ["rows_read", "rows_kept", "rows_rejected", "duplicates_removed"]
    total = {"source_file": "TOTAL", **table[counts].sum().to_dict()}
    total["first_date"], total["last_date"] = good["date"].min(), good["date"].max()
    return pd.concat([table, pd.DataFrame([total])], ignore_index=True)


def build_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Aggregate the clean data into the breakdown tables shown in the report."""
    total = df["revenue"].sum()
    month = df["date"].dt.strftime("%Y-%m")

    by_month = (
        df.assign(month=month)
        .groupby("month", as_index=False)
        .agg(transactions=("revenue", "size"), units=("quantity", "sum"), revenue=("revenue", "sum"))
    )
    by_month["revenue"] = by_month["revenue"].round(2)
    by_month["avg_transaction_value"] = (by_month["revenue"] / by_month["transactions"]).round(2)
    by_month["growth_vs_prev_month"] = by_month["revenue"].pct_change()

    by_product = (
        df.groupby("product", as_index=False)
        .agg(units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values(["revenue", "product"], ascending=[False, True], ignore_index=True)
    )
    by_product["revenue"] = by_product["revenue"].round(2)
    by_product["share_of_revenue"] = by_product["revenue"] / total
    by_product["avg_unit_price"] = (by_product["revenue"] / by_product["units"]).round(2)

    by_region = (
        df.groupby("region", as_index=False)
        .agg(transactions=("revenue", "size"), units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values(["revenue", "region"], ascending=[False, True], ignore_index=True)
    )
    by_region["revenue"] = by_region["revenue"].round(2)
    by_region["share_of_revenue"] = by_region["revenue"] / total

    region_by_month = (
        df.assign(month=month)
        .pivot_table(index="month", columns="region", values="revenue", aggfunc="sum", fill_value=0)
        .reindex(columns=by_region["region"])
    )
    region_by_month["Total"] = region_by_month.sum(axis=1)
    region_by_month = region_by_month.round(2).reset_index()
    region_by_month.columns.name = None

    return {
        "By Month": by_month,
        "By Product": by_product,
        "By Region": by_region,
        "Region by Month": region_by_month,
    }


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #

def _excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Replace missing values with empty cells so openpyxl can write them."""
    return df.astype(object).where(df.notna(), None)


def _style_header(cells) -> None:
    for cell in cells:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _page_setup(ws) -> None:
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToHeight = 0


def _format_table_sheet(ws, formats: dict[str, str], default_format: str | None = None) -> None:
    """Style a sheet that holds one table: header, widths, number formats, frozen header."""
    headers = [cell.value for cell in ws[1]]
    for idx, (header, column) in enumerate(zip(headers, ws.columns), start=1):
        number_format = formats.get(header, default_format if idx > 1 else None)
        if number_format:
            for cell in column[1:]:
                cell.number_format = number_format
        if isinstance(header, str) and header == header.lower():
            # 'unit_price' -> 'Unit Price'; region names such as 'NSW' are left alone
            column[0].value = HEADER_LABELS.get(header, header.replace("_", " ").title())
        longest = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 3, 12), 60)
    _style_header(ws[1])
    ws.freeze_panes = "A2"
    _page_setup(ws)


def _chart(kind, title: str, ws, value_col: int, anchor: str, target, y_title: str, y_format: str) -> None:
    """Add one single-series chart reading its data from a table sheet."""
    chart = kind()
    chart.title = title
    chart.legend = None
    chart.height, chart.width = 7.5, 15
    chart.y_axis.title = y_title
    chart.y_axis.number_format = y_format
    # openpyxl 3.1 writes axes as deleted unless told otherwise, which hides them in Excel.
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.add_data(Reference(ws, min_col=value_col, min_row=1, max_row=ws.max_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    series = chart.series[0]
    if kind is LineChart:
        series.smooth = False  # straight segments: a smoothed curve invents values
        series.graphicalProperties.line.solidFill = CHART_COLOUR
        series.graphicalProperties.line.width = 28575
    else:
        chart.gapWidth = 60
        series.graphicalProperties.solidFill = CHART_COLOUR
    target.add_chart(chart, anchor)


def _write_summary(wb, tables: dict, good: pd.DataFrame, quality: pd.DataFrame, money: str) -> None:
    """Build the Summary sheet: key figures, data quality counts and four charts."""
    ws = wb.create_sheet("Summary", 0)
    by_month, by_product, by_region = tables["By Month"], tables["By Product"], tables["By Region"]
    best_month = by_month.loc[by_month["revenue"].idxmax()]
    totals = quality.iloc[-1]

    ws["A1"] = "Sales Report"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    ws["A2"] = f"{good['date'].min():%d %b %Y} to {good['date'].max():%d %b %Y}"
    ws["A2"].font = Font(italic=True, color="595959")

    figures = [
        ("Total revenue", float(round(good["revenue"].sum(), 2)), money),
        ("Transactions", len(good), INT_FORMAT),
        ("Units sold", int(good["quantity"].sum()), INT_FORMAT),
        ("Average transaction value", float(round(good["revenue"].mean(), 2)), money),
        ("Top product by revenue", by_product.iloc[0]["product"], None),
        ("Top region by revenue", by_region.iloc[0]["region"], None),
        ("Best month", best_month["month"], None),
    ]
    checks = [
        ("Files read", len(quality) - 1, INT_FORMAT),
        ("Rows read", int(totals["rows_read"]), INT_FORMAT),
        ("Rows kept", int(totals["rows_kept"]), INT_FORMAT),
        ("Rows rejected", int(totals["rows_rejected"]), INT_FORMAT),
        ("Duplicates removed", int(totals["duplicates_removed"]), INT_FORMAT),
    ]

    row = 4
    for heading, block in (("Key figures", figures), ("Data quality", checks)):
        ws.cell(row=row, column=1, value=heading)
        ws.cell(row=row, column=2, value="Value")
        _style_header(ws[row][:2])
        for label, value, number_format in block:
            row += 1
            ws.cell(row=row, column=1, value=label)
            cell = ws.cell(row=row, column=2, value=value)
            cell.alignment = Alignment(horizontal="right")
            if number_format:
                cell.number_format = number_format
        row += 2
    note = ws.cell(row=row - 1, column=1, value="Details: Data Quality, Rejected Rows, Duplicates Removed")
    note.font = Font(italic=True, size=9, color="595959")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 3

    money_axis = money.replace(".00", "")
    month_cols, product_cols, region_cols = list(by_month.columns), list(by_product.columns), list(by_region.columns)
    _chart(LineChart, "Revenue by month", wb["By Month"], month_cols.index("revenue") + 1, "D4", ws,
           "Revenue", money_axis)
    _chart(BarChart, "Revenue by region", wb["By Region"], region_cols.index("revenue") + 1, "N4", ws,
           "Revenue", money_axis)
    _chart(BarChart, "Revenue by product", wb["By Product"], product_cols.index("revenue") + 1, "D20", ws,
           "Revenue", money_axis)
    _chart(BarChart, "Units sold by product", wb["By Product"], product_cols.index("units") + 1, "N20", ws,
           "Units", INT_FORMAT)
    _page_setup(ws)
    wb.active = 0


def write_report(
    tables: dict[str, pd.DataFrame],
    good: pd.DataFrame,
    rejected: pd.DataFrame,
    duplicates: pd.DataFrame,
    quality: pd.DataFrame,
    output: Path,
    currency_symbol: str = "$",
) -> None:
    """Write every sheet of the Excel report."""
    money = f'"{currency_symbol}"#,##0.00'
    formats = {
        "revenue": money, "unit_price": money, "avg_transaction_value": money, "avg_unit_price": money,
        "share_of_revenue": PERCENT_FORMAT, "growth_vs_prev_month": GROWTH_FORMAT,
        "date": DATE_FORMAT, "first_date": DATE_FORMAT, "last_date": DATE_FORMAT,
        "transactions": INT_FORMAT, "units": INT_FORMAT, "quantity": INT_FORMAT,
        "rows_read": INT_FORMAT, "rows_kept": INT_FORMAT, "rows_rejected": INT_FORMAT,
        "duplicates_removed": INT_FORMAT,
    }
    if good["order_id"].isna().all():  # no source had order IDs: leave the empty column out
        good, rejected, duplicates = (frame.drop(columns="order_id") for frame in (good, rejected, duplicates))

    sheets = dict(tables)
    sheets["Clean Data"] = good
    sheets["Data Quality"] = quality
    if not rejected.empty:
        sheets["Rejected Rows"] = rejected
    if not duplicates.empty:
        sheets["Duplicates Removed"] = duplicates

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, table in sheets.items():
            _excel_safe(table).to_excel(writer, sheet_name=name, index=False)
        wb = writer.book
        for name in sheets:
            if name == "Rejected Rows":  # raw text exactly as it was in the source file
                _format_table_sheet(wb[name], {})
            else:
                _format_table_sheet(wb[name], formats, money if name == "Region by Month" else None)
        wb["Clean Data"].auto_filter.ref = wb["Clean Data"].dimensions
        for cell in wb["Data Quality"][wb["Data Quality"].max_row]:
            cell.font = Font(bold=True)
        _write_summary(wb, tables, good, quality, money)


# --------------------------------------------------------------------------- #
# Command line
# --------------------------------------------------------------------------- #

@dataclass
class RunResult:
    """What a run did, for logging and for callers that import this module."""

    output: Path
    files_read: int
    rows_read: int
    rows_kept: int
    rows_rejected: int
    duplicates_removed: int
    total_revenue: float
    skipped_files: list[tuple[str, str]] = field(default_factory=list)


def run(input_dir: Path, output: Path, config: dict | None = None, keep_duplicates: bool = False) -> RunResult:
    """Run the whole pipeline: read, clean, aggregate, write the report."""
    config = config or load_config(None)
    files = find_input_files(Path(input_dir), output)
    raw, skipped = load_files(files, config)
    good, rejected, duplicates = clean(raw, config, keep_duplicates)
    if good.empty:
        raise ValueError("No valid rows left after cleaning")
    quality = build_quality_table(raw, good, rejected, duplicates)
    tables = build_tables(good)
    write_report(tables, good, rejected, duplicates, quality, output, config["currency_symbol"])
    return RunResult(
        output=Path(output),
        files_read=raw["source_file"].nunique(),
        rows_read=len(raw),
        rows_kept=len(good),
        rows_rejected=len(rejected),
        duplicates_removed=len(duplicates),
        total_revenue=float(round(good["revenue"].sum(), 2)),
        skipped_files=skipped,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Merge sales CSV and Excel exports into one Excel report.")
    parser.add_argument("--input", type=Path, default=Path("data"), help="folder with CSV/Excel files (default: data)")
    parser.add_argument("--output", type=Path, default=Path("report.xlsx"), help="report to create (default: report.xlsx)")
    parser.add_argument("--config", type=Path, help="JSON file with extra column names, date formats or currency")
    parser.add_argument(
        "--keep-duplicates", action="store_true",
        help="keep identical rows (use when repeat sales on the same day are normal)",
    )
    parser.add_argument("--log-file", type=Path, help="also append the log to this file (useful for scheduled runs)")
    parser.add_argument("--quiet", action="store_true", help="only print warnings and errors")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if args.log_file:
        handlers.append(logging.FileHandler(args.log_file, encoding="utf-8"))
    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s" if args.log_file else "%(levelname)-7s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )

    try:
        result = run(args.input, args.output, load_config(args.config), args.keep_duplicates)
    except PermissionError:
        log.error("Cannot write %s. Close the file in Excel and run again.", args.output)
        return 1
    except Exception as exc:
        log.error("%s", exc)
        return 1

    log.info(
        "Done: %d files, %d rows read = %d kept + %d rejected + %d duplicates removed",
        result.files_read, result.rows_read, result.rows_kept, result.rows_rejected, result.duplicates_removed,
    )
    log.info("Report written to %s", result.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
