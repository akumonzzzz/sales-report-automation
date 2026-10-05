# Sales Report Automation

Turns a folder of messy sales CSV exports into one clean Excel report, in one command.

![Report preview](docs/report-preview.png)

## The problem

A small retailer gets a sales export from each store and from its online shop. Every file uses different column names, date formats and currency formatting, and some rows are duplicated or broken. Someone spends hours each month copying them into one spreadsheet by hand.

## What the script does

1. Reads every CSV in a folder.
2. Maps each file's column names to one schema (`order_date`, `Sold_On` and `Date` all become `date`).
3. Parses dates in six formats (`2026-03-12`, `12/03/2026`, `Mar 12, 2026`, ...).
4. Cleans numbers such as `$1,299.50`.
5. Gives product names one spelling (`USB-C HUB` and `usb-c hub` become `USB-C Hub`).
6. Sets aside invalid rows and duplicate rows, each in its own sheet, so nothing disappears silently.
7. Writes a formatted Excel report with summary figures and charts.

## Sample run

```
$ python sales_report.py --input data --output report.xlsx
INFO  Loaded online_orders.csv (43 rows)
INFO  Loaded store_melbourne.csv (56 rows)
INFO  Loaded store_sydney.csv (60 rows)
INFO  Removed 5 duplicate rows
INFO  Done: 159 rows read, 149 kept, 5 rejected, 5 duplicates removed -> report.xlsx
```

The report has these sheets:

| Sheet | Contents |
| --- | --- |
| Summary | Total revenue, orders, units, average order value, best product, top region, two charts |
| By Month | Orders, units and revenue per month |
| By Product | Units and revenue per product |
| By Region | Orders and revenue per region |
| Clean Data | Every valid row after cleaning |
| Rejected Rows | Invalid rows with the reason (bad date, missing product, invalid quantity or price) |
| Duplicates Removed | Identical rows that were dropped |

## How to run it

Requires Python 3.9 or newer.

```
pip install -r requirements.txt
python make_sample_data.py      # creates the sample CSVs in data/
python sales_report.py --input data --output report.xlsx
```

Options:

| Option | Default | Meaning |
| --- | --- | --- |
| `--input` | `data` | Folder containing the CSV files |
| `--output` | `report.xlsx` | Excel file to create |
| `--keep-duplicates` | off | Keep identical rows instead of removing them |

## Notes on duplicates

The sample data has no order ID, so two rows with the same date, product, quantity, price and region cannot be told apart from a genuine repeat sale. The script removes them by default and lists every removed row in the `Duplicates Removed` sheet for review. Use `--keep-duplicates` when repeat sales on the same day are normal.

## Adapting it to other data

- New column name: add it to `COLUMN_ALIASES` in `sales_report.py`.
- New date format: add it to `DATE_FORMATS`.
- A file that cannot be read is skipped with a warning, and the rest of the run continues.

The sample data is generated and does not come from a real business.

## Built with

Python, pandas, openpyxl.
