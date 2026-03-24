# Advance Material Order

Built by Hbradroc@uwo.ca.

This is a simple non GUI based script that reads BOM and lead-time data, then creates AMO output sheets in Excel.

## Files needed in this folder

- `SYSCAD_AMO.py`
- `allbom.xlsx` (exported BOM from SysCAD)
- `component_db.csv` (component database)
- `leadtime.csv` (lead time document)

## Required columns

### `component_db.csv`
- `Product no`
- `Component no`
- `Component name`
- `Component description`
- `Make/buy`
- `Quantity`
- `Comp To date`

### `leadtime.csv`
- `Item no`
- `Item name`
- `Lead time`

## How the script works

1. Reads every sheet from `allbom.xlsx`.
2. Cleans the sheet and detects product number + quantity columns.
3. Uses `component_db.csv` to explode BOM kits into leaf components.
4. Uses `leadtime.csv` to fetch lead time, supplier, and acquisition info.
5. Keeps parts that meet AMO rules:
   - purchased item
   - lead time greater than 40
   - Comp To date is blank or `99999999`
6. Writes one AMO sheet per source sheet into `AMO_Output_AllSheets.xlsx`.

## Dashboard layout (risk-first)

The `AMO_Dashboard` sheet is organized in this order:

1. `Breakdown by AMO Sheet`
2. `No Safety Stock Items (<= 0)`
3. `Low Safety Stock Items (<= 1.0)`
4. `Longest Lead Time Items`
5. `Supplier Quantities (Low Safety only)`

Charts are placed on the right side of the dashboard to avoid overlap with table content.

## Run

Install dependencies:

```bash
pip install pandas openpyxl
```

Run:

```bash
python SYSCAD_AMO.py
```
