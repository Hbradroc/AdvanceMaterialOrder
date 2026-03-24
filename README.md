# AMO Extractor

Generate AMO candidate sheets from:
- exported BOM workbook
- component database CSV
- lead time CSV

This project includes a reconstructed Python script (`amo_tool.py`) equivalent to the provided executable workflow.

## Required Input Files

Place these files in the same folder as `amo_tool.py`:

1. `allbom.xlsx`  
   Exported BOM from SysCAD (can contain multiple sheets).

2. `component_db.csv`  
   Component database file. Required columns:
   - `Product no`
   - `Component no`
   - `Component name`
   - `Component description`
   - `Make/buy`
   - `Quantity`
   - `Comp To date`

3. `leadtime.csv`  
   Lead-time document. Required columns:
   - `Item no`
   - `Item name`
   - `Lead time`

   Optional columns (auto-detected if present):
   - Acquisition: `Acquisition code` / `Acquisition` / `Acq code` / `Acq. code`
   - Supplier/Vendor variants (for supplier display in output)
   - Safety stock variants (for safety stock display/filtering)

## Output

Running the script creates:
- `AMO_Output_AllSheets.xlsx`

For each BOM sheet in `allbom.xlsx`, the script creates one output sheet prefixed with `AMO_`.

## How AMO Selection Works

An item is included in AMO output only when all conditions pass:

1. Item appears in demanded BOM quantity (including exploded child quantities).
2. Item is **purchased** (from `Acq` in leadtime or `Make/buy` fallback).
3. `Lead time > 40`.
4. `Comp To date` is blank or equals `99999999`.

## Run Locally

## 1) Install Python dependencies

```bash
pip install pandas openpyxl
```

## 2) Run

```bash
python amo_tool.py
```

## Expected console output

```text
Done! Created: AMO_Output_AllSheets.xlsx
Processed sheets: ...
Applied filter Safety stock = 0 to all AMO sheets.
```

## Notes

- Sheet names are sanitized/truncated to Excel limits.
- An Excel auto-filter is applied on `Safety stock` to show zero values.
- If input filenames differ, rename them to the required names above (or update constants in `amo_tool.py`).
