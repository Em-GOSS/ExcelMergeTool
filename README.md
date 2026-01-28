# ExcelMergeTool

A small GUI tool to append a data unit Excel file into an existing `AllDataList` workbook and regenerate six charts.

## Requirements

- Python 3
- `openpyxl`

Install dependencies:

```bash
pip install openpyxl
```

## Usage

```bash
python app.py
```

1. Choose the `AllDataList` Excel file.
2. Choose a data unit Excel file, or select a folder to batch import all `.xlsx` files inside.
3. Click **开始处理** to append the data and rebuild charts. The workbook will open on the `Charts` sheet after saving.

The tool writes the data block starting in the next available column (7 columns per block with one empty spacer column) and creates/overwrites a `Charts` worksheet containing six line charts.
