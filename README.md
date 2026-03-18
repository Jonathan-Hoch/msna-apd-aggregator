# MSNA APD Aggregator (v1.2.0)

## Overview

This tool automates the aggregation of multi-sheet Excel exports from muscle sympathetic nerve activity (MSNA) microneurography action potential detection software (APD) analysis. Each input file represents a single recording; the tool merges all files in a folder into a single master spreadsheet (`APD_Master.xlsx`) and computes burst-level statistics across files. Example data (randomized with realistic values) are provided in the `example_data` folder.

## Key Features

- **GUI Interface:** Simple folder selection for non-technical users — no coding required.

![GUI Interface](screenshots/FolderSelection.png)

- **Statistical QA:** Automatic Z-score flagging (1σ, 2σ, 3σ) of outliers across burst amplitude, AP frequency, and APs/burst, applied as color-coded cell formatting in the output sheet.
- **Audit Logging:** Generates a processing log (`processing_log.txt`) listing every file processed, with errors noted by filename.
- **Headless Mode:** Can be run without the GUI by setting file paths directly in the script (see below).

Upon successful execution, a confirmation window displays the output filename.

![Success State](screenshots/Success.png)

## Output

Running the tool produces two files in the selected export folder:

- `APD_Master.xlsx` — master spreadsheet with one row per input file, computed statistics, and SD-based color flagging
- `processing_log.txt` — list of all files processed, with any errors noted

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/Jonathan-Hoch/msna-apd-aggregator.git
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the app:
   ```
   python msna_aggregate.py
   ```

## Headless Use (No GUI)

If you prefer to run the script directly without the GUI:

1. Open `msna_aggregate.py`
2. Set `INPUT_DIR` and `OUTPUT_FILE` in the config block near the top of the file
3. Call `process_data(INPUT_DIR, OUTPUT_FILE.parent)` directly
4. Delete or ignore everything below the `process_data()` function

## Technical Notes

Built with Python using pandas for data transformation and openpyxl for Excel writing and conditional formatting. `xlrd` is included for legacy `.xls` file support; modern `.xlsx` files are handled by openpyxl.

## Requirements

See `requirements.txt`. Key dependencies:

```
pandas>=3.0.1
openpyxl>=3.1.5
customtkinter>=5.2.2
numpy>=2.4.3
xlrd>=2.0.1
```

> Note: pandas 3.0+ requires Python 3.11 or higher. If you are on an older Python version, use `pandas>=2.2.0` instead.

## License

MIT
