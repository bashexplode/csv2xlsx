# CSVs to Excel (one sheet per CSV)

Combine all CSV files from a folder into a single Excel workbook, with each CSV becoming its own worksheet (tab). The script auto-detects CSV delimiters, sanitizes sheet names to satisfy Excel rules, and streams data efficiently for large files.

## Features
- One worksheet per CSV file
- Auto-detects common delimiters (comma, semicolon, tab, pipe)
- Optional delimiter override (e.g., --delimiter "\t")
- Handles UTF-8 with BOM by default (utf-8-sig)
- Sanitizes and deduplicates sheet names to comply with Excel limits (max 31 chars, no illegal characters)
- Optional recursive directory scan
- Streams to Excel using write-only mode (memory-friendly)
- Sorted, case-insensitive processing of CSV filenames

## Requirements
- Python 3.8+
- openpyxl

Install dependency:
```bash
pip install openpyxl
```

## Usage
Save the script as csvs_to_excel.py (or any name you prefer), then run:

```bash
python csv2xlsx.py -i /path/to/csv_folder -o combined.xlsx
```

### Options
- -i, --input-dir: Directory containing CSV files (required)
- -o, --output: Output Excel path (e.g., combined.xlsx) (required)
- --encoding: CSV encoding (default: utf-8-sig)
- --delimiter: CSV delimiter override (e.g., , ; \t |). If omitted, the script auto-detects.
- -r, --recursive: Recurse into subdirectories
- -q, --quiet: Suppress progress output

## Examples

Basic combine:
```bash
python csvs_to_excel.py -i ./csvs -o combined.xlsx
```

Force tab-delimited:
```bash
python csvs_to_excel.py -i ./csvs -o combined.xlsx --delimiter "\t"
```

Recurse through subfolders:
```bash
python csvs_to_excel.py -i ./csvs -o combined.xlsx -r
```

Quiet mode:
```bash
python csvs_to_excel.py -i ./csvs -o combined.xlsx -q
```

Custom encoding (e.g., Windows-1252):
```bash
python csvs_to_excel.py -i ./csvs -o combined.xlsx --encoding cp1252
```

## Notes
- Sheet names are derived from CSV filenames (without extension) and sanitized to remove illegal characters: [ ] : * ? / \ '
- If a sanitized name collides or exceeds 31 characters, the script truncates and appends a numeric suffix (e.g., _1, _2) to ensure uniqueness.
- Auto-detection uses Python’s csv.Sniffer; if detection fails, it defaults to comma.
- Rows are written as-is (strings). Excel may auto-format known types upon open; precise typing would require a different workflow (e.g., pandas + dtype handling).

## Limitations
- Excel sheet limits apply (max ~1,048,576 rows and 16,384 columns per sheet).
- Very large CSVs are supported via streaming, but disk I/O can still be the bottleneck.
- The script writes values as plain cells; it does not preserve formulas or styles from CSV (CSV doesn’t carry styles).

## Troubleshooting
- “No CSV files found”: Verify the -i path and that files end with .csv (case-insensitive). Use -r if files are in subfolders.
- Garbled characters: Try a different --encoding (e.g., cp1252). The default utf-8-sig strips BOMs commonly found in exported CSVs.
- Wrong delimiter detected: Override with --delimiter (e.g., --delimiter ";", or --delimiter "\t").

## Contributing
Issues and pull requests are welcome. If you add features (e.g., XLSX styling, pandas integration, or ZIP ingestion), please include tests and documentation updates.

## License
idc. steal this.
