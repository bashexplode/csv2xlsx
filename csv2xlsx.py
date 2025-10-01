#!/usr/bin/env python3
import os
import argparse
import csv
from openpyxl import Workbook

ILLEGAL_SHEET_CHARS = set(['[', ']', ':', '*', '?', '/', '\\', "'"])
MAX_SHEET_LEN = 31

def sanitize_sheet_name(name, used_names):
    # Remove extension and illegal characters
    name = os.path.splitext(name)[0]
    name = "".join("_" if ch in ILLEGAL_SHEET_CHARS else ch for ch in name)
    name = name.strip() or "Sheet"

    # Truncate to Excel's limit
    base = name[:MAX_SHEET_LEN]

    # Ensure uniqueness
    candidate = base
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        max_base_len = MAX_SHEET_LEN - len(suffix)
        candidate = f"{base[:max_base_len]}{suffix}"
        counter += 1

    used_names.add(candidate)
    return candidate

def sniff_dialect(sample_bytes, provided_delimiter=None):
    # Fallback to comma if sniffer fails
    if provided_delimiter:
        class SimpleDialect(csv.excel):
            delimiter = provided_delimiter
        return SimpleDialect()
    try:
        sample_text = sample_bytes.decode("utf-8", errors="ignore")
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample_text, delimiters=[",", ";", "\t", "|"])
        return dialect
    except Exception:
        return csv.excel  # default comma

def read_csv_rows(path, encoding="utf-8-sig", delimiter=None):
    # Try to detect delimiter from a small sample if not provided
    with open(path, "rb") as fb:
        sample = fb.read(4096)
    dialect = sniff_dialect(sample, provided_delimiter=delimiter)

    with open(path, "r", newline="", encoding=encoding, errors="replace") as f:
        reader = csv.reader(f, dialect)
        for row in reader:
            yield row

def find_csv_files(input_dir, recursive=False):
    files = []
    if recursive:
        for root, _, filenames in os.walk(input_dir):
            for fn in filenames:
                if fn.lower().endswith(".csv"):
                    files.append(os.path.join(root, fn))
    else:
        for fn in os.listdir(input_dir):
            if fn.lower().endswith(".csv"):
                files.append(os.path.join(input_dir, fn))
    # Sort case-insensitive by filename
    return sorted(files, key=lambda p: os.path.basename(p).lower())

def combine_csvs_to_excel(input_dir, output_path, encoding="utf-8-sig", delimiter=None, recursive=False, verbose=True):
    csv_files = find_csv_files(input_dir, recursive=recursive)
    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in: {input_dir}")

    wb = Workbook(write_only=True)
    used_names = set()

    for csv_path in csv_files:
        sheet_name = sanitize_sheet_name(os.path.basename(csv_path), used_names)
        ws = wb.create_sheet(title=sheet_name)

        if verbose:
            print(f"Adding sheet: {sheet_name} from {csv_path}")

        try:
            for row in read_csv_rows(csv_path, encoding=encoding, delimiter=delimiter):
                ws.append(row)
        except Exception as e:
            if verbose:
                print(f"Warning: failed to process {csv_path}: {e}")

    # Ensure output directory exists
    outdir = os.path.dirname(os.path.abspath(output_path)) or "."
    os.makedirs(outdir, exist_ok=True)
    wb.save(output_path)
    if verbose:
        print(f"Wrote Excel workbook: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Combine a folder of CSVs into a single Excel workbook, one sheet per CSV.")
    parser.add_argument("-i", "--input-dir", required=True, help="Directory containing CSV files")
    parser.add_argument("-o", "--output", required=True, help="Output Excel path (e.g., combined.xlsx)")
    parser.add_argument("--encoding", default="utf-8-sig", help="CSV encoding (default utf-8-sig)")
    parser.add_argument("--delimiter", help="CSV delimiter override (e.g., , ; \\t |). If omitted, auto-detect.")
    parser.add_argument("-r", "--recursive", action="store_true", help="Recurse into subdirectories")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress progress output")
    args = parser.parse_args()

    if not os.path.isdir(args.input_dir):
        raise NotADirectoryError(f"Not a directory: {args.input_dir}")

    delimiter = None
    if args.delimiter:
        delimiter = args.delimiter
        if delimiter.lower() == "\\t":
            delimiter = "\t"

    combine_csvs_to_excel(
        input_dir=args.input_dir,
        output_path=args.output,
        encoding=args.encoding,
        delimiter=delimiter,
        recursive=args.recursive,
        verbose=not args.quiet,
    )

if __name__ == "__main__":
    main()