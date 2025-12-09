#!/usr/bin/env python3
"""
Standalone Python script for converting Excel to CSV
This script will be bundled with the Electron app
"""

import pandas as pd
import sys
import os


def main():
    if len(sys.argv) < 4:
        print(
            "ERROR: Invalid arguments. Usage: python_converter.py <input_file> <sheet_name> <output_file> [max_rows]"
        )
        sys.exit(1)

    input_file = sys.argv[1]
    sheet_name = sys.argv[2]
    output_file = sys.argv[3]
    max_rows = int(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4] != "None" else None

    try:
        # Read Excel file
        print(f"Reading {input_file}...")
        df = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl")
        print(f"Read {len(df)} rows")

        # Remove completely empty rows
        df = df.dropna(how="all")
        print(f"After removing empty rows: {len(df)} rows")

        if len(df) == 0:
            print("ERROR: No data found")
            sys.exit(1)

        # Split into multiple files if max_rows is specified
        if max_rows and len(df) > max_rows:
            num_files = (len(df) + max_rows - 1) // max_rows
            print(f"Splitting into {num_files} files...")

            for i in range(num_files):
                start_idx = i * max_rows
                end_idx = min((i + 1) * max_rows, len(df))
                chunk = df.iloc[start_idx:end_idx]

                if i == 0:
                    output_path = output_file
                else:
                    base_name = os.path.splitext(output_file)[0]
                    output_path = f"{base_name}_part{i + 1}.csv"

                chunk.to_csv(output_path, index=False, encoding="utf-8")
                print(f"Created {output_path} with {len(chunk)} rows")
        else:
            # Single file
            df.to_csv(output_file, index=False, encoding="utf-8")
            print(f"Created {output_file} with {len(df)} rows")

        print(f"SUCCESS: {len(df)} rows")
        sys.exit(0)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
