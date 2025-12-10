#!/usr/bin/env python3
"""
Standalone Python script for converting Excel to CSV
This script will be bundled with the Electron app
"""

import pandas as pd
import sys
import os
import os.path
import json


def validate_file_path(file_path, allow_write=False):
    """
    Validate file paths to prevent path traversal attacks

    Args:
        file_path (str): The file path to validate
        allow_write (bool): Whether to allow write operations

    Returns:
        str: Normalized absolute path if valid

    Raises:
        ValueError: If path is invalid or dangerous
    """
    if not file_path:
        raise ValueError("File path cannot be empty")

    # Normalize path and resolve any .. components
    normalized_path = os.path.normpath(os.path.abspath(file_path))

    # Check for path traversal attempts
    if '..' in file_path or file_path.startswith('/'):
        if not normalized_path.startswith(os.getcwd()):
            raise ValueError("Path traversal detected: file must be within current working directory")

    # For input files, check if they exist and are readable
    if not allow_write:
        if not os.path.exists(normalized_path):
            raise ValueError(f"Input file does not exist: {normalized_path}")
        if not os.path.isfile(normalized_path):
            raise ValueError(f"Path is not a file: {normalized_path}")
        if not os.access(normalized_path, os.R_OK):
            raise ValueError(f"File is not readable: {normalized_path}")

    # For output files, check if directory exists and is writable
    if allow_write:
        output_dir = os.path.dirname(normalized_path)
        if not os.path.exists(output_dir):
            raise ValueError(f"Output directory does not exist: {output_dir}")
        if not os.access(output_dir, os.W_OK):
            raise ValueError(f"Output directory is not writable: {output_dir}")

    return normalized_path


def parse_arguments():
    """
    Safely parse and validate command line arguments

    Returns:
        dict: Validated arguments
    """
    if len(sys.argv) < 4:
        print(
            "ERROR: Invalid arguments. Usage: python_converter.py <input_file> <sheet_name> <output_file> [max_rows] [max_file_size_mb]"
        )
        sys.exit(1)

    try:
        # Validate and normalize file paths
        input_file = validate_file_path(sys.argv[1], allow_write=False)
        output_file = validate_file_path(sys.argv[3], allow_write=True)

        # Validate sheet name (basic sanitization)
        sheet_name = sys.argv[2].strip()
        if not sheet_name:
            raise ValueError("Sheet name cannot be empty")

        # Validate optional numeric arguments
        max_rows = None
        if len(sys.argv) > 4 and sys.argv[4] != "None":
            max_rows = int(sys.argv[4])
            if max_rows <= 0:
                raise ValueError("max_rows must be positive")

        max_file_size_mb = None
        if len(sys.argv) > 5 and sys.argv[5] != "None":
            max_file_size_mb = float(sys.argv[5])
            if max_file_size_mb <= 0:
                raise ValueError("max_file_size_mb must be positive")

        return {
            'input_file': input_file,
            'sheet_name': sheet_name,
            'output_file': output_file,
            'max_rows': max_rows,
            'max_file_size_mb': max_file_size_mb
        }

    except (ValueError, TypeError) as e:
        print(f"ERROR: Invalid argument - {str(e)}")
        sys.exit(1)


def main():
    args = parse_arguments()

    try:
        # Read Excel file
        print(f"Reading {args['input_file']}...")
        df = pd.read_excel(args['input_file'], sheet_name=args['sheet_name'], engine="openpyxl")
        print(f"Read {len(df)} rows")

        # Remove completely empty rows
        df = df.dropna(how="all")
        print(f"After removing empty rows: {len(df)} rows")

        if len(df) == 0:
            print("ERROR: No data found")
            sys.exit(1)

        # If no limits, create single file
        if not args['max_rows'] and not args['max_file_size_mb']:
            df.to_csv(args['output_file'], index=False, encoding="utf-8")
            print(f"Created {args['output_file']} with {len(df)} rows")
        else:
            # Need to split based on limits
            file_index = 1
            current_chunk = []
            current_chunk_rows = 0
            current_chunk_size_bytes = 0
            max_size_bytes = args['max_file_size_mb'] * 1024 * 1024 if args['max_file_size_mb'] else None
            
            # Get header size once (approximate)
            header_size_bytes = len(df.columns.to_csv(index=False).encode('utf-8')) if len(df.columns) > 0 else 0

            for idx, row in df.iterrows():
                # Convert row to CSV to estimate size
                row_csv = row.to_csv(index=False, header=False)
                row_size_bytes = len(row_csv.encode('utf-8'))
                
                # Check if adding this row would exceed limits
                would_exceed_rows = args['max_rows'] and current_chunk_rows >= args['max_rows']
                would_exceed_size = max_size_bytes and (current_chunk_size_bytes + header_size_bytes + row_size_bytes) > max_size_bytes

                # If current chunk would exceed limits, save it and start new one
                if current_chunk and (would_exceed_rows or would_exceed_size):
                    chunk_df = pd.DataFrame(current_chunk)
                    if file_index == 1:
                        output_path = args['output_file']
                    else:
                        base_name = os.path.splitext(args['output_file'])[0]
                        output_path = f"{base_name}_part{file_index}.csv"
                    
                    chunk_df.to_csv(output_path, index=False, encoding="utf-8")
                    file_size = os.path.getsize(output_path)
                    print(f"Created {output_path} with {len(chunk_df)} rows ({file_size / 1024 / 1024:.2f} MB)")
                    
                    file_index += 1
                    current_chunk = []
                    current_chunk_rows = 0
                    current_chunk_size_bytes = 0

                current_chunk.append(row)
                current_chunk_rows += 1
                current_chunk_size_bytes += row_size_bytes

            # Write remaining chunk
            if current_chunk:
                chunk_df = pd.DataFrame(current_chunk)
                if file_index == 1:
                    output_path = args['output_file']
                else:
                    base_name = os.path.splitext(args['output_file'])[0]
                    output_path = f"{base_name}_part{file_index}.csv"
                
                chunk_df.to_csv(output_path, index=False, encoding="utf-8")
                file_size = os.path.getsize(output_path)
                print(f"Created {output_path} with {len(chunk_df)} rows ({file_size / 1024 / 1024:.2f} MB)")

        print(f"SUCCESS: {len(df)} rows")
        sys.exit(0)

    except Exception as e:
        print(f"ERROR: {str(e)}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
