#!/usr/bin/env python3
"""
process_book.py

Clean and style an Excel file (book_1.xlsx) using pandas and xlsxwriter.

Features:
- Remove duplicate rows
- Handle missing values sensibly
- Parse and standardize date columns
- Normalize and format phone numbers
- Write styled Excel output with a corporate look
- Error handling and logging

Usage:
    python process_book.py --input book_1.xlsx --output cleaned_book_1.xlsx
"""
from __future__ import annotations

import argparse
import logging
import os
import re
from typing import Optional

import pandas as pd


logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names: strip, lower, replace spaces with underscores."""
    df = df.copy()
    df.columns = [re.sub(r"\s+", "_", c.strip()).lower() for c in df.columns]
    return df


def format_phone_number(value: object) -> Optional[str]:
    """Return a formatted phone number or None.

    Attempts to extract digits and format as +1 (AAA) BBB-CCCC when possible.
    If digits length is 10, assume US number. If 11 and starts with '1', keep country code.
    Otherwise return the digits-only string.
    """
    if pd.isna(value):
        return None
    s = str(value)
    digits = re.sub(r"\D", "", s)
    if not digits:
        return None
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"+1 ({digits[0:3]}) {digits[3:6]}-{digits[6:]}"
    # fallback: return digits
    return digits


def parse_date_series(s: pd.Series) -> pd.Series:
    """Try to parse a pandas Series to datetimes; coerce errors to NaT."""
    try:
        return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    except Exception:
        return pd.to_datetime(s, errors="coerce")


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Perform cleaning steps on the DataFrame and return a cleaned copy."""
    if df.empty:
        return df

    df = normalize_column_names(df)

    # Drop exact duplicate rows
    before = len(df)
    df = df.drop_duplicates()
    logger.info("Dropped %d duplicate rows", before - len(df))

    # Handle missing values: object->empty string, numeric->0
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("")

    # Parse date-like columns (heuristic: column name contains 'date')
    for col in df.columns:
        if "date" in col:
            parsed = parse_date_series(df[col])
            num_parsed = parsed.notna().sum()
            logger.info("Parsed %d values in date column '%s'", int(num_parsed), col)
            df[col] = parsed

    # Format phone columns (heuristic: column name contains 'phone' or 'tel')
    for col in df.columns:
        if any(k in col for k in ("phone", "tel", "mobile", "contact")):
            df[col] = df[col].apply(format_phone_number)

    return df


def style_and_write(df: pd.DataFrame, output_path: str) -> None:
    """Write the DataFrame to Excel with styling using xlsxwriter."""
    try:
        with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
            df.to_excel(writer, sheet_name="Cleaned", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Cleaned"]

            # Formats
            header_fmt = workbook.add_format({
                "bold": True,
                "bg_color": "#2F5597",
                "font_color": "#FFFFFF",
                "border": 1,
            })
            cell_fmt = workbook.add_format({"border": 1})
            date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd", "border": 1})
            odd_row = workbook.add_format({"bg_color": "#F7F7F7"})

            # Apply header format
            for col_num, value in enumerate(df.columns):
                worksheet.write(0, col_num, value, header_fmt)

            # Set column widths
            for i, col in enumerate(df.columns):
                series = df[col].astype(str).where(df[col].notna(), "")
                max_len = max(series.map(len).max(), len(col)) + 2
                max_len = min(max_len, 50)
                worksheet.set_column(i, i, max_len, cell_fmt)

            # Apply date format to date columns
            for i, col in enumerate(df.columns):
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    worksheet.set_column(i, i, None, date_fmt)

            # Freeze header row and add autofilter
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

            # Apply alternating row banding for readability
            for row in range(1, len(df) + 1):
                if row % 2 == 0:
                    worksheet.set_row(row, None, odd_row)

            logger.info("Wrote styled Excel to %s", output_path)
    except Exception as exc:
        logger.exception("Failed to write styled Excel: %s", exc)
        raise


def read_input(input_path: str) -> pd.DataFrame:
    """Read input Excel file into a DataFrame with basic checks."""
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
    try:
        df = pd.read_excel(input_path, engine="openpyxl")
        logger.info("Read %d rows and %d columns from %s", len(df), len(df.columns), input_path)
        return df
    except Exception as exc:
        logger.exception("Error reading Excel file: %s", exc)
        raise


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Clean and style an Excel workbook")
    parser.add_argument("--input", "-i", default="book_1.xlsx", help="Input Excel file")
    parser.add_argument("--output", "-o", default="cleaned_book_1.xlsx", help="Output Excel file")
    args = parser.parse_args(argv)

    try:
        df = read_input(args.input)
        cleaned = clean_dataframe(df)
        style_and_write(cleaned, args.output)
        logger.info("Processing completed successfully")
        return 0
    except Exception as exc:
        logger.error("Processing failed: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
