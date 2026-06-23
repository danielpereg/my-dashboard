#!/usr/bin/env python3
"""
create_dashboard.py

Read raw Excel data, clean it with pandas, and produce a polished
business dashboard Excel file (`Client_Dashboard.xlsx`) using xlsxwriter.

Key steps implemented:
- Data cleaning: remove 'unnamed' cols, drop rows with missing values, remove duplicates
- Dashboard styling: corporate navy headers, merged title, hidden gridlines,
  alternating row stripes, conditional formatting for `purchase_amount`, auto-fit columns,
  and 120% zoom for presentation

Usage:
    python create_dashboard.py --input book_1.xlsx --output Client_Dashboard.xlsx
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


def read_excel(path: str) -> pd.DataFrame:
    """Read an Excel file into a DataFrame using openpyxl engine."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    try:
        df = pd.read_excel(path, engine="openpyxl")
        logger.info("Loaded %d rows and %d cols from %s", len(df), len(df.columns), path)
        return df
    except Exception:
        logger.exception("Failed to read Excel file: %s", path)
        raise


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Clean DataFrame according to business rules.

    - Remove columns whose name contains 'unnamed' (case-insensitive)
    - Drop rows with any missing values
    - Drop duplicate rows
    - Normalize column names (strip and lower)
    """
    df = df.copy()

    # Normalize column names
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    # Drop columns containing 'unnamed'
    unnamed_cols = [c for c in df.columns if isinstance(c, str) and re.search(r"unnamed", c, re.I)]
    if unnamed_cols:
        df = df.drop(columns=unnamed_cols)
        logger.info("Dropped unnamed columns: %s", unnamed_cols)

    # Drop rows with any missing values (user requested)
    before = len(df)
    df = df.dropna()
    logger.info("Dropped %d rows with missing values", before - len(df))

    # Drop exact duplicates
    before = len(df)
    df = df.drop_duplicates()
    logger.info("Dropped %d duplicate rows", before - len(df))

    # Normalize column names to safe identifiers for downstream use
    df.columns = [re.sub(r"\s+", "_", str(c).strip()).lower() for c in df.columns]

    return df


def autofit_columns(worksheet, df: pd.DataFrame, startrow: int, formats: dict):
    """Adjust column widths to fit the longest cell in each column.

    - `startrow` is the row index where the dataframe header was written.
    - `formats` is a dict mapping column index -> cell format (optional)
    """
    for i, col in enumerate(df.columns):
        series = df[col].astype(str).where(df[col].notna(), "")
        max_len = max(series.map(len).max(), len(str(col))) + 2
        max_len = min(max_len, 60)
        fmt = formats.get(i) if formats else None
        worksheet.set_column(i, i, max_len, fmt)


def create_dashboard_excel(df: pd.DataFrame, output_path: str) -> None:
    """Write a polished dashboard Excel file using xlsxwriter styling."""
    try:
        with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
            # Write dataframe starting at row 2 so we can place a merged title above
            startrow = 2
            df.to_excel(writer, sheet_name="Dashboard", index=False, startrow=startrow)

            workbook = writer.book
            worksheet = writer.sheets["Dashboard"]

            # --- Define formats ---
            title_fmt = workbook.add_format({
                "bold": True,
                "font_size": 16,
                "align": "center",
                "valign": "vcenter",
                "font_color": "#FFFFFF",
                "bg_color": "#2F5597",
            })

            header_fmt = workbook.add_format({
                "bold": True,
                "bg_color": "#2F5597",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            })

            cell_fmt = workbook.add_format({"border": 1, "align": "left"})
            stripe_fmt = workbook.add_format({"bg_color": "#F5F5F5"})

            # Light green for conditional formatting
            highlight_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})

            # --- Title: merged at very top ---
            last_col = len(df.columns) - 1 if len(df.columns) else 0
            worksheet.merge_range(0, 0, 0, last_col, "Client Performance Dashboard", title_fmt)

            # Hide gridlines for a clean, modern look
            worksheet.hide_gridlines(2)

            # Write header formats (headers were written by pandas at startrow)
            for col_num, value in enumerate(df.columns):
                worksheet.write(startrow, col_num, value, header_fmt)

            # Apply alternating row stripes for readability
            for row in range(startrow + 1, startrow + 1 + len(df)):
                if (row - startrow) % 2 == 0:
                    worksheet.set_row(row, None, stripe_fmt)

            # Conditional formatting for 'purchase_amount' to highlight high values
            col_name = "purchase_amount"
            if col_name in df.columns:
                # define threshold as 90th percentile (top 10%)
                try:
                    threshold = float(df[col_name].quantile(0.9))
                except Exception:
                    threshold = None

                if threshold is not None:
                    col_idx = df.columns.get_loc(col_name)
                    first_row = startrow + 1
                    last_row = startrow + len(df)
                    cell_range = (first_row, col_idx, last_row, col_idx)
                    worksheet.conditional_format(first_row, col_idx, last_row - 1, col_idx, {
                        'type': 'cell',
                        'criteria': '>=',
                        'value': threshold,
                        'format': highlight_fmt,
                    })
                    logger.info("Applied conditional formatting on '%s' for values >= %s", col_name, threshold)

            # Auto-fit columns and set zoom for presentation
            autofit_columns(worksheet, df, startrow, formats={})
            worksheet.set_zoom(120)

            # Freeze panes to keep headers and title visible
            worksheet.freeze_panes(startrow + 1, 0)

            logger.info("Saved dashboard to %s", output_path)
    except Exception:
        logger.exception("Failed to create dashboard file: %s", output_path)
        raise


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Create Client Dashboard Excel file")
    parser.add_argument("--input", "-i", default="book_1.xlsx", help="Input Excel file path")
    parser.add_argument("--output", "-o", default="Client_Dashboard.xlsx", help="Output Excel file path")
    args = parser.parse_args(argv)

    try:
        raw = read_excel(args.input)
        cleaned = clean_data(raw)
        create_dashboard_excel(cleaned, args.output)
        logger.info("Dashboard generation completed successfully")
        return 0
    except Exception as exc:
        logger.error("Dashboard generation failed: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
