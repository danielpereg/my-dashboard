# Excel Cleaner

This repository contains `process_book.py`, a script to clean and style `book_1.xlsx`.

Requirements
- Python 3.8+
- Install dependencies:

```bash
pip install -r requirements.txt
```

Usage

```bash
python process_book.py --input book_1.xlsx --output cleaned_book_1.xlsx
```

Run the Streamlit dashboard with:

```bash
streamlit run app.py
```

The script will remove duplicates, handle missing values, standardize date columns, format phone numbers, and produce a professionally styled Excel file.
