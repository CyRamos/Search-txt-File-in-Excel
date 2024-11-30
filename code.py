import pandas as pd
from pathlib import Path


def search_excel_workbook(excel_file, search_term):
    """
    Search for a term in all sheets of an Excel workbook using pandas.
    Returns a list of tuples containing (sheet_name, row, column) where term was found.
    """
    # Read all sheets into a dictionary of dataframes
    all_sheets = pd.read_excel(excel_file, sheet_name=None)
    results = []

    for sheet_name, df in all_sheets.items():
        # Convert all values to strings and replace NaN with empty string
        df_str = df.astype(str).replace('nan', '')

        # Search in each column
        for col in df_str.columns:
            # Find matching rows (using contains instead of exact match)
            matches = df_str[df_str[col].str.lower().str.contains(search_term.lower(), na=False)]

            if not matches.empty:
                # For each matching row
                for idx in matches.index:
                    # Add 1 to row and column numbers to match Excel's 1-based indexing
                    results.append((sheet_name, idx + 1, df_str.columns.get_loc(col) + 1, matches.at[idx, col]))

    return results


def main():
    # Get file paths from user
    txt_file = input("Enter the path to your text file: ")
    excel_file = input("Enter the path to your Excel file: ")

    # Read search terms from text file
    try:
        with open(txt_file, 'r', encoding='utf-8') as file:
            search_terms = [line.strip() for line in file if line.strip()]
    except FileNotFoundError:
        print(f"Error: Text file '{txt_file}' not found.")
        return

    # Check if Excel file exists
    if not Path(excel_file).exists():
        print(f"Error: Excel file '{excel_file}' not found.")
        return

    # Search for each term
    print("\nResults:")
    for term in search_terms:
        results = search_excel_workbook(excel_file, term)
        if results:
            print(f"\n{term}:")
            for sheet_name, row, col, value in results:
                print(f"Found in Sheet: {sheet_name}, Row: {row}, Column: {col}")


if __name__ == "__main__":
    main()