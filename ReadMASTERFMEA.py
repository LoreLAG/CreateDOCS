import pandas as pd
import os


def format_masters_db(file_path):
    """Formats the Excel file, removes duplicates, and smartly unmerges cells."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist!")

    # 1. Read the 'MASTERS' sheet
    try:
        df = pd.read_excel(file_path, sheet_name='MASTERS')
    except ValueError:
        df = pd.read_excel(file_path, sheet_name=0)

    df.dropna(how='all', inplace=True)

    if df.empty:
        return 0

    phases_column = df.columns[0]

    # 2. SMART UNMERGE OF COMBINED CELLS
    # A. First, fill empty cells only in the first column (Phase Name)
    df[phases_column] = df[phases_column].ffill()

    # B. Then, fill the rest of the columns, GROUPED BY PHASE!
    # This way, a new row (empty except for the name) will NEVER inherit the defects of the phase above.
    data_columns = df.columns[1:]
    df[data_columns] = df.groupby(phases_column, sort=False)[data_columns].ffill()

    # 3. TEXT FORMATTING (Capitalize first letter only)
    def format_text(value):
        if isinstance(value, str) and value.strip():
            return value.strip().capitalize()
        return value

    for col in df.columns:
        df[col] = df[col].apply(format_text)

    df.columns = [format_text(col) for col in df.columns]

    # 4. REMOVAL OF EXACT DUPLICATES
    initial_rows = len(df)
    df.drop_duplicates(inplace=True)
    removed_duplicates = initial_rows - len(df)

    # 5. SAVE IN THE "functional_database" SHEET
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='functional_database', index=False)

    return removed_duplicates


if __name__ == "__main__":
    file_master = "MASTERS-FMEA.xlsx"
    try:
        print(f"🔄 Starting cleaning and formatting of {file_master}...")
        removed = format_masters_db(file_master)
        print(f"✅ Operation completed! Removed {removed} duplicates.")
    except Exception as e:
        print(f"❌ Error: {e}")
