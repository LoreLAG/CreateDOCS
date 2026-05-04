import pandas as pd
import os


def format_master_cp_db(file_path):
    """
    Formats the Control Plan Excel file, merging the two header rows,
    removes duplicates, formats the text, and unmerges combined cells.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist!")

    # 1. Ignore the first 8 rows of company headers and use Excel rows 9 and 10 (indices 8 and 9) as column names
    df = pd.read_excel(file_path, sheet_name=0, header=[8, 9])

    # REMOVE ghost columns and rows (completely empty) to prevent crashes
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(axis=0, how='all', inplace=True)

    # Flattens the two-row header into a single string (e.g., "Phase_Machine")
    # Ignoring empty columns and DEDUPLICATING identical names
    new_columns = []
    seen = {}
    for i, col in enumerate(df.columns.values):
        name = ' '.join([str(c) for c in col if "Unnamed" not in str(c) and pd.notna(c)]).strip()

        # If the resulting name is empty, assign a dummy name
        if not name:
            name = f"Unnamed_Column_{i}"

        # Anti-Crash System for columns with the exact same name
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1

        new_columns.append(name)

    df.columns = new_columns

    if df.empty:
        return 0

    phases_column = df.columns[0]
    data_columns = df.columns[1:]

    # 2. SMART UNMERGE OF COMBINED CELLS
    # A. Fill empty cells of the Phase (column 0)
    df[phases_column] = df[phases_column].ffill()

    # B. Fill the rest of the columns, grouped by phase (executed column by column to avoid misalignment)
    for col in data_columns:
        df[col] = df.groupby(phases_column, sort=False)[col].ffill()

    # 3. TEXT FORMATTING
    def format_text_uppercase(value):
        """Forces everything to UPPERCASE (used for the first column)."""
        if isinstance(value, str) and value.strip():
            return value.strip().upper()
        return value

    def format_text_normal(value):
        """Capitalizes the first letter, leaves the rest intact (to preserve acronyms like SPC, CMM)."""
        if isinstance(value, str) and value.strip():
            text = value.strip()
            return text[0].upper() + text[1:]
        return value

    # Apply ALL UPPERCASE to the first column
    df[phases_column] = df[phases_column].apply(format_text_uppercase)

    # Apply "first letter uppercase, preserve acronyms" to the rest of the columns
    for col in data_columns:
        df[col] = df[col].apply(format_text_normal)

    # Apply "normal" formatting to the newly flattened column titles as well
    df.columns = [format_text_normal(col) for col in df.columns]

    # 4. REMOVAL OF EXACT DUPLICATES
    initial_rows = len(df)
    df.drop_duplicates(inplace=True)
    removed_duplicates = initial_rows - len(df)

    # 5. SAVE IN THE "functional_database" SHEET
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='functional_database', index=False)

    return removed_duplicates


if __name__ == "__main__":
    file_master_cp = "MASTERS-CP.xlsx"

    try:
        print(f"🔄 Starting cleaning and formatting of {file_master_cp}...")
        removed = format_master_cp_db(file_master_cp)
        print(f"✅ Operation completed! Removed {removed} duplicates.")
        print(f"📁 'functional_database' sheet successfully generated/updated.")
    except Exception as e:
        print(f"❌ Error: {e}")
