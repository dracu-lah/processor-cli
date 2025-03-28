import pandas as pd
import os

def process_excel():
    # Get the current directory
    current_dir = os.getcwd()
    
    # Find the first Excel file in the directory
    excel_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
    if not excel_files:
        print("No Excel files found in the current directory.")
        return
    
    file_path = os.path.join(current_dir, excel_files[0])
    print(f"Processing file: {file_path}")
    
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Display column names for selection
    print("\nAvailable columns:")
    for idx, col in enumerate(df.columns, start=1):
        print(f"{idx}. {col}")
    
    # User input for unique column selection
    use_whole_row = input("Do you want to check uniqueness based on the whole row? (yes/no): ").strip().lower()
    
    if use_whole_row == "yes":
        selected_columns = df.columns.tolist()
    else:
        selected_indices = input("Enter the column numbers to check for uniqueness (comma-separated): ")
        selected_columns = [df.columns[int(i) - 1] for i in selected_indices.split(',')]
    
    # Ask user whether to use "AND" or "OR" for multiple columns
    if len(selected_columns) > 1:
        operation = input('Check for uniqueness using "AND" (both must match) or "OR" (either can match)? (and/or): ').strip().lower()
    else:
        operation = "and"  # Default to "AND" when only one column is selected
    
    # Identify duplicated and unique rows
    if operation == "and":
        duplicated_rows = df[df.duplicated(subset=selected_columns, keep=False)]
    else:  # OR operation: check duplicates for each column separately and combine the results
        duplicated_rows = pd.DataFrame()
        for col in selected_columns:
            dupes = df[df.duplicated(subset=[col], keep=False)]
            duplicated_rows = pd.concat([duplicated_rows, dupes]).drop_duplicates()

    unique_rows = df.drop(duplicated_rows.index)

    # Sorting Feature
    sort_choice = input("Do you want to sort the data? (yes/no): ").strip().lower()
    if sort_choice == "yes":
        sort_column_index = input("Enter the column number to sort by: ").strip()
        sort_order = input("Enter sorting order (asc/desc): ").strip().lower()
        
        try:
            sort_column = df.columns[int(sort_column_index) - 1]
            ascending = True if sort_order == "asc" else False
            duplicated_rows = duplicated_rows.sort_values(by=sort_column, ascending=ascending)
            unique_rows = unique_rows.sort_values(by=sort_column, ascending=ascending)
        except (IndexError, ValueError):
            print("Invalid column selection for sorting. Skipping sorting.")

    # Create output directories
    output_dir = os.path.join(current_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # Save outputs to new Excel files in output directory
    duplicated_file = os.path.join(output_dir, "duplicated_rows.xlsx")
    unique_file = os.path.join(output_dir, "unique_rows.xlsx")
    duplicated_rows.to_excel(duplicated_file, index=False)
    unique_rows.to_excel(unique_file, index=False)
    
    print(f"\nProcessed successfully!")
    print(f"Duplicated rows saved to: {duplicated_file}")
    print(f"Unique rows saved to: {unique_file}")

    # Splitting Feature
    split_choice = input("Do you want to split the unique rows into smaller files? (yes/no): ").strip().lower()
    if split_choice == "yes":
        try:
            num_rows_per_file = int(input("Enter the number of rows per file: ").strip())
            if num_rows_per_file > 0:
                split_dir = os.path.join(output_dir, "split_files")
                os.makedirs(split_dir, exist_ok=True)
                
                num_splits = (len(unique_rows) // num_rows_per_file) + (1 if len(unique_rows) % num_rows_per_file != 0 else 0)
                
                for i in range(num_splits):
                    start_idx = i * num_rows_per_file
                    end_idx = start_idx + num_rows_per_file
                    split_df = unique_rows.iloc[start_idx:end_idx]
                    split_file = os.path.join(split_dir, f"split_{i+1}.xlsx")
                    split_df.to_excel(split_file, index=False)
                
                print(f"Unique rows split into {num_splits} files inside: {split_dir}")
            else:
                print("Invalid number of rows. Skipping splitting.")
        except ValueError:
            print("Invalid input for row count. Skipping splitting.")

if __name__ == "__main__":
    process_excel()
