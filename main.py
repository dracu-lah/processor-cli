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

if __name__ == "__main__":
    process_excel()
