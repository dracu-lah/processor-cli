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
    print("Available columns:")
    for idx, col in enumerate(df.columns, start=1):
        print(f"{idx}. {col}")
    
    # User input for unique column selection
    selected_indices = input("Enter the column numbers to check for uniqueness (comma-separated): ")
    selected_columns = [df.columns[int(i) - 1] for i in selected_indices.split(',')]
    
    # Identify duplicated and unique rows
    duplicated_rows = df[df.duplicated(subset=selected_columns, keep=False)]
    unique_rows = df.drop(duplicated_rows.index)
    
    # Create output directories
    output_dir = os.path.join(current_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    
    # Save outputs to new Excel files in output directory
    duplicated_file = os.path.join(output_dir, "duplicated_rows.xlsx")
    unique_file = os.path.join(output_dir, "unique_rows.xlsx")
    duplicated_rows.to_excel(duplicated_file, index=False)
    unique_rows.to_excel(unique_file, index=False)
    
    print(f"Processed successfully!\nDuplicated rows saved to {duplicated_file}\nUnique rows saved to {unique_file}")

if __name__ == "__main__":
    process_excel()
