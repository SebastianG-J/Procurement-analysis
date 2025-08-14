"""
Data Merger Script - Documented Version
=======================================

This script merges two Excel files containing analysis results by appending
the data from one file to another. It's used to combine different batches
of scraped or analyzed data into a single comprehensive dataset.

Purpose:
- Combine results from different analysis runs
- Merge scraped data from multiple sessions
- Create consolidated datasets for further analysis

Author: [Your Name]
Date: [Date]
"""

import pandas as pd

def merge_excel_files(file1_path, file2_path, output_path):
    """
    Merge two Excel files by concatenating their data.
    
    Args:
        file1_path (str): Path to the first Excel file (base data)
        file2_path (str): Path to the second Excel file (data to append)
        output_path (str): Path where the merged file will be saved
    
    Returns:
        pandas.DataFrame: The merged dataset
    """
    # Load both Excel files into pandas DataFrames
    # This reads the data from the specified Excel files
    df_simplified = pd.read_excel(file1_path)
    df_new = pd.read_excel(file2_path)
    
    # Append new data to the bottom of the existing data
    # ignore_index=True ensures that the index is reset to 0, 1, 2, etc.
    # This prevents duplicate index values that could cause issues
    df_merged = pd.concat([df_simplified, df_new], ignore_index=True)
    
    # Save the merged data to a new Excel file
    # index=False prevents the DataFrame index from being saved as a column
    df_merged.to_excel(output_path, index=False)
    
    return df_merged

def main():
    """
    Main function that executes the merge operation with default file paths.
    """
    # Define the input and output file paths
    # These are the default file names used in the original script
    input_file1 = "results_merged.xlsx"  # Base dataset
    input_file2 = "results_new.xlsx"     # New data to append
    output_file = "results_merged1.xlsx" # Output file name
    
    # Execute the merge operation
    merged_data = merge_excel_files(input_file1, input_file2, output_file)
    
    # Note: Print statement removed to avoid exposing confidential information
    # The merge operation is complete when the function returns without errors

if __name__ == "__main__":
    # Execute the main function when the script is run directly
    main()
