"""
Data Exploration Script - Documented Version
============================================

This script performs initial data exploration and analysis of the procurement dataset.
It provides insights into the structure, quality, and basic statistics of the transaction data.

Purpose:
- Load and examine the main transaction dataset
- Understand data structure and column types
- Analyze transaction type distribution
- Perform supplier activity analysis
- Assess data quality and identify potential issues

Key Features:
- Comprehensive data loading and validation
- Statistical analysis of transaction data
- Supplier performance overview
- Data quality assessment

Author: [Your Name]
Date: [Date]
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from IPython.display import display

# ===== DATA LOADING =====
def load_transaction_data(file_path="Vareposter Alfotech.xlsx"):
    """
    Load the main transaction dataset from Excel file.
    
    Args:
        file_path (str): Path to the Excel file containing transaction data
    
    Returns:
        pandas.DataFrame: Loaded transaction data
    """
    print("Loading transaction data...")
    sales_df = pd.read_excel(file_path, header=0)
    print(f"Data loaded successfully. Shape: {sales_df.shape}")
    return sales_df

# ===== DATA OVERVIEW =====
def explore_data_structure(df):
    """
    Explore the basic structure and information about the dataset.
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("=== DATA STRUCTURE ANALYSIS ===")
    
    # Display basic information about the dataset
    print("\n1. Dataset Info:")
    print(f"   - Total rows: {len(df)}")
    print(f"   - Total columns: {len(df.columns)}")
    print(f"   - Memory usage: {df.memory_usage(deep=True).sum() / 1024**2:.2f} MB")
    
    # Display column information
    print("\n2. Column Information:")
    df_info = df.info()
    
    # Display data types summary
    print("\n3. Data Types Summary:")
    dtype_counts = df.dtypes.value_counts()
    for dtype, count in dtype_counts.items():
        print(f"   - {dtype}: {count} columns")
    
    # Display first few rows
    print("\n4. First 5 rows of data:")
    print(df.head())
    
    return df_info

def analyze_missing_data(df):
    """
    Analyze missing data in the dataset.
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("\n=== MISSING DATA ANALYSIS ===")
    
    # Calculate missing values per column
    missing_data = df.isnull().sum()
    missing_percentage = (missing_data / len(df)) * 100
    
    missing_summary = pd.DataFrame({
        'Missing_Count': missing_data,
        'Missing_Percentage': missing_percentage
    }).sort_values('Missing_Percentage', ascending=False)
    
    print("Missing data summary:")
    print(missing_summary[missing_summary['Missing_Count'] > 0])
    
    return missing_summary

# ===== TRANSACTION TYPE ANALYSIS =====
def analyze_transaction_types(df):
    """
    Analyze the distribution of transaction types (Sales vs Purchases).
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("\n=== TRANSACTION TYPE ANALYSIS ===")
    
    # Count transaction types
    transaction_counts = df['Posttype'].value_counts()
    transaction_percentages = (transaction_counts / len(df)) * 100
    
    print("Transaction type distribution:")
    for posttype, count in transaction_counts.items():
        percentage = transaction_percentages[posttype]
        print(f"   - {posttype}: {count:,} transactions ({percentage:.1f}%)")
    
    # Create a summary DataFrame
    transaction_summary = pd.DataFrame({
        'Count': transaction_counts,
        'Percentage': transaction_percentages
    })
    
    return transaction_summary

# ===== SUPPLIER ANALYSIS =====
def analyze_supplier_activity(df):
    """
    Analyze supplier activity and performance.
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("\n=== SUPPLIER ACTIVITY ANALYSIS ===")
    
    # Count transactions per supplier
    supplier_counts = df['Leverandørnr.'].value_counts()
    
    print(f"Total unique suppliers: {len(supplier_counts)}")
    print(f"Suppliers with transactions: {len(supplier_counts[supplier_counts > 0])}")
    
    # Top suppliers by transaction count
    print("\nTop 10 suppliers by transaction count:")
    top_suppliers = supplier_counts.head(10)
    for supplier, count in top_suppliers.items():
        print(f"   - {supplier}: {count:,} transactions")
    
    # Supplier activity statistics
    print(f"\nSupplier activity statistics:")
    print(f"   - Average transactions per supplier: {supplier_counts.mean():.1f}")
    print(f"   - Median transactions per supplier: {supplier_counts.median():.1f}")
    print(f"   - Maximum transactions per supplier: {supplier_counts.max():,}")
    
    return supplier_counts

# ===== DATA QUALITY ASSESSMENT =====
def assess_data_quality(df):
    """
    Assess the overall quality of the dataset.
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("\n=== DATA QUALITY ASSESSMENT ===")
    
    # Check for duplicate rows
    duplicate_count = df.duplicated().sum()
    print(f"Duplicate rows: {duplicate_count}")
    
    # Check for empty product numbers
    empty_varenr = df['Varenr.'].isnull().sum()
    print(f"Empty product numbers: {empty_varenr}")
    
    # Check for negative quantities (which might indicate returns)
    negative_quantities = (df['Antal'] < 0).sum()
    print(f"Negative quantities (returns): {negative_quantities}")
    
    # Check for zero quantities
    zero_quantities = (df['Antal'] == 0).sum()
    print(f"Zero quantities: {zero_quantities}")
    
    # Check date range
    date_range = df['Bogføringsdato'].agg(['min', 'max'])
    print(f"Date range: {date_range['min']} to {date_range['max']}")
    
    return {
        'duplicates': duplicate_count,
        'empty_varenr': empty_varenr,
        'negative_quantities': negative_quantities,
        'zero_quantities': zero_quantities,
        'date_range': date_range
    }

# ===== STATISTICAL ANALYSIS =====
def perform_statistical_analysis(df):
    """
    Perform basic statistical analysis on numeric columns.
    
    Args:
        df (pandas.DataFrame): The transaction dataset
    """
    print("\n=== STATISTICAL ANALYSIS ===")
    
    # Select numeric columns for analysis
    numeric_columns = df.select_dtypes(include=[np.number]).columns
    print(f"Numeric columns: {list(numeric_columns)}")
    
    # Generate descriptive statistics
    print("\nDescriptive statistics:")
    stats = df[numeric_columns].describe()
    print(stats)
    
    return stats

# ===== MAIN EXECUTION =====
def main():
    """
    Main function that orchestrates the data exploration process.
    """
    print("=== DATA EXPLORATION SCRIPT ===")
    print("This script performs comprehensive analysis of the transaction dataset.\n")
    
    # Load the data
    try:
        df = load_transaction_data()
    except FileNotFoundError:
        print("Error: Transaction data file not found.")
        return
    except Exception as e:
        print(f"Error loading data: {e}")
        return
    
    # Perform data exploration
    explore_data_structure(df)
    analyze_missing_data(df)
    analyze_transaction_types(df)
    analyze_supplier_activity(df)
    assess_data_quality(df)
    perform_statistical_analysis(df)
    
    print("\n=== DATA EXPLORATION COMPLETE ===")
    print("The dataset has been analyzed and key insights have been identified.")
    print("Review the output above for data quality issues and patterns.")

if __name__ == "__main__":
    main()
