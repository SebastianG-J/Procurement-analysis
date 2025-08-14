"""
Overall Analysis Script - Documented Version
============================================

This script performs comprehensive analysis of the procurement dataset, including
product analysis, supplier performance metrics, data enrichment, and quality assessment.

Purpose:
- Analyze product sales and purchase patterns
- Evaluate supplier performance and relationships
- Identify data quality issues and inconsistencies
- Enrich product data with additional information
- Generate comprehensive reports for business insights

Key Features:
- Product identification and categorization
- Sales vs purchase volume analysis
- Multi-supplier product analysis
- Data quality assessment and cleaning
- Product description matching and enrichment

Author: [Your Name]
Date: [Date]
"""

import pandas as pd
import numpy as np
import openpyxl
from pathlib import Path

# ===== DATA LOADING AND PREPARATION =====
def load_main_datasets():
    """
    Load the main datasets for analysis.
    
    Returns:
        tuple: (vareposter_df, varer_df) - Main transaction data and product master data
    """
    print("Loading main datasets...")
    
    # Load main transaction data
    vareposter = pd.read_excel('Vareposter Alfotech.xlsx', header=0)
    print(f"Transaction data loaded: {vareposter.shape}")
    
    # Load product master data
    varer = pd.read_excel('Varer.xlsx', header=0)
    print(f"Product master data loaded: {varer.shape}")
    
    return vareposter, varer

def extract_unique_products(vareposter_df):
    """
    Extract unique product numbers from the transaction data.
    
    Args:
        vareposter_df (pandas.DataFrame): Main transaction dataset
    
    Returns:
        pandas.DataFrame: DataFrame with unique product numbers
    """
    print("Extracting unique product numbers...")
    
    # Extract unique product numbers
    unikke_varenumre = vareposter_df['Varenr.'].unique()
    unikke_varenumre = pd.DataFrame(unikke_varenumre, columns=['Varenr.'])
    
    print(f"Found {len(unikke_varenumre)} unique products")
    
    # Save to Excel for reference
    unikke_varenumre.to_excel('unikke_varenumre.xlsx', index=False)
    print("Unique products saved to 'unikke_varenumre.xlsx'")
    
    return unikke_varenumre

def separate_transaction_types(vareposter_df):
    """
    Separate sales and purchase transactions.
    
    Args:
        vareposter_df (pandas.DataFrame): Main transaction dataset
    
    Returns:
        tuple: (sales_df, purchases_df) - Separated sales and purchase data
    """
    print("Separating transaction types...")
    
    # Separate sales transactions (including Montageforbrug)
    vareposter_salg = vareposter_df[vareposter_df['Posttype'].isin(['Salg', 'Montageforbrug'])]
    
    # Separate purchase transactions
    vareposter_køb = vareposter_df[vareposter_df['Posttype'] == 'Køb']
    
    print(f"Sales transactions: {len(vareposter_salg)}")
    print(f"Purchase transactions: {len(vareposter_køb)}")
    
    # Save separated data
    vareposter_salg.to_excel('vareposter_salg.xlsx', index=False)
    vareposter_køb.to_excel('vareposter_køb.xlsx', index=False)
    
    return vareposter_salg, vareposter_køb

# ===== PRODUCT ANALYSIS =====
def analyze_sales_by_product_supplier(sales_df):
    """
    Analyze sales quantities by product and supplier.
    
    Args:
        sales_df (pandas.DataFrame): Sales transaction data
    
    Returns:
        pandas.DataFrame: Sales analysis by product and supplier
    """
    print("Analyzing sales by product and supplier...")
    
    # Calculate total quantities sold per product and supplier
    antal_per_varenr_leverandor = (
        sales_df.groupby(['Varenr.', 'Leverandørnr.'])['Antal']
        .sum()
        .reset_index()
        .sort_values('Antal', ascending=True)  # Most sold products at top (negative values)
    )
    
    # Save analysis results
    antal_per_varenr_leverandor.to_excel('antal_per_varenr_leverandor.xlsx', index=False)
    print("Sales analysis saved to 'antal_per_varenr_leverandor.xlsx'")
    
    return antal_per_varenr_leverandor

def analyze_purchases_by_product_supplier(purchases_df):
    """
    Analyze purchase quantities by product and supplier.
    
    Args:
        purchases_df (pandas.DataFrame): Purchase transaction data
    
    Returns:
        pandas.DataFrame: Purchase analysis by product and supplier
    """
    print("Analyzing purchases by product and supplier...")
    
    # Calculate total quantities purchased per product and supplier
    antal_per_varenr_leverandor_purchased = (
        purchases_df.groupby(['Varenr.', 'Leverandørnr.'])['Antal']
        .sum()
        .reset_index()
        .sort_values('Antal', ascending=False)  # Most purchased products at top
    )
    
    # Save analysis results
    antal_per_varenr_leverandor_purchased.to_excel('antal_per_varenr_leverandor_purchased.xlsx', index=False)
    print("Purchase analysis saved to 'antal_per_varenr_leverandor_purchased.xlsx'")
    
    return antal_per_varenr_leverandor_purchased

def analyze_products_without_supplier(sales_df):
    """
    Analyze products sold without supplier information.
    
    Args:
        sales_df (pandas.DataFrame): Sales transaction data
    
    Returns:
        pandas.DataFrame: Products sold without supplier information
    """
    print("Analyzing products sold without supplier information...")
    
    # Find products sold without supplier information
    antal_uden_leverandør = (
        sales_df.loc[sales_df["Leverandørnr."].isnull()]
            .groupby("Varenr.", as_index=False)["Antal"]
            .sum()
            .sort_values("Antal", ascending=True)
            .reset_index(drop=True)
    )
    
    # Save analysis results
    antal_uden_leverandør.to_excel('antal_uden_leverandør.xlsx', index=False)
    print("Products without supplier analysis saved to 'antal_uden_leverandør.xlsx'")
    
    return antal_uden_leverandør

# ===== MULTI-SUPPLIER ANALYSIS =====
def analyze_shared_products_sales(sales_df):
    """
    Analyze products sold by multiple suppliers.
    
    Args:
        sales_df (pandas.DataFrame): Sales transaction data
    
    Returns:
        pandas.DataFrame: Analysis of products sold by multiple suppliers
    """
    print("Analyzing products sold by multiple suppliers...")
    
    # Normalize product numbers to avoid hidden duplicates
    df = sales_df.copy()
    df["Varenr."] = df["Varenr."].astype(str).str.strip().str.upper()
    df["Leverandørnr."] = df["Leverandørnr."].astype(str).str.strip()
    
    # Keep only actual sales (negative quantities)
    df_sales = df[df["Antal"] < 0].copy()
    
    # Count unique suppliers per product
    supplier_counts = (
        df_sales.groupby("Varenr.")["Leverandørnr."]
        .nunique()
        .reset_index(name="Antal leverandører")
    )
    
    # Filter to products sold by more than one supplier
    shared_products = supplier_counts[supplier_counts["Antal leverandører"] > 1]["Varenr."]
    
    if shared_products.empty:
        print("No products found with more than one supplier in sales.")
        return pd.DataFrame()
    
    # Restrict analysis to shared products
    df_filtered = df_sales[df_sales["Varenr."].isin(shared_products)].copy()
    
    # Sum sold per product and supplier
    summary = (
        df_filtered
        .groupby(["Varenr.", "Leverandørnr."])["Antal"]
        .sum()
        .reset_index()
        .rename(columns={"Antal": "Total_Antal"})
    )
    
    # Convert to positive sold quantity
    summary["Antal_solgt"] = -summary["Total_Antal"]
    
    # Merge back supplier counts
    summary = summary.merge(supplier_counts, on="Varenr.", how="left")
    
    # Compute share of total sold for each supplier
    total_per_product = summary.groupby("Varenr.")["Antal_solgt"].transform("sum")
    summary["Andel_pct"] = (summary["Antal_solgt"] / total_per_product * 100).round(1)
    
    # Sort for readability
    summary = summary.sort_values(["Varenr.", "Antal_solgt"], ascending=[True, False])
    
    # Save results
    summary.to_excel("salg_med_flere_leverandoerer.xlsx", index=False)
    print("Multi-supplier analysis saved to 'salg_med_flere_leverandoerer.xlsx'")
    
    return summary

# ===== DATA QUALITY ANALYSIS =====
def analyze_supplier_number_mismatches(purchases_df):
    """
    Analyze cases where Kildenr. and Leverandørnr. don't match.
    
    Args:
        purchases_df (pandas.DataFrame): Purchase transaction data
    
    Returns:
        pandas.DataFrame: Analysis of supplier number mismatches
    """
    print("Analyzing supplier number mismatches...")
    
    # Find cases where Kildenr. and Leverandørnr. don't match
    not_same_kildenr_leverandørnr = purchases_df[purchases_df['Kildenr.'] != purchases_df['Leverandørnr.']]
    
    print(f"Found {len(not_same_kildenr_leverandørnr)} transactions with mismatched supplier numbers")
    
    # Save analysis results
    not_same_kildenr_leverandørnr.to_excel('not_same_kildenr_leverandørnr.xlsx', index=False)
    
    # Analyze by Kildenr.
    df = not_same_kildenr_leverandørnr.copy()
    df['Antal'] = pd.to_numeric(df['Antal'], errors='coerce').fillna(0)
    
    # Group by Kildenr and sum quantities
    sum_by_kildenr = (
        df.groupby('Kildenr.', dropna=False, observed=False)['Antal']
        .sum()
        .reset_index(name='Total_Antal')
        .sort_values('Total_Antal', ascending=False)
    )
    
    # Create detailed summary
    df['Kildenr.'] = df['Kildenr.'].astype(str).str.strip()
    df['Varenr.'] = df['Varenr.'].astype(str).str.strip()
    df['Antal_per_Varenr'] = df.groupby(['Kildenr.', 'Varenr.'])['Antal'].transform('sum')
    
    summary = (
        df.drop_duplicates(subset=['Kildenr.', 'Varenr.'])
        .loc[:, ['Kildenr.', 'Varenr.', 'Antal_per_Varenr']]
        .sort_values(['Antal_per_Varenr'], ascending=False)
        .reset_index(drop=True)
    )
    
    # Save detailed analysis
    summary.to_excel('kildenr_køb_data.xlsx', index=False)
    print("Supplier mismatch analysis saved to 'kildenr_køb_data.xlsx'")
    
    return not_same_kildenr_leverandørnr, summary

# ===== PRODUCT DATA ENRICHMENT =====
def match_products_with_master_data(unikke_varenumre_df, varer_df):
    """
    Match unique products with master data for enrichment.
    
    Args:
        unikke_varenumre_df (pandas.DataFrame): Unique product numbers
        varer_df (pandas.DataFrame): Product master data
    
    Returns:
        pandas.DataFrame: Matched products with master data
    """
    print("Matching products with master data...")
    
    # Define columns to extract from master data
    key_col = "Nummer"
    wanted_columns = [
        "Beskrivelse",
        "Beskrivelse 2",
        "Beskrivelse 3",
        "Basisenhed",
        "Kostpris",
        "Enhedspris",
    ]
    
    # Prepare data for matching
    unikke = unikke_varenumre_df.copy()
    unikke[key_col] = unikke[key_col].astype(str).str.strip()
    
    varer = varer_df.copy()
    varer[key_col] = varer[key_col].astype(str).str.strip()
    
    # Check for missing columns
    required_cols = [key_col] + wanted_columns
    missing = [c for c in required_cols if c not in varer.columns]
    if missing:
        raise ValueError(f"Missing expected columns: {missing}")
    
    # Perform the match
    matched = unikke.merge(varer, on=key_col, how="inner")
    
    # Report unmatched products
    not_found = unikke.loc[~unikke[key_col].isin(varer[key_col])].drop_duplicates()
    if not not_found.empty:
        print(f"{len(not_found)} products had no match in master data")
    
    # Save matched results
    matched.to_excel("matched_varenumre.xlsx", index=False)
    print(f"Matched {len(matched)} products with master data")
    
    return matched

def match_products_without_supplier_data(products_without_supplier_df, varer_df):
    """
    Match products sold without supplier information with master data.
    
    Args:
        products_without_supplier_df (pandas.DataFrame): Products without supplier info
        varer_df (pandas.DataFrame): Product master data
    
    Returns:
        pandas.DataFrame: Matched products with master data
    """
    print("Matching products without supplier data...")
    
    # Define columns to extract
    key_col = "Nummer"
    wanted_columns = [
        "Beskrivelse",
        "Beskrivelse 2",
        "Beskrivelse 3",
        "Basisenhed",
        "Kostpris",
        "Enhedspris",
    ]
    
    # Prepare data for matching
    unikke = products_without_supplier_df.copy()
    unikke[key_col] = unikke[key_col].astype(str).str.strip()
    unikke["Antal"] = pd.to_numeric(unikke["Antal"], errors="coerce")
    
    varer = varer_df.copy()
    varer[key_col] = varer[key_col].astype(str).str.strip()
    
    # Check for missing columns
    required_cols = [key_col] + wanted_columns
    missing = [c for c in required_cols if c not in varer.columns]
    if missing:
        raise ValueError(f"Missing expected columns: {missing}")
    
    # Perform the match
    matched = unikke.merge(varer, on=key_col, how="inner")
    
    # Report unmatched products
    not_found = unikke.loc[~unikke[key_col].isin(varer[key_col])].drop_duplicates()
    if not not_found.empty:
        print(f"{len(not_found)} products had no match in master data")
    
    # Save matched results
    matched.to_excel("matched_varenumre_uden_lev.xlsx", index=False)
    print(f"Matched {len(matched)} products without supplier data")
    
    return matched

# ===== MAIN EXECUTION =====
def main():
    """
    Main function that orchestrates the comprehensive analysis process.
    """
    print("=== OVERALL ANALYSIS SCRIPT ===")
    print("This script performs comprehensive analysis of the procurement dataset.\n")
    
    try:
        # Load main datasets
        vareposter, varer = load_main_datasets()
        
        # Extract unique products
        unikke_varenumre = extract_unique_products(vareposter)
        
        # Separate transaction types
        vareposter_salg, vareposter_køb = separate_transaction_types(vareposter)
        
        # Analyze sales by product and supplier
        antal_per_varenr_leverandor = analyze_sales_by_product_supplier(vareposter_salg)
        
        # Analyze purchases by product and supplier
        antal_per_varenr_leverandor_purchased = analyze_purchases_by_product_supplier(vareposter_køb)
        
        # Analyze products without supplier information
        antal_uden_leverandør = analyze_products_without_supplier(vareposter_salg)
        
        # Analyze shared products (sold by multiple suppliers)
        shared_products_analysis = analyze_shared_products_sales(vareposter_salg)
        
        # Analyze supplier number mismatches
        supplier_mismatches, supplier_summary = analyze_supplier_number_mismatches(vareposter_køb)
        
        # Match products with master data
        matched_products = match_products_with_master_data(unikke_varenumre, varer)
        
        # Match products without supplier data
        matched_products_no_supplier = match_products_without_supplier_data(antal_uden_leverandør, varer)
        
        print("\n=== ANALYSIS COMPLETE ===")
        print("All analysis steps have been completed successfully.")
        print("Review the generated Excel files for detailed results.")
        
    except Exception as e:
        print(f"Error during analysis: {e}")
        raise

if __name__ == "__main__":
    main()
