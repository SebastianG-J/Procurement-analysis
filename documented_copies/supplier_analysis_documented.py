"""
Supplier Analysis Script - Documented Version
=============================================

This script performs supplier-specific analysis for key suppliers in the procurement dataset.
It provides detailed insights into individual supplier performance, product relationships,
and business patterns.

Purpose:
- Analyze individual supplier performance and patterns
- Identify supplier-specific product relationships
- Evaluate supplier reliability and data quality
- Generate supplier-specific reports and insights
- Support supplier relationship management decisions

Key Features:
- Supplier-specific transaction analysis
- Product relationship mapping
- Performance metrics calculation
- Data quality assessment per supplier
- Comparative supplier analysis

Author: [Your Name]
Date: [Date]
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ===== DATA LOADING =====
def load_supplier_data():
    """
    Load the main transaction data for supplier analysis.
    
    Returns:
        pandas.DataFrame: Main transaction dataset
    """
    print("Loading transaction data for supplier analysis...")
    
    # Load main transaction data
    vareposter = pd.read_excel('Vareposter Alfotech.xlsx', header=0)
    print(f"Transaction data loaded: {vareposter.shape}")
    
    return vareposter

# ===== SUPPLIER-SPECIFIC ANALYSIS =====
def analyze_supplier_performance(df, supplier_id):
    """
    Analyze performance metrics for a specific supplier.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_id (str): Supplier identifier to analyze
    
    Returns:
        dict: Dictionary containing supplier performance metrics
    """
    print(f"Analyzing performance for supplier: {supplier_id}")
    
    # Filter data for the specific supplier
    supplier_data = df[df['Leverandørnr.'] == supplier_id].copy()
    
    if len(supplier_data) == 0:
        print(f"No data found for supplier: {supplier_id}")
        return {}
    
    # Calculate basic metrics
    total_transactions = len(supplier_data)
    unique_products = supplier_data['Varenr.'].nunique()
    
    # Separate sales and purchases
    sales_data = supplier_data[supplier_data['Posttype'].isin(['Salg', 'Montageforbrug'])]
    purchase_data = supplier_data[supplier_data['Posttype'] == 'Køb']
    
    # Calculate transaction counts
    sales_count = len(sales_data)
    purchase_count = len(purchase_data)
    
    # Calculate monetary metrics
    total_sales_value = sales_data['Salgsbeløb (faktisk)'].sum()
    total_cost_value = purchase_data['Kostbeløb (faktisk)'].sum()
    
    # Calculate quantity metrics
    total_sales_quantity = abs(sales_data['Antal'].sum())
    total_purchase_quantity = purchase_data['Antal'].sum()
    
    # Compile results
    performance_metrics = {
        'supplier_id': supplier_id,
        'total_transactions': total_transactions,
        'unique_products': unique_products,
        'sales_count': sales_count,
        'purchase_count': purchase_count,
        'total_sales_value': total_sales_value,
        'total_cost_value': total_cost_value,
        'total_sales_quantity': total_sales_quantity,
        'total_purchase_quantity': total_purchase_quantity,
        'sales_purchase_ratio': sales_count / purchase_count if purchase_count > 0 else float('inf')
    }
    
    print(f"Analysis complete for {supplier_id}")
    return performance_metrics

def analyze_supplier_products(df, supplier_id):
    """
    Analyze product relationships for a specific supplier.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_id (str): Supplier identifier to analyze
    
    Returns:
        pandas.DataFrame: Product analysis for the supplier
    """
    print(f"Analyzing products for supplier: {supplier_id}")
    
    # Filter data for the specific supplier
    supplier_data = df[df['Leverandørnr.'] == supplier_id].copy()
    
    if len(supplier_data) == 0:
        print(f"No data found for supplier: {supplier_id}")
        return pd.DataFrame()
    
    # Analyze products
    product_analysis = (
        supplier_data.groupby('Varenr.')
        .agg({
            'Antal': ['sum', 'count'],
            'Salgsbeløb (faktisk)': 'sum',
            'Kostbeløb (faktisk)': 'sum',
            'Posttype': lambda x: list(x.unique())
        })
        .reset_index()
    )
    
    # Flatten column names
    product_analysis.columns = [
        'Varenr.', 'Total_Quantity', 'Transaction_Count',
        'Total_Sales_Value', 'Total_Cost_Value', 'Transaction_Types'
    ]
    
    # Add additional metrics
    product_analysis['Avg_Quantity_per_Transaction'] = (
        product_analysis['Total_Quantity'] / product_analysis['Transaction_Count']
    )
    
    # Sort by total quantity
    product_analysis = product_analysis.sort_values('Total_Quantity', ascending=False)
    
    print(f"Product analysis complete for {supplier_id}")
    return product_analysis

def analyze_supplier_trends(df, supplier_id):
    """
    Analyze temporal trends for a specific supplier.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_id (str): Supplier identifier to analyze
    
    Returns:
        pandas.DataFrame: Temporal analysis for the supplier
    """
    print(f"Analyzing trends for supplier: {supplier_id}")
    
    # Filter data for the specific supplier
    supplier_data = df[df['Leverandørnr.'] == supplier_id].copy()
    
    if len(supplier_data) == 0:
        print(f"No data found for supplier: {supplier_id}")
        return pd.DataFrame()
    
    # Ensure date column is datetime
    supplier_data['Bogføringsdato'] = pd.to_datetime(supplier_data['Bogføringsdato'])
    
    # Add month and year columns for grouping
    supplier_data['Year'] = supplier_data['Bogføringsdato'].dt.year
    supplier_data['Month'] = supplier_data['Bogføringsdato'].dt.month
    supplier_data['YearMonth'] = supplier_data['Bogføringsdato'].dt.to_period('M')
    
    # Analyze monthly trends
    monthly_trends = (
        supplier_data.groupby('YearMonth')
        .agg({
            'Antal': 'sum',
            'Salgsbeløb (faktisk)': 'sum',
            'Kostbeløb (faktisk)': 'sum',
            'Varenr.': 'nunique'
        })
        .reset_index()
    )
    
    # Rename columns
    monthly_trends.columns = [
        'YearMonth', 'Total_Quantity', 'Total_Sales_Value',
        'Total_Cost_Value', 'Unique_Products'
    ]
    
    print(f"Trend analysis complete for {supplier_id}")
    return monthly_trends

# ===== COMPARATIVE ANALYSIS =====
def compare_suppliers(df, supplier_ids):
    """
    Compare performance across multiple suppliers.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_ids (list): List of supplier identifiers to compare
    
    Returns:
        pandas.DataFrame: Comparative analysis results
    """
    print(f"Comparing {len(supplier_ids)} suppliers...")
    
    comparison_results = []
    
    for supplier_id in supplier_ids:
        # Get performance metrics for each supplier
        performance = analyze_supplier_performance(df, supplier_id)
        if performance:
            comparison_results.append(performance)
    
    # Create comparison DataFrame
    comparison_df = pd.DataFrame(comparison_results)
    
    # Calculate additional comparison metrics
    if len(comparison_df) > 0:
        comparison_df['sales_efficiency'] = (
            comparison_df['total_sales_value'] / comparison_df['total_transactions']
        )
        comparison_df['product_diversity'] = (
            comparison_df['unique_products'] / comparison_df['total_transactions']
        )
    
    print("Supplier comparison complete")
    return comparison_df

# ===== DATA QUALITY ASSESSMENT =====
def assess_supplier_data_quality(df, supplier_id):
    """
    Assess data quality for a specific supplier.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_id (str): Supplier identifier to analyze
    
    Returns:
        dict: Data quality metrics for the supplier
    """
    print(f"Assessing data quality for supplier: {supplier_id}")
    
    # Filter data for the specific supplier
    supplier_data = df[df['Leverandørnr.'] == supplier_id].copy()
    
    if len(supplier_data) == 0:
        print(f"No data found for supplier: {supplier_id}")
        return {}
    
    # Calculate data quality metrics
    quality_metrics = {
        'supplier_id': supplier_id,
        'total_records': len(supplier_data),
        'missing_varenr': supplier_data['Varenr.'].isnull().sum(),
        'missing_beskrivelse': supplier_data['Beskrivelse'].isnull().sum(),
        'duplicate_records': supplier_data.duplicated().sum(),
        'zero_quantities': (supplier_data['Antal'] == 0).sum(),
        'negative_quantities': (supplier_data['Antal'] < 0).sum(),
        'date_range_start': supplier_data['Bogføringsdato'].min(),
        'date_range_end': supplier_data['Bogføringsdato'].max()
    }
    
    # Calculate percentages
    quality_metrics['missing_varenr_pct'] = (
        quality_metrics['missing_varenr'] / quality_metrics['total_records'] * 100
    )
    quality_metrics['missing_beskrivelse_pct'] = (
        quality_metrics['missing_beskrivelse'] / quality_metrics['total_records'] * 100
    )
    
    print(f"Data quality assessment complete for {supplier_id}")
    return quality_metrics

# ===== REPORT GENERATION =====
def generate_supplier_report(df, supplier_id, output_path=None):
    """
    Generate a comprehensive report for a specific supplier.
    
    Args:
        df (pandas.DataFrame): Main transaction dataset
        supplier_id (str): Supplier identifier to analyze
        output_path (str): Path to save the report (optional)
    
    Returns:
        dict: Complete supplier analysis report
    """
    print(f"Generating comprehensive report for supplier: {supplier_id}")
    
    # Perform all analyses
    performance = analyze_supplier_performance(df, supplier_id)
    products = analyze_supplier_products(df, supplier_id)
    trends = analyze_supplier_trends(df, supplier_id)
    quality = assess_supplier_data_quality(df, supplier_id)
    
    # Compile complete report
    report = {
        'supplier_id': supplier_id,
        'performance_metrics': performance,
        'product_analysis': products,
        'trend_analysis': trends,
        'data_quality': quality,
        'summary': {
            'total_transactions': performance.get('total_transactions', 0),
            'unique_products': performance.get('unique_products', 0),
            'data_quality_score': 100 - quality.get('missing_varenr_pct', 0)
        }
    }
    
    # Save report if output path provided
    if output_path:
        # Save performance metrics
        if performance:
            pd.DataFrame([performance]).to_excel(
                f"{output_path}_{supplier_id}_performance.xlsx", index=False
            )
        
        # Save product analysis
        if not products.empty:
            products.to_excel(f"{output_path}_{supplier_id}_products.xlsx", index=False)
        
        # Save trend analysis
        if not trends.empty:
            trends.to_excel(f"{output_path}_{supplier_id}_trends.xlsx", index=False)
        
        print(f"Report saved to {output_path}_{supplier_id}_*.xlsx")
    
    print(f"Report generation complete for {supplier_id}")
    return report

# ===== MAIN EXECUTION =====
def main():
    """
    Main function that orchestrates the supplier analysis process.
    """
    print("=== SUPPLIER ANALYSIS SCRIPT ===")
    print("This script performs comprehensive analysis of individual suppliers.\n")
    
    try:
        # Load data
        df = load_supplier_data()
        
        # Define suppliers to analyze (example supplier IDs)
        # Note: Replace with actual supplier IDs from your data
        example_suppliers = ['SUPPLIER1', 'SUPPLIER2', 'SUPPLIER3']
        
        # Get actual top suppliers from data
        top_suppliers = df['Leverandørnr.'].value_counts().head(5).index.tolist()
        print(f"Top 5 suppliers by transaction count: {top_suppliers}")
        
        # Analyze each supplier
        for supplier_id in top_suppliers[:3]:  # Analyze top 3 suppliers
            print(f"\n{'='*50}")
            print(f"ANALYZING SUPPLIER: {supplier_id}")
            print(f"{'='*50}")
            
            # Generate comprehensive report
            report = generate_supplier_report(
                df, 
                supplier_id, 
                output_path=f"supplier_analysis_{supplier_id}"
            )
            
            # Print summary
            summary = report['summary']
            print(f"Summary for {supplier_id}:")
            print(f"  - Total transactions: {summary['total_transactions']:,}")
            print(f"  - Unique products: {summary['unique_products']:,}")
            print(f"  - Data quality score: {summary['data_quality_score']:.1f}%")
        
        # Perform comparative analysis
        print(f"\n{'='*50}")
        print("COMPARATIVE SUPPLIER ANALYSIS")
        print(f"{'='*50}")
        
        comparison = compare_suppliers(df, top_suppliers[:5])
        if not comparison.empty:
            comparison.to_excel("supplier_comparison.xlsx", index=False)
            print("Comparative analysis saved to 'supplier_comparison.xlsx'")
        
        print("\n=== SUPPLIER ANALYSIS COMPLETE ===")
        print("All supplier analyses have been completed successfully.")
        print("Review the generated Excel files for detailed results.")
        
    except Exception as e:
        print(f"Error during supplier analysis: {e}")
        raise

if __name__ == "__main__":
    main()
