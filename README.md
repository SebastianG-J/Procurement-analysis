# Procurement Analysis Project

A data analysis project for procurement and sales data analysis, focusing on supplier performance, product analysis, and automated data collection.

## 📋 Project Overview

This project provides tools and analysis for:
- Sales and purchase data analysis
- Supplier performance metrics
- Product data enrichment through web scraping
- Data quality assessment

## 📁 Project Structure

```
BI project/
├── 📊 Data Analysis Notebooks
│   ├── data exploration.ipynb          # Initial data exploration
│   ├── overall_analysis.ipynb          # Comprehensive analysis
│   └── prisudvikling_stål.ipynb        # Price development analysis
│
├── 🤖 Web Scraping Scripts
│   ├── H1_scraper_script.py            # Main product data scraper
│   ├── meter_pr_rulle_script.py        # Specialized data collection
│   └── results_merger.py               # Data merging utilities
│
├── 📈 Supplier Analysis
│   └── Suppliers/
│       ├── data_analysis_dicsa.ipynb   # Supplier-specific analysis
│       ├── ingerslev_analysis.ipynb    # Supplier-specific analysis
│       └── MTG_analysis.ipynb          # Supplier-specific analysis
│
└── 📄 Data Files
    ├── Vareposter Alfotech.xlsx        # Main transaction data
    ├── Varer.xlsx                      # Product master data
    └── Various analysis outputs...
```

## 🔍 Key Components

### 1. Data Exploration (`data exploration.ipynb`)
- Data loading and structure analysis
- Transaction type distribution
- Supplier activity analysis
- Data quality assessment

### 2. Overall Analysis (`overall_analysis.ipynb`)
- Product identification and categorization
- Sales vs purchase volume analysis
- Supplier performance metrics
- Data enrichment and matching

### 3. Web Scraping Automation
- **H1 Scraper**: Automated product specification collection
- **Meter per Roll Script**: Specialized roll-based product data
- **Results Merger**: Data combination utilities

## 🛠️ Technical Requirements

### Python Dependencies
```bash
pip install pandas openpyxl matplotlib seaborn numpy selenium
```

### Web Scraping Requirements
- Chrome browser
- ChromeDriver (automatically managed by Selenium)
- Internet connection

## 🔧 Usage Instructions

### 1. Data Analysis
```python
# Load main transaction data
import pandas as pd
vareposter = pd.read_excel('Vareposter Alfotech.xlsx')

# Separate sales and purchases
sales_data = vareposter[vareposter['Posttype'].isin(['Salg', 'Montageforbrug'])]
purchase_data = vareposter[vareposter['Posttype'] == 'Køb']
```

### 2. Web Scraping
```bash
# Run the main scraper
python H1_scraper_script.py --input-file your_data_file.xlsx

# Run meter per roll scraper
python meter_pr_rulle_script.py --input-file your_product_list.xlsx
```

### 3. Data Merging
```bash
# Merge scraped results
python results_merger.py
```

## 📈 Output Files

The analysis generates several Excel files:
- `results_simplified.xlsx` - Core analysis results
- `results_fast.xlsx` - Quick analysis output
- `results_merged.xlsx` - Combined analysis results
- Various intermediate analysis files

## 🤝 Contributing

For modifications or extensions:
1. Ensure data privacy and security compliance
2. Test web scraping scripts with appropriate delays
3. Validate analysis outputs against business requirements
4. Document any changes to data processing logic

## 📝 Notes

- All monetary values are in local currency
- Dates follow European format (DD-MM-YYYY)
- Product numbers are standardized as strings for consistency
- Web scraping includes appropriate delays to respect supplier websites

---

**Last Updated**: April 2025