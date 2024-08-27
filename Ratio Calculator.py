#!/usr/bin/env python
# coding: utf-8

# In[1]:


from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from collections import Counter
from datetime import datetime
import os
from collections import defaultdict
from openpyxl.utils import get_column_letter
import requests
import numpy as np
import re
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from datetime import datetime, timedelta


# In[86]:


file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\FILT_MSFT.xlsx"  # Replace with your actual file path
sheet_name = "Filtered Tables and Rows"

# Read the Excel sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)
# Define the mapping of variables to search terms and skip flag
mapping = {
    'Quick': {'terms': ['Total cash', 'cash equivalents', 'Short-term investments','marketable securities','Total Current Liabilities','Receivable'], 'skip': True},
    'Current': {'terms': ['Total Current Assets', 'Total Current Liabilities'], 'skip': True},
    'Cash' : {'terms': ['Total cash', 'cash equivalents','Total Current Liabilities'], 'skip': True},
    'Receivables Turnover Ratio' : {'terms': ['Total Revenue','Receivable'], 'skip': True},
    'Inventory turnover ratio': {'terms': ['Cost of Revenue', 'Inventories'], 'skip': True},
    'Payables turnover ratio': {'terms': ['Cost of Revenue', 'Payable'], 'skip': True},
    'Working capital turnover ratio': {'terms': ['Total Revenue','Total Current Assets', 'Total Current Liabilities'], 'skip': True},
    'Fixed asset turnover ratio': {'terms': ['Total Revenue','Property and equipment, net'], 'skip': True},
    'Total asset turnover ratio': {'terms': ['Total Revenue','Total Assets'], 'skip': True},
    'Gross profit margin': {'terms': ['gross','Total Revenue'], 'skip': True},
    'Operating profit margin': {'terms': ['operating income','Total Revenue'], 'skip': True},
    'Pretax margin': {'terms': ['operating income','interest expense','total revenue'], 'skip': True},
    'Net profit margin': {'terms': ['net income','total revenue'], 'skip': True},
    'Operating return on assets': {'terms': ['operating income','total assets'], 'skip': True},
    'Return on assets': {'terms': ['net income','total assets'], 'skip': True},
    'Return on equity': {'terms': ['net income',"total stockholders' equity","total shareholders' equity"], 'skip': True},
    'Return on invested capital': {'terms': ['Income before income tax', 'interest expense','short-term debt','long-term debt',"total stockholders' equity","total shareholders' equity"], 'skip': True},
    'effective tax rate': {'terms': ['Income before income tax', 'tax expense',], 'skip': True},
    'Return on equity': {'terms': ['net income', "total stockholders' equity","total shareholders' equity"], 'skip': True},
    'Tax burden': {'terms': ['net income', 'Income before income tax'], 'skip': True},
    'Financial Leverage Ratio': {'terms': ['total assets', "total stockholders' equity","total shareholders' equity"], 'skip': True},
    'Debt to assets': {'terms': ['short-term debt','long-term debt','total assets'], 'skip': True},
    'Debt to equity': {'terms': ['short-term debt','long-term debt',"total stockholders' equity","total shareholders' equity"], 'skip': True},
    'Debt to capital': {'terms': ['short-term debt','long-term debt',"total stockholders' equity","total shareholders' equity"], 'skip': True},
    'Interest Coverage Ratio': {'terms': ['Income before income tax', 'interest expense'], 'skip': True},
    # Add more mappings as needed
}

def is_small_float(value, threshold=100.0):
    try:
        # Remove parentheses and commas, if present
        if isinstance(value, str):
            value = value.replace('(', '').replace(')', '').replace(',', '')
            if '%' in value:
                # Convert percentage to float
                value = float(value.replace('%', '').strip())
            else:
                value = float(value.strip())
        return abs(float(value)) < threshold
    except (ValueError, TypeError):
        return False

# Initialize an empty list to collect results
results = []

# Iterate through each variable and its associated search terms
for variable, settings in mapping.items():
    terms = settings['terms']
    skip = settings.get('skip', False)
    
    print(f"Searching for: {variable}")
    
    for term in terms:
        print(f"\nSearching for term: {term}")
        
        # Filter the DataFrame to include only rows where the target column contains the current term, case-insensitive
        filtered_df = df[df[target_column].str.contains(term, case=False, na=False)]
        
        # Iterate through filtered rows
        for index, row in filtered_df.iterrows():
            should_skip = False

            if skip:
                # Find the index of the target column and access the cell to the right
                target_col_index = df.columns.get_loc(target_column)
                
                # Check if the target_col_index + 1 is within bounds
                if target_col_index + 1 < len(row):
                    next_cell_value = row.iloc[target_col_index + 1]  # Accessing the cell to the right
                    print(f"Checking next cell value for row {index}: {next_cell_value}")  # Debug information

                    if is_small_float(next_cell_value):
                        print(f"Skipped row {index} due to small float or percentage: {next_cell_value}")
                        should_skip = True

            if not should_skip:
                # Collect data in results list
                result = {
                    'Variable': variable,
                    'Term': term,
                    'Row Index': index,
                    'Table Name': row.get('Table Name', 'N/A'),
                    'Table Index': row.get('Table Index', 'N/A'),
                    'Column 1': row[target_column],
                }
                for col in df.columns:
                    result[col] = row[col]
                
                results.append(result)
                print(f"Row {index}:")
                for col in df.columns:
                    # Print column name and value
                    print(f"{col}: {row[col]}")
                print()  # Add a blank line for readability
                break  # Stop after printing the first valid matching row for the current term

    # Add a separator between different variables for clarity
    print("="*40)

# Convert results list to a DataFrame
results_df = pd.DataFrame(results)

# Display the results DataFrame
print("Results DataFrame:")
print(results_df)

# Save results DataFrame to a new Excel file
results_df.to_excel(r"C:\Users\thoma\Desktop\Projects\Project Oversight\search_results.xlsx", index=False)


# In[ ]:




