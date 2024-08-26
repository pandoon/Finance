#!/usr/bin/env python
# coding: utf-8

# In[15]:


import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import re
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
import os
from datetime import datetime, timedelta


# In[165]:


link = "https://app.quotemedia.com/data/downloadFiling?webmasterId=90423&ref=318088749&type=HTML&symbol=NVDA&cdn=432b30401cba4ad888d63effbbb3098a&companyName=NVIDIA+Corporation&formType=10-K&formDescription=Annual+report+pursuant+to+Section+13+or+15%28d%29&dateFiled=2024-02-21"


# In[169]:



headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
}

page = requests.get(link,headers=headers)

soup = BeautifulSoup(page.text, 'html.parser')

def contains_number(text):
    """Check if the text contains any integer or float."""
    return bool(re.search(r'\b\d+(\.\d+)?\b', text))

def contains_keyword(text, keywords):
    """Check if the text contains any of the keywords."""
    return any(keyword.lower() in text.lower() for keyword in keywords)

keywords = [
    'Revenue', 'Gross Margin', 'Operating Expenses', 'Operating Income', 
    'Net Income', 'Net Income per Share', 'Segment', 'Interest Income/Expense', 
    'Other Income (Expense), Net', 'Cash and Cash Equivalents', 
    'Marketable Securities', 'Net Cash Provided by (Used in)', 
    'Outstanding Indebtedness', 'Consolidated Statements of Income', 
    'Consolidated Statements of Comprehensive Income', 
    'Consolidated Balance Sheets', 'Shareholders\' Equity', 'Cash Flows', 
    'Leases', 'Fair Value of Financial Assets/Liabilities', 
    'Balance Sheet Components', 'Accrued and Other Liabilities', 
    'Deferred Revenue', 'Debt', 'Income Taxes', 'Segment Information', 
    'Geographic Revenue', 'Specialized Market Segment', 
    'Long-Lived Assets by Region', 'Segments', 'Tax' , 'Balance', 'Income', 'Assets','Liability', 'Property', 'Equipment', 'Inventories','Inventory',
    'Carrying', 'Amortized', 'Fiscal', 'Research', 'Development', 'Expense', 'Total'
]

def process_row_data(cells):
    row_data = []
    i = 0
    
    while i < len(cells):
        cell_value = cells[i].get_text(strip=True)
        
        # Handle cases where a cell starts with '%' and ends with '$'
        if cell_value.startswith('%') and cell_value.endswith('$'):
            if row_data:
                row_data[-1] += cell_value[0]  # Concatenate '%' with the previous cell
            if i + 1 < len(cells):
                next_cell_value = cells[i + 1].get_text(strip=True)
                row_data.append(cell_value[1:-1] + next_cell_value)  # Concatenate the middle part with the next cell
                i += 1  # Skip the next cell since we've processed it
            else:
                row_data.append(cell_value[1:])  # Add the middle part if no next cell
        elif cell_value.startswith('%'):
            # If the cell starts with '%', concatenate '%' to the previous cell
            if row_data:
                row_data[-1] += cell_value
            else:
                # If no previous cell, just add the cell
                row_data.append(cell_value)
        elif cell_value.endswith('$'):
            # If the cell ends with '$', handle concatenation with the next cell
            if i + 1 < len(cells):
                next_cell_value = cells[i + 1].get_text(strip=True)
                row_data.append(cell_value[:-1] + next_cell_value)  # Concatenate with the next cell
                i += 1  # Skip the next cell since we've processed it
            else:
                row_data.append(cell_value)
        elif cell_value == ')':
            # If the cell contains only ')', concatenate it with the previous cell
            if row_data:
                row_data[-1] += cell_value
            else:
                row_data.append(cell_value)
        else:
            row_data.append(cell_value)
        
        i += 1
    
    return row_data

def remove_empty_cells_after_column_b(data):
    """Remove empty cells to the right of column B but keep column B."""
    cleaned_data = []
    for row in data:
        # Keep column A and B
        if len(row) > 2:
            cleaned_row = row[:2] + [cell for cell in row[2:] if cell.strip()]
        else:
            cleaned_row = row
        cleaned_data.append(cleaned_row)
    return cleaned_data

# Your existing code for processing tables and creating DataFrames
tables = soup.find_all('table')

all_tables_data = []

for table_index, table in enumerate(tables):
    print(f"Table {table_index + 1}")

    # Initialize lists to hold rows
    rows = table.find_all('tr')
    
    table_contains_keyword = False
    
    for row_index, row in enumerate(rows):
        cells = row.find_all('td')
        
        # Check if any cell in this row contains a keyword
        row_contains_keyword = any(contains_keyword(cell.get_text(strip=True), keywords) for cell in cells)
        
        if row_contains_keyword:
            table_contains_keyword = True
            break  # No need to check further rows if we've found a keyword
    
    if table_contains_keyword:
        print(f"Table {table_index + 1} contains keywords:")
        
        # Extract headers
        header_row = rows[0]  # Assuming the first row is the header
        headers = [header.get_text(strip=True) for header in header_row.find_all('th')]
        if not headers:  # If headers are not found in <th> tags, use <td> instead
            headers = [cell.get_text(strip=True) for cell in rows[0].find_all('td')]
        
        # Check if headers are empty and create default headers
        num_columns = len(headers)
        if all(header == '' for header in headers):
            headers = [f'Column {i+1}' for i in range(num_columns)]
        
        print("Headers:", headers)
        print(f"Number of columns: {num_columns}")

        # List to hold data for the current table
        table_data = []

        # Process rows
        for row_index, row in enumerate(rows):
            cells = row.find_all('td')
            row_data = process_row_data(cells)
            
            # Pad row data if it has fewer columns
            if len(row_data) < num_columns:
                row_data.extend([''] * (num_columns - len(row_data)))
            
            # Add row data to the table_data list
            table_data.append(row_data)
        
        # Remove empty cells to the right of column B
        table_data = remove_empty_cells_after_column_b(table_data)
        
        # Create DataFrame and add to all_tables_data list
        try:
            df = pd.DataFrame(table_data)
            
            # Ensure headers are set properly and unique
            if len(df.columns) > 0:
                df.columns = headers[:len(df.columns)]
            
            # Optional: Add a marker row for separation
            df.insert(0, 'Table Index', f'Table {table_index + 1}')
            
            # Reset index to avoid any potential index issues
            df.reset_index(drop=True, inplace=True)
            
            all_tables_data.append(df)
        except Exception as e:
            print(f"Failed to create DataFrame for Table {table_index + 1}: {e}")

# Concatenate all DataFrames into one
if all_tables_data:
    try:
        # Concatenate DataFrames
        all_tables_df = pd.concat(all_tables_data, ignore_index=True)
        
        # Define the base filename and get a unique filename
        base_filename = 'output.xlsx'
        unique_filename = get_unique_filename(base_filename)
        
        # Save all tables to a single sheet in Excel
        with pd.ExcelWriter(unique_filename, engine='openpyxl') as writer:
            all_tables_df.to_excel(writer, sheet_name='All_Tables', index=False)
        print(f"Data has been written to {unique_filename}")
    except Exception as e:
        print(f"Failed to write data to Excel: {e}")
else:
    print("No tables with keywords were found.")


# In[25]:





# In[ ]:





# In[ ]:





# In[ ]:




