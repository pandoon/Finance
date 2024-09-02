#!/usr/bin/env python
# coding: utf-8

# In[12]:


import pandas as pd
import numpy as np
from twelvedata import TDClient
import matplotlib.pyplot as plt
import os
from datetime import datetime
from typing import Dict
import time
import openpyxl


# In[14]:


file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Technology\path_to_modified_file.xlsx"
td = TDClient(apikey="eeef2551d821410d94cd1b93b70692cc")#Get your own api key from twelve data or use mine lol


# In[15]:


call_stock(file_path)
#Run all cells below one by one. Then come up here and run call_stock


# In[3]:


def process_time_series(symbol, interval='1month', outputsize=172, timezone='America/New_York'):
    # Fetch time series data
    ts = td.time_series(
        symbol=symbol,
        interval=interval,
        outputsize=outputsize,
        timezone=timezone,
    )
    
    # Convert to DataFrame
    df = ts.as_pandas()
    
    # Drop unnecessary columns
    df = df.drop(columns=['open', 'low', 'high'])
    
    # Reset index if 'datetime' is the index
    if df.index.name == 'datetime':
        df.reset_index(inplace=True)
    
    # Check if 'datetime' column exists and convert to 'Year Quarter' format
    if 'datetime' in df.columns:
        # Convert datetime to Year Quarter format and add a space between year and quarter
        df['Period End'] = df['datetime'].dt.to_period('Q').astype(str)
        df['Period End'] = df['Period End'].str.replace(r'(\d{4})(Q\d)', r'\1 \2', regex=True)
        
        # Drop duplicate quarters, keeping only the most recent entry
        df = df.sort_values(by='datetime', ascending=False)  # Sort by datetime descending
        df = df.drop_duplicates(subset='Period End', keep='first')  # Keep only the most recent entry per quarter
        df = df.drop(columns = ['datetime'])
    else:
        print("Column 'datetime' not found.")
    time.sleep(3)
    return df


# In[4]:


def add_ticker_column_to_sheets(file_path):
    # Read the Excel file and get all sheet names
    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names
    
    # Use ExcelWriter to write back changes to the same file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # Loop through each sheet
        for sheet_name in sheet_names:
            # Read the sheet into a DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Check if the 'Ticker' column already exists
            #if 'Ticker' in df.columns:
                #continue  # Skip to the next sheet

            # Extract only the ticker symbol from the sheet name (before the comma)
            ticker = sheet_name.split(',')[0]
            
            # Add a new column named 'Ticker' with the ticker value in every cell
            df['Ticker'] = ticker
            
            # Write the updated DataFrame back to the sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Print the modified DataFrame for each sheet
        return df


# In[5]:


def months_ago_min_date(file_path, ticker):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=ticker)
    
    # Define a function to convert 'Period End' to a comparable format
    def period_to_date(period):
        if isinstance(period, str):
            year, quarter = period.split(' ')
            year = int(year)
            quarter = int(quarter[1])
            return pd.Timestamp(year=year, month=quarter*3, day=1)
        else:
            return pd.NaT  # Return Not a Time for non-string entries

    # Convert 'Period End' to datetime for comparison
    df['Period End Date'] = df['Period End'].apply(period_to_date)

    # Find the minimum date
    min_date = df['Period End Date'].min()

    # Calculate the difference in months between the current date and min_date
    today = pd.Timestamp(datetime.today())
    months_ago = (today.year - min_date.year) * 12 + today.month - min_date.month

    # Drop the temporary 'Period End Date' column
    df = df.drop(columns=['Period End Date'])

    # Return the number of months ago
    return months_ago


# In[13]:


def call_stock(file_path):
    check_and_remove_placeholder(file_path)
    add_ticker_column_to_sheets(file_path)
    origin = pd.ExcelFile(file_path)
    sheet_names = origin.sheet_names
    tickers = [name.split(',')[0] for name in sheet_names]
    merged_df = pd.DataFrame()
    
    for sheet_name, ticker in zip(sheet_names[:], tickers[:]):
        stock_sheet = process_time_series(ticker)
        data_sheet = pd.read_excel(io=file_path, sheet_name=sheet_name)
        
        if 'Period End' in stock_sheet.columns and 'Period End' in data_sheet.columns:
            # Merge the dataframes
            merged = pd.merge(stock_sheet, data_sheet, on='Period End', how='inner')
            
            # Identify columns with '_x' and '_y' suffixes
            x_columns = [col for col in merged.columns if col.endswith('_x')]
            y_columns = [col for col in merged.columns if col.endswith('_y')]
            
            # Drop '_y' columns
            merged = merged.drop(columns=y_columns)
            
            # Rename '_x' columns by removing the '_x' suffix
            rename_dict = {col: col[:-2] for col in x_columns}
            merged = merged.rename(columns=rename_dict)
            
            # Ensure 'close' and 'volume' columns are rightmost
            cols = [col for col in merged.columns if col not in ['close', 'volume']]
            cols += ['close', 'volume']
            merged = merged[cols]
            merged['Next Period Close'] = merged['close'].shift(1)
            merged['P/E'] = merged['close'] / merged['Basic EPS']
            
            output_file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Technology\path_to_modified_file.xlsx"
            with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                merged.to_excel(writer, sheet_name=sheet_name, index=False)
        
        time.sleep(5)  # Delay to manage rate limits or server load
    add_ticker_column_to_sheets(file_path)
    return print('Stock prices and Volume data added to file.')


# In[7]:


def check_and_remove_placeholder(file_path):
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File '{file_path}' does not exist.")
        return
    
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(file_path)
        
        # Check if 'Placeholder' sheet exists
        if 'Sheet' in workbook.sheetnames:
            # Delete the 'Placeholder' sheet
            del workbook['Sheet']
            print("'Sheet' sheet has been removed.")
            
            # Save the workbook by overwriting the existing file
            workbook.save(file_path)
            print(f"Workbook '{file_path}' has been updated and saved.")
        else:
            print("'Sheet' sheet does not exist in the workbook.")
    
    except Exception as e:
        print(f"An error occurred: {e}")


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




