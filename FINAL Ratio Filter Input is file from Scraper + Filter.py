#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[14]:


calculate_ratios(r"C:\Users\thoma\Desktop\Projects\Project Oversight\financial_data.xlsx", 'MSFT,0000789019')


# In[2]:


def sort_and_fill_gaps(df):
    # Define a function to convert 'Period End' to a comparable format
    def period_to_date(period):
        if isinstance(period, str):
            year, quarter = period.split(' ')
            year = int(year)
            quarter = int(quarter[1])
            return pd.Timestamp(year=year, month=quarter*3, day=1)
        else:
            return pd.NaT  # Return Not a Time for non-string entries

    # Convert 'Period End' to datetime for sorting and gap calculation
    df['Period End Date'] = df['Period End'].apply(period_to_date)

    # Filter for '2024 Q4' row, if it exists
    q4_2024_row = df[df['Period End'] == '2024 Q4']

    # If '2024 Q4' exists, remove it from the main DataFrame
    if not q4_2024_row.empty:
        df = df[df['Period End'] != '2024 Q4']
        # Sort the remaining DataFrame by 'Period End Date' in descending order
        df_sorted = df.sort_values(by='Period End Date', ascending=False).reset_index(drop=True)
        # Add '2024 Q4' row at the beginning
        df_sorted = pd.concat([q4_2024_row, df_sorted], ignore_index=True)
    else:
        # Sort the DataFrame by 'Period End Date' in descending order if '2024 Q4' doesn't exist
        df_sorted = df.sort_values(by['Period End Date'], ascending=False).reset_index(drop=True)

    # Initialize the result DataFrame
    result_df = pd.DataFrame(columns=df.columns)
    
    # Iterate over the sorted DataFrame
    for i in range(len(df_sorted)):
        current_row = df_sorted.iloc[[i]]  # Make it a DataFrame
        result_df = pd.concat([result_df, current_row], ignore_index=True)
        
        if i < len(df_sorted) - 1:
            # Check the next row's period
            next_row = df_sorted.iloc[[i + 1]]  # Make it a DataFrame
            current_date = period_to_date(current_row['Period End'].values[0])
            next_date = period_to_date(next_row['Period End'].values[0])
            
            # Calculate the number of quarters between the current and next row
            if pd.notna(current_date) and pd.notna(next_date):
                quarters_diff = (current_date.year - next_date.year) * 4 + (current_date.month // 3 - next_date.month // 3)
                
                # Ensure quarters_diff is a positive integer
                if isinstance(quarters_diff, int) and quarters_diff > 1:
                    # Add empty rows if there's a gap
                    for _ in range(quarters_diff - 1):
                        empty_row = pd.DataFrame([[np.nan] * len(df.columns)], columns=df.columns)
                        result_df = pd.concat([result_df, empty_row], ignore_index=True)
    
    # Drop the 'Period End Date' column and reset index
    result_df = result_df.drop(columns=['Period End Date']).reset_index(drop=True)
    
    return result_df


# In[5]:


def calculate_financial_metrics(xls, target_sheet_name):
    # Load the data from the specified sheet
    df = pd.read_excel(xls, sheet_name=target_sheet_name)

    # Filter rows for each 'Fact' value and create separate DataFrames
    fact_columns = [
        'AccountsReceivableNetCurrent', 'AccountsPayableCurrent', 'InventoryNet', 'AssetsCurrent','LiabilitiesCurrent',
        'GrossProfit', 'CostOfGoodsAndServicesSold', 'PropertyPlantAndEquipmentGross',
        'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment', 'Assets',
        'OperatingIncomeLoss', 'NetIncomeLoss', 'IncomeTaxExpenseBenefit', 'InterestExpense',
        'StockholdersEquity', 'CashAndCashEquivalentsAtCarryingValue', 'LongTermDebt'
    ]
    
    dataframes = {fact: df[df['Fact'] == fact].copy() for fact in fact_columns}
    
    # Filter with gaps for quarters
    for key in dataframes:
        dataframes[key] = sort_and_fill_gaps(dataframes[key])

    # Determine the minimum length of the DataFrames
    min_length = min(len(df) for df in dataframes.values())
    
    # Truncate all DataFrames to the minimum length
    for key in dataframes:
        dataframes[key] = dataframes[key].iloc[:min_length]

    combined_data = []

    # Process the DataFrames to calculate metrics
    for i in range(min_length - 1):  # Ignore the last row since it has no next row with averages
        metrics = {}
        metrics['Period End'] = dataframes['AccountsReceivableNetCurrent'].iloc[i]['Period End']

        # Define a function to calculate the average for given DataFrame and fact key
        def calculate_average(df_key):
            current_value = dataframes[df_key].iloc[i]['Value']
            next_index = i + 1
            while next_index < min_length and pd.isna(dataframes[df_key].iloc[next_index]['Value']):
                next_index += 1
            if next_index < min_length:
                next_value = dataframes[df_key].iloc[next_index]['Value']
                return (current_value + next_value) / 2
            return current_value

        # Calculate the required metrics
        metrics['Average Receivables'] = calculate_average('AccountsReceivableNetCurrent')
        metrics['Average Payables'] = calculate_average('AccountsPayableCurrent')
        metrics['Average Inventory'] = calculate_average('InventoryNet')
        metrics['Revenue'] = dataframes['CostOfGoodsAndServicesSold'].iloc[i]['Value'] + dataframes['GrossProfit'].iloc[i]['Value']
        metrics['Gross Profit'] = dataframes['GrossProfit'].iloc[i]['Value']
        metrics['Cost of Goods Sold'] = dataframes['CostOfGoodsAndServicesSold'].iloc[i]['Value']
        
        # Calculate average Gross PPE
        current_grossppe = dataframes['PropertyPlantAndEquipmentGross'].iloc[i]['Value']
        next_index = i + 1
        while next_index < min_length and pd.isna(dataframes['PropertyPlantAndEquipmentGross'].iloc[next_index]['Value']):
            next_index += 1
        if next_index < min_length:
            next_grossppe = dataframes['PropertyPlantAndEquipmentGross'].iloc[next_index]['Value']
            average_grossppe = (current_grossppe + next_grossppe) / 2
        else:
            average_grossppe = current_grossppe

        metrics['Average Net PPE'] = average_grossppe - calculate_average('AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment')
        
        # Calculate average total assets
        metrics['Average Total Assets'] = calculate_average('Assets')

        # Other metrics
        metrics['Operating Income'] = dataframes['OperatingIncomeLoss'].iloc[i]['Value']
        metrics['Net Income'] = dataframes['NetIncomeLoss'].iloc[i]['Value']
        metrics['EBT'] = metrics['Net Income'] + dataframes['IncomeTaxExpenseBenefit'].iloc[i]['Value']
        metrics['EBIT'] = metrics['EBT'] + dataframes['InterestExpense'].iloc[i]['Value']
        metrics['Effective Tax Rate'] = dataframes['IncomeTaxExpenseBenefit'].iloc[i]['Value'] / metrics['EBT']
        metrics['Current Assets'] = dataframes['AssetsCurrent'].iloc[i]['Value']
        metrics['Current Liabilities'] = dataframes['LiabilitiesCurrent'].iloc[i]['Value']
        
        # Average equity and invested capital
        metrics['Average Shareholder Equity'] = calculate_average('StockholdersEquity')
        current_equity = dataframes['StockholdersEquity'].iloc[i]['Value']
        current_ltdebt = dataframes['LongTermDebt'].iloc[i]['Value']
        current_cce = dataframes['CashAndCashEquivalentsAtCarryingValue'].iloc[i]['Value']
        next_index = i + 1
        while next_index < min_length and pd.isna(dataframes['StockholdersEquity'].iloc[next_index]['Value']):
            next_index += 1
        if next_index < min_length:
            next_equity = dataframes['StockholdersEquity'].iloc[next_index]['Value']
            next_ltdebt = dataframes['LongTermDebt'].iloc[next_index]['Value']
            next_cce = dataframes['CashAndCashEquivalentsAtCarryingValue'].iloc[next_index]['Value']
            next_equitydebtlesscash = next_equity + next_ltdebt - next_cce
            metrics['Average Invested Capital'] = (current_equity + current_ltdebt - current_cce + next_equity + next_ltdebt - next_cce) / 2

        combined_data.append(metrics)

    # Convert the combined data to a DataFrame and return
    results_df = pd.DataFrame(combined_data)
    return results_df


# In[13]:


def calculate_ratios(file_path, ticker):
    # Call the calculate_financial_metrics function
    df = calculate_financial_metrics(file_path, ticker)
    
    # Check if 'Current Assets' and 'Current Liabilities' columns are present
    if 'Current Assets' not in df.columns or 'Current Liabilities' not in df.columns:
        raise ValueError("Columns 'Current Assets' and 'Current Liabilities' are required in the DataFrame.")
    
    # Calculate Ratios
    df['Current'] = df['Current Assets'] / df['Current Liabilities']
    df['Receivables Turnover'] = df['Revenue'] / df['Average Receivables']
    df['Payables Turnover'] = df['Cost of Goods Sold'] / df['Average Payables']
    df['Inventory Turnover'] = df['Cost of Goods Sold'] / df['Average Inventory']
    df['Days Sales Outstanding'] = 90 / df['Receivables Turnover']
    df['Days Inventory On Hand'] = 90 / df['Inventory Turnover']
    df['Days Payables Outstanding'] = 90 / df['Payables Turnover']
    df['Cash Conversion Cycle'] = df['Days Inventory On Hand'] + df['Days Sales Outstanding'] - df['Days Payables Outstanding']
    df['Fixed Asset Turnover'] = df['Revenue'] / df['Average Net PPE']
    df['Total Asset Turnover'] = df['Revenue'] / df['Average Total Assets']
    df['Gross Margin'] = df['Gross Profit'] / df['Revenue']
    df['Operating Margin'] = df['Operating Income'] / df['Revenue']
    df['Pretax Margin'] = df['EBT'] / df['Revenue']
    df['Net Profit Margin'] = df['Net Income'] / df['Revenue']
    df['Operating Return on Assets'] = df['Operating Income'] / df['Average Total Assets']
    df['Return on Assets'] = df['Net Income'] / df['Average Total Assets']
    df['Return on Equity'] = df['Net Income'] / df['Average Shareholder Equity']
    df['ROIC'] = (df['EBIT'] * (1-df['Effective Tax Rate'])) / df['Average Invested Capital']
    df['Tax Burden'] = df['Net Income'] / df['EBT']
    df['Interest Burden'] = df['EBT'] / df['EBIT']
    df['EBIT Margin'] = df['EBIT'] / df['Revenue']
    df['Financial Leverage Ratio'] = df['Average Total Assets'] / df['Average Shareholder Equity']
    
    
    # Create a DataFrame with the results
    result_df = df[['Period End', 'Current','Receivables Turnover','Payables Turnover','Inventory Turnover',
                   'Days Sales Outstanding','Days Inventory On Hand','Days Payables Outstanding','Cash Conversion Cycle',
                   'Fixed Asset Turnover','Total Asset Turnover', 'Gross Margin','Operating Margin','Pretax Margin',
                   'Net Profit Margin','Operating Return on Assets', 'Return on Assets', 'Return on Equity',
                   'ROIC', 'Tax Burden', 'Interest Burden', 'EBIT Margin', 'Financial Leverage Ratio']].copy()
    
    # Optionally, you can save this result to a new Excel file or return it
    result_df.to_excel(r"C:\Users\thoma\Desktop\Projects\Project Oversight\current_ratios.xlsx", index=False)
    
    return result_df

# Example usage
#result_df = calculate_ratios(r"C:\Users\thoma\Desktop\Projects\Project Oversight\financial_data.xlsx", 'MSFT,0000789019')
#print(result_df)


# In[ ]:




