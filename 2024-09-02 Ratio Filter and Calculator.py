#!/usr/bin/env python
# coding: utf-8

# In[12]:


import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
import openpyxl


# In[13]:


#calculate_ratios(r"C:\Users\thoma\Desktop\Projects\Project Oversight\financial_data.xlsx", 'MSFT,0000789019')
#For Individual prints
calculate_all(r"C:\Users\thoma\Desktop\Projects\Project Oversight\financial_data.xlsx")
#Calculates all in an excel file.


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

    # Sort the DataFrame by 'Period End Date' in descending order
    df_sorted = df.sort_values(by='Period End Date', ascending=False).reset_index(drop=True)

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


# In[3]:


def calculate_financial_metrics(xls, target_sheet_name):
    # Load the data from the specified sheet
    df = pd.read_excel(xls, sheet_name=target_sheet_name)
    df = remove_receivable(file_path = xls, ticker=target_sheet_name)
    df = add_ppeg(file_path = xls, ticker=target_sheet_name)
    df = remove_COGS(file_path = xls, ticker = target_sheet_name)#You'll need this to deal with companies that have 0 CoGs (MasterCard)
    df = remove_interest_expense(file_path= xls, ticker=target_sheet_name)#You'll need this for companies without interest expense. (VIZIO)
    # Filter rows for each 'Fact' value and create separate DataFrames
    fact_columns = [
        'AccountsReceivableNetCurrent', 'AssetsCurrent','LiabilitiesCurrent',
        'RevenueFromContractWithCustomerExcludingAssessedTax', 'CostOfGoodsAndServicesSold', 'PropertyPlantAndEquipmentGross',
        'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment', 'Assets','StockholdersEquity',
        'OperatingIncomeLoss', 'NetIncomeLoss', 'IncomeTaxExpenseBenefit', 'InterestExpense', 'LiabilitiesAndStockholdersEquity',
        'CashAndCashEquivalentsAtCarryingValue', 'LongTermDebt','EarningsPerShareBasic',
    ]
    
    dataframes = {fact: df[df['Fact'] == fact].copy() for fact in fact_columns}
    
    # Filter with gaps for quarters
    for key in dataframes:
        dataframes[key] = sort_and_fill_gaps(dataframes[key])

    # Determine the minimum length of the DataFrames
    min_length = min(len(df) for df in dataframes.values())
    for name, df in dataframes.items():
        print(f"Length of {name}: {len(df)}")
    
    print(f"Minimum length of DataFrames: {min_length}")
    # Truncate all DataFrames to the minimum length
    for key in dataframes:
        dataframes[key] = dataframes[key].iloc[:min_length]

    combined_data = []

    # Process the DataFrames to calculate metrics
    for i in range(min_length - 1):  # Ignore the last row since it has no next row with averages
        metrics = {}
        metrics['Period End'] = dataframes['LiabilitiesCurrent'].iloc[i]['Period End']

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
        metrics['Revenue'] = dataframes['RevenueFromContractWithCustomerExcludingAssessedTax'].iloc[i]['Value']
        metrics['Gross Profit'] = dataframes['RevenueFromContractWithCustomerExcludingAssessedTax'].iloc[i]['Value'] - dataframes['CostOfGoodsAndServicesSold'].iloc[i]['Value']
        metrics['Cost of Goods Sold'] = dataframes['CostOfGoodsAndServicesSold'].iloc[i]['Value']
        metrics['Basic EPS'] = dataframes['EarningsPerShareBasic'].iloc[i]['Value']
        metrics['Basic EPS'] = dataframes['EarningsPerShareBasic'].iloc[i]['Value']
        metrics['Total Equity'] = dataframes['StockholdersEquity'].iloc[i]['Value']
        metrics['Total Liabilities'] = dataframes['LiabilitiesAndStockholdersEquity'].iloc[i]['Value'] - metrics['Total Equity']
        
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


# In[4]:


def calculate_ratios(file_path, ticker):
    # Call the calculate_financial_metrics function
    df = calculate_financial_metrics(file_path, ticker)
    
    # Check if necessary columns are present
    required_columns = [
        'Current Assets', 'Current Liabilities', 'Revenue', 'Average Receivables', 'Cost of Goods Sold',
        'Average Net PPE', 'Average Total Assets', 'Gross Profit',
        'Operating Income', 'EBT', 'Net Income', 'Effective Tax Rate', 'Average Shareholder Equity',
        'Average Invested Capital'
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        # Print missing columns
        print(f"Missing columns: {', '.join(missing_columns)}")
        raise ValueError(f"Missing columns: {', '.join(missing_columns)}")
    
    # Calculate Ratios
    df['Current'] = df['Current Assets'] / df['Current Liabilities']
    df['Receivables Turnover'] = df['Revenue'] / df['Average Receivables']
    df['Days Sales Outstanding'] = 90 / df['Receivables Turnover']
    df['Fixed Asset Turnover'] = df['Revenue'] / df['Average Net PPE']
    df['Total Asset Turnover'] = df['Revenue'] / df['Average Total Assets']
    df['Gross Margin'] = df['Gross Profit'] / df['Revenue']
    df['Operating Margin'] = df['Operating Income'] / df['Revenue']
    df['Pretax Margin'] = df['EBT'] / df['Revenue']
    df['Net Profit Margin'] = df['Net Income'] / df['Revenue']
    df['Operating Return on Assets'] = df['Operating Income'] / df['Average Total Assets']
    df['Return on Assets'] = df['Net Income'] / df['Average Total Assets']
    df['Return on Equity'] = df['Net Income'] / df['Average Shareholder Equity']
    df['ROIC'] = (df['EBIT'] * (1 - df['Effective Tax Rate'])) / df['Average Invested Capital']
    df['Tax Burden'] = df['Net Income'] / df['EBT']
    df['Interest Burden'] = df['EBT'] / df['EBIT']
    df['EBIT Margin'] = df['EBIT'] / df['Revenue']
    df['Financial Leverage Ratio'] = df['Average Total Assets'] / df['Average Shareholder Equity']
    df['Basic EPS'] = df['Basic EPS']
    df['Debt to Equity'] = df['Total Liabilities'] / df['Total Equity']
    
    # Create a DataFrame with the results
    result_df = df[['Period End', 'Current','Receivables Turnover',
                   'Days Sales Outstanding',
                   'Fixed Asset Turnover','Total Asset Turnover', 'Gross Margin','Operating Margin','Pretax Margin',
                   'Net Profit Margin','Operating Return on Assets', 'Return on Assets', 'Return on Equity', 'Debt to Equity',
                   'ROIC', 'Tax Burden', 'Interest Burden', 'EBIT Margin', 'Financial Leverage Ratio','Basic EPS']].copy()
    
    # Define the output file path
    output_file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\ratios.xlsx"
    if not os.path.exists(output_file_path):
        workbook = Workbook()
        workbook.create_sheet(title=ticker)
        workbook.save(output_file_path)
    
    # Save the result to a new Excel file or update existing one
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        result_df.to_excel(writer, sheet_name=ticker, index=False)
    
    return result_df
# Example usage
#file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\financial_data.xlsx"
#ticker = 'NVDA,0001045810'
#result_df = calculate_ratios(file_path, ticker)
#print(result_df)


# In[5]:


def get_sheet_names(file_path):
    # Load the Excel file
    excel_file = pd.ExcelFile(file_path)
    
    # Get the sheet names
    sheet_names = excel_file.sheet_names
    
    return sheet_names


# In[6]:


def calculate_all(file_path):
    list_of_names = get_sheet_names(file_path)
    
    for name in list_of_names:
        try:
            # Calculate and save ratios for each sheet
            calculate_ratios(file_path, name)
        except Exception as e:
            print(f"Error processing {name}: {e}")
    # Print a completion message
    print('All ratios have been calculated and saved.')


# In[7]:


def add_ppeg(file_path, ticker):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=ticker)
    
    # Check if 'PropertyPlantAndEquipmentGross' exists in the 'Fact' column
    if 'PropertyPlantAndEquipmentGross' in df['Fact'].values:
        return df
    else:
        if ('AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment' in df['Fact'].values and 
        'PropertyPlantAndEquipmentNet' in df['Fact'].values):
        
            # Get the matching rows for each fact
            accumulated_depreciation = df[df['Fact'] == 'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment']
            ppe_net = df[df['Fact'] == 'PropertyPlantAndEquipmentNet']
        
            # Merge the DataFrames on 'Period End'
            merged = pd.merge(accumulated_depreciation, ppe_net, on='Period End', suffixes=('_accumulated', '_net'))
        
            # Calculate the PropertyPlantAndEquipmentGross
            merged['PropertyPlantAndEquipmentGross'] = (
                merged['Value_net'] + merged['Value_accumulated']
            )
        
            # Prepare the new DataFrame row
            new_rows = merged[['Period End']].copy()
            new_rows['Fact'] = 'PropertyPlantAndEquipmentGross'
            new_rows['Value'] = merged['PropertyPlantAndEquipmentGross']
        
            # Append the new rows to the original DataFrame
            updated_df = pd.concat([df, new_rows], ignore_index=True)
        
            # Save the updated DataFrame back to the Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                updated_df.to_excel(writer, sheet_name=ticker, index=False)
    
        return updated_df      


# In[8]:


def remove_COGS(file_path, ticker):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=ticker)
    
    # Check if 'InventoryNet' exists in the 'Fact' column
    if 'CostOfGoodsAndServicesSold' in df['Fact'].values:
        return df
    
    # Check if 'LiabilitiesCurrent' exists
    if 'LiabilitiesCurrent' in df['Fact'].values:
        # Create a DataFrame for rows where 'Fact' is 'LiabilitiesCurrent'
        liabilities_current_df = df[df['Fact'] == 'LiabilitiesCurrent'].copy()
        
        # Create new rows for 'InventoryNet' with the same 'Period End' as 'LiabilitiesCurrent'
        new_rows = liabilities_current_df[['Period End']].copy()
        new_rows['Fact'] = 'CostOfGoodsAndServicesSold'
        new_rows['Value'] = 0
        
        # Append the new rows to the original DataFrame
        updated_df = pd.concat([df, new_rows], ignore_index=True)
        
        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_df.to_excel(writer, sheet_name=ticker, index=False)
        
        return updated_df
    
    else:
        return df


# In[9]:


def remove_interest_expense(file_path, ticker):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=ticker)
    
    # Check if 'InventoryNet' exists in the 'Fact' column
    if 'InterestExpense' in df['Fact'].values:
        return df
    
    # Check if 'LiabilitiesCurrent' exists
    if 'LiabilitiesCurrent' in df['Fact'].values:
        # Create a DataFrame for rows where 'Fact' is 'LiabilitiesCurrent'
        liabilities_current_df = df[df['Fact'] == 'LiabilitiesCurrent'].copy()
        
        # Create new rows for 'InventoryNet' with the same 'Period End' as 'LiabilitiesCurrent'
        new_rows = liabilities_current_df[['Period End']].copy()
        new_rows['Fact'] = 'InterestExpense'
        new_rows['Value'] = 0
        
        # Append the new rows to the original DataFrame
        updated_df = pd.concat([df, new_rows], ignore_index=True)
        
        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_df.to_excel(writer, sheet_name=ticker, index=False)
        
        return updated_df
    
    else:
        return df


# In[10]:


def remove_receivable(file_path, ticker):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path, sheet_name=ticker)
    
    # Check if 'InventoryNet' exists in the 'Fact' column
    if 'AccountsReceivableNetCurrent' in df['Fact'].values:
        return df
    
    # Check if 'LiabilitiesCurrent' exists
    if 'LiabilitiesCurrent' in df['Fact'].values:
        # Create a DataFrame for rows where 'Fact' is 'LiabilitiesCurrent'
        liabilities_current_df = df[df['Fact'] == 'LiabilitiesCurrent'].copy()
        
        # Create new rows for 'InventoryNet' with the same 'Period End' as 'LiabilitiesCurrent'
        new_rows = liabilities_current_df[['Period End']].copy()
        new_rows['Fact'] = 'AccountsReceivableNetCurrent'
        new_rows['Value'] = 0
        
        # Append the new rows to the original DataFrame
        updated_df = pd.concat([df, new_rows], ignore_index=True)
        
        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_df.to_excel(writer, sheet_name=ticker, index=False)
        
        return updated_df
    
    else:
        return df


# In[ ]:




