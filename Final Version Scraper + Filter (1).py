#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime


# In[2]:


headers = {'User-Agent': "mafred416@gmail.com"}


# In[3]:


desired_tickers = ['AAPL', 'MSFT', 'AMZN']


# In[6]:


tickers = requests.get(
"https://www.sec.gov/files/company_tickers.json",
headers = headers
)
tickers_json = tickers.json()

def ensure_space_before_quarter(df):
    def correct_period_end(period_end):
        # Check if the 'Q' is preceded by a space
        if 'Q' in period_end:
            parts = period_end.split('Q')
            # If the part before 'Q' does not end with a space, add one
            if not parts[0].endswith(' '):
                return f"{parts[0]} Q{parts[1]}"
        return period_end

    df['Period End'] = df['Period End'].apply(correct_period_end)
    return df

filtered_entries = []
for key, entry in tickers_json.items():
    ticker = entry['ticker']
    
    if ticker in desired_tickers:
        entry['cik_str'] = str(entry['cik_str']).zfill(10)
        filtered_entries.append(entry)
        print(f"Saved {ticker}: CIK {entry['cik_str']}, Title {entry['title']}")

# Fetch company facts
fact_storage = {}

for entry in filtered_entries:
    cik = entry['cik_str']
    ticker = entry['ticker']
    
    companyFacts = requests.get(
        f'https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json',
        headers=headers
    )
    
    if companyFacts.status_code == 200:
        fact_storage[cik] = companyFacts.json()
        print(f"Data for {ticker} ({cik}) retrieved and stored successfully.")
    else:
        print(f"Failed to retrieve data for {ticker} ({cik}). Status code: {companyFacts.status_code}")

# Define terms of interest
terms_of_interest = [
    'AccountsReceivableNet', 'AccountsPayableCurrent', 'AccountsReceivableNetCurrent',
    'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment', 'Assets', 'AssetsCurrent',
    'OperatingIncomeLoss', 'MarketableSecuritiesCurrent', 'CashAndCashEquivalentsAtCarryingValue', 'Cash', 
    'CommonStockSharesOutstanding', 'CostOfGoodsAndServicesSold', 'CostOfRevenue', 'Depreciation',
    'EffectiveIncomeTaxRateContinuingOperations', 'LongTermDebt', 'IncomeTaxExpenseBenefit', 'InterestExpense',
    'InventoryNet', 'LeaseAndRentalExpense', 'LiabilitiesAndStockholdersEquity', 'Liabilities', 'LiabilitiesCurrent',
    'NetCashProvidedByUsedInOperatingActivities', 'NetIncomeLoss', 'OperatingExpenses', 'PropertyPlantAndEquipmentGross',
    'PaymentsOfDividendsCommonStock', 'PropertyPlantAndEquipmentNet', 'StockholdersEquity', 'GrossProfit',
    'WeightedAverageNumberOfSharesOutstandingBasic', 'DebtLongtermAndShorttermCombinedAmount', 
    'SellingGeneralAndAdministrativeExpense'
]

# Filter facts of interest
filtered_fact_storage = {}
for cik, data in fact_storage.items():
    filtered_facts = {'us-gaap': {}}
    us_gaap_facts = data['facts']['us-gaap']
    
    for term in terms_of_interest:
        if term in us_gaap_facts:
            filtered_facts['us-gaap'][term] = us_gaap_facts[term]
    
    filtered_fact_storage[cik] = {'facts': filtered_facts}

# Extract fact values and period end dates
fact_values_with_dates = {}
for cik, data in filtered_fact_storage.items():
    fact_values_with_dates[cik] = {}
    us_gaap_facts = data['facts']['us-gaap']
    
    for fact, details in us_gaap_facts.items():
        fact_values_with_dates[cik][fact] = []
        
        if 'units' in details and 'USD' in details['units']:
            for unit in details['units']['USD']:
                value = unit.get('val')
                fy = unit.get('fy')  # Get fiscal year
                fp = unit.get('fp')  # Get fiscal period
                if value is not None and fy is not None and fp is not None:
                    if fp == 'FY':
                        formatted_period_end = f"{fy} Q4"  # Convert 'FY' to 'Q4'
                    else:
                        formatted_period_end = f"{fy}{fp}"  # Format the date using fiscal year and period
                    fact_values_with_dates[cik][fact].append({
                        'val': value,
                        'end': formatted_period_end  # Use formatted date
                    })

# Save data to Excel
excel_writer = pd.ExcelWriter('financial_data.xlsx', engine='openpyxl')

for cik, facts in fact_values_with_dates.items():
    ticker = None
    for entry in filtered_entries:
        if entry['cik_str'] == cik:
            ticker = entry['ticker']
            break
    
    data_for_sheet = []
    latest_entries = {}  # Dictionary to track the latest values by Fact and Period End
    
    for fact, entries in facts.items():
        for entry in entries:
            period_end = entry['end']
            key = (fact, period_end)
            latest_entries[key] = {
                'Fact': fact,
                'Value': entry['val'],
                'Period End': period_end
            }
    
    # Convert the latest_entries dictionary to a list
    data_for_sheet.extend(latest_entries.values())
    
    df = pd.DataFrame(data_for_sheet)
    
    # Fix spacing issue
    df = ensure_space_before_quarter(df)
    
    # Remove duplicate rows
    df = df.drop_duplicates()
    
    sheet_name = f"{ticker},{cik}"
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

excel_writer.save()
print("Data has been successfully saved to 'financial_data.xlsx'")


# In[7]:


fact_storage['0001018724']['facts']['us-gaap']


# In[ ]:




