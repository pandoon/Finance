#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import pandas as pd
import matplotlib.pyplot as plt
import os


# In[2]:


headers = {'User-Agent': "mafred416@gmail.com"}


# In[3]:


desired_tickers = ['AAPL', 'MSFT', 'AMZN']


# In[4]:


# Retrieve tickers JSON
tickers = requests.get(
"https://www.sec.gov/files/company_tickers.json",
headers = headers
)
tickers_json = tickers.json()

filtered_entries = []

for key, entry in tickers_json.items():
    ticker = entry['ticker']
    
    if ticker in desired_tickers:
        entry['cik_str'] = str(entry['cik_str']).zfill(10)
        filtered_entries.append(entry)
        print(f"Saved {ticker}: CIK {entry['cik_str']}, Title {entry['title']}")

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

terms_of_interest = [
    'AccountsReceivableNet','AccountsPayableCurrent',
    'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment','Assets','AssetsCurrent',
    'OperatingIncomeLoss','AvailableForSaleSecurities','CashAndCashEquivalentsAtCarryingValue','Cash','CommonStockSharesOutstanding',
    'CashAndCashEquivalentsAtCarryingValue','CostOfGoodsSold','CostOfRevenue','Depreciation','EffectiveIncomeTaxRateContinuingOperations',
    'LongTermDebt','IncomeTaxExpenseBenefit','InterestExpense','InventoryNet','LeaseAndRentalExpense','LiabilitiesAndStockholdersEquity',
    'Liabilities','LiabilitiesCurrent', 'NetCashProvidedByUsedInOperatingActivities','NetIncomeLoss','OperatingExpenses',
    'OperatingIncomeLoss','PaymentsOfDividendsCommonStock','PropertyPlantAndEquipmentNet','Revenues','StockholdersEquity',
    'WeightedAverageNumberOfSharesOutstandingBasic','DebtLongtermAndShorttermCombinedAmount',
]

filtered_fact_storage = {}

for cik, data in fact_storage.items():
    filtered_facts = {'us-gaap': {}}
    us_gaap_facts = data['facts']['us-gaap']
    
    for term in terms_of_interest:
        if term in us_gaap_facts:
            filtered_facts['us-gaap'][term] = us_gaap_facts[term]
    
    filtered_fact_storage[cik] = {'facts': filtered_facts}

fact_values_with_dates = {}

for cik, data in filtered_fact_storage.items():
    fact_values_with_dates[cik] = {}
    us_gaap_facts = data['facts']['us-gaap']
    
    for fact, details in us_gaap_facts.items():
        fact_values_with_dates[cik][fact] = []
        
        if 'units' in details and 'USD' in details['units']:
            for unit in details['units']['USD']:
                value = unit.get('val')
                filed_date = unit.get('filed')
                if value is not None and filed_date is not None:
                    fact_values_with_dates[cik][fact].append({
                        'val': value,
                        'filed': filed_date
                    })

excel_writer = pd.ExcelWriter('financial_data.xlsx', engine='openpyxl')

for cik, facts in fact_values_with_dates.items():
    ticker = None
    for entry in filtered_entries:
        if entry['cik_str'] == cik:
            ticker = entry['ticker']
            break
    
    data_for_sheet = []
    
    for fact, entries in facts.items():
        for entry in entries:
            data_for_sheet.append({
                'Fact': fact,
                'Value': entry['val'],
                'Filed Date': entry['filed']
            })
    
    df = pd.DataFrame(data_for_sheet)
    sheet_name = f"{ticker},{cik}"
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

excel_writer.save()
print("Data has been successfully saved to 'financial_data.xlsx'")


# In[ ]:




