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


# In[9]:


#Saved Desired Tickers
#Industrials = desired_tickers = ['GE', 'CAT', 'RTX', 'UNP', 'LMT', 'HON', 'ETN', 'UPS', 'BA', 'DE', 'WM', 'GD', 'TT', 'CTAS', 'CP', 'PH', 'TDG', 'NOC', 'ITW', 'CNI', 'MMM', 'FDX', 'CSX', 'CARR', 'RSG', 'EMR', 'NSC', 'CPRT', 'PCAR', 'URI', 'JCI', 'WCN', 'GWW', 'LHX', 'CMI', 'ODFL', 'PWR', 'AME', 'HWM', 'FAST', 'VRSK', 'EFX', 'OTIS', 'IR', 'HEI', 'XYL', 'VRT', 'ROK', 'HEI.A', 'WAB', 'GPN', 'VLTO', 'AXON', 'DAL', 'DOV', 'HUBB', 'LII', 'BAH', 'BLDR', 'CSL', 'J', 'AER', 'EME', 'WSO.B', 'WSO', 'JBHT', 'EXPD', 'MAS', 'LUV', 'ZTO', 'TXT', 'RBA', 'SWK', 'IEX', 'SNA', 'PNR', 'OC', 'NDSN', 'UAL', 'GGG', 'UHAL', 'UHAL.B', 'POOL', 'ACM', 'XPO', 'CLH', 'FTAI', 'CNH', 'TTEK', 'FIX', 'AOS', 'WMS', 'CHRW', 'CW', 'ALLE', 'BLD', 'ITT', 'NVT', 'RRX', 'HII', 'LECO', 'ULS', 'SAIA', 'WWD', 'FBIN', 'APG', 'ARMK', 'TTC', 'BWXT', 'GNRC', 'CNM', 'KBR', 'ESLT', 'CR', 'MTZ', 'DCI', 'RBC', 'FLR', 'KNX', 'MLI', 'WCC', 'ASR','FCN', 'ATI', 'AIT', 'LTM', 'AYI', 'AAON', 'MIDD', 'SPXC', 'DRS', 'WSC', 'CRS', 'MSA', 'OSK', 'AAL', 'KEX', 'TREX', 'LPX', 'DLB', 'AGCO', 'LOAR', 'ADT', 'WTS', 'FLS', 'LSTR', 'RHI', 'SITE', 'ESAB', 'MOG.A', 'MOG.B', 'R', 'CWST', 'AZEK', 'GXO', 'TKR', 'FSS', 'VMI', 'AVAV', 'BECN', 'MMS', 'ZWS', 'AWI', 'SRCL', 'EXPO', 'CSWI', 'MDU', 'HXL', 'AL', 'GTLS', 'DY', 'TNET', 'GATX', 'BCO', 'SNDR', 'FELE', 'MATX', 'MSM', 'GTES', 'ALK', 'VRRM', 'ACA', 'HRI', 'HAFN', 'SPR', 'ENS', 'AEIS', 'RXO', 'KFY', 'TEX', 'KAI', 'IESC', 'CBZ', 'STRL', 'ABM', 'JOBY', 'UNF', 'NSP', 'BRC', 'MAN', 'ROAD', 'KTOS', 'GMS', 'NPO', 'CXT', 'ATKR', 'MWA', 'GVA', 'GFF', 'HAYW', 'RKLB', 'SKYW', 'ICFI', 'PRIM', 'SEB', 'ATMU', 'REZI', 'HUBG', 'CAR', 'JBT', 'FA', 'TRN', 'BE', 'MGRC', 'ENOV', 'SBLK', 'HNI', 'ARCB', 'AZZ', 'CMPR', 'GOGL', 'CAAP', 'ENR', 'AIR', 'HI', 'WOR', 'WERN', 'MRCY', 'MIR', 'EPAC', 'ALG', 'ASPN', 'ROCK', 'SXI','BWLP', 'B', 'KMT', 'POWL', 'SYM', 'PRG', 'HLMN', 'GEO', 'VSTS', 'TNC', 'HURN', 'VVX', 'JBLU', 'DSGR', 'HEES', 'VSEC', 'NSSC', 'CMRE', 'ENVX', 'CODI', 'MYRG', 'PLUG', 'REVG', 'NMM', 'SCS', 'JBI', 'AMRC', 'SFL', 'DAC', 'NVEE', 'CXW', 'BV', 'GBX', 'CDRE', 'HLIO', 'APOG', 'MRTN', 'DNOW', 'LNN', 'GIC', 'PBI', 'TPC', 'UP', 'KFRC', 'ACHR', 'JELD', 'LZ', 'ARLO', 'CRAI', 'TRNS', 'MEG', 'ULH', 'ATSG', 'HY', 'TILE', 'TGI', 'THR', 'AGX', 'FIP', 'CCEC', 'GRC', 'CECO', 'PCT']
#Tech = ['AAPL', 'MSFT', 'NVDA', 'AVGO', 'ORCL', 'ASML', 'ADBE', 'CRM', 'AMD', 'ACN', 'CSCO', 'TXN', 'QCOM', 'IBM', 'INTU', 'NOW', 'AMAT', 'UBER', 'ARM', 'SONY', 'PANW', 'ADI', 'ADP', 'ANET', 'KLAC', 'MU', 'LRCX', 'FI', 'SHOP', 'INTC', 'DELL', 'APH', 'SNPS', 'MSI', 'CDNS', 'PLTR', 'WDAY', 'CRWD', 'MRVL', 'NXPI', 'ROP', 'FTNT', 'ADSK', 'TTD', 'PAYX', 'TEL', 'MPWR', 'FIS', 'MCHP', 'TEAM', 'FICO', 'SQ', 'DDOG', 'CTSH', 'SNOW', 'IT', 'GLW', 'GRMN', 'HPQ', 'ON', 'APP', 'ZS', 'CDW', 'STM', 'ANSS', 'NET', 'KEYS', 'FTV', 'MSTR', 'SMCI', 'HUBS', 'HPE', 'TYL', 'BR', 'NTAP', 'FSLR', 'GDDY', 'IOT', 'WDC', 'TER', 'CPAY', 'CHKP', 'PTC', 'MDB', 'LDOS', 'ZM', 'STX', 'TDY', 'SSNC', 'VRSN', 'ZBRA', 'SWKS', 'ENTG', 'PSTG', 'ENPH', 'BSY', 'GEN', 'MANH', 'NTNX', 'AKAM', 'DT', 'AZPN', 'TOST', 'TRMB', 'LOGI', 'AFRM', 'OKTA', 'MNDY', 'FLEX', 'JNPR', 'GRAB', 'JKHY', 'JBL', 'CYBR', 'GWRE', 'DOCU', 'COHR','FFIV', 'UI', 'EPAM', 'QRVO', 'CACI', 'NICE', 'ONTO', 'SNX', 'PSN', 'TWLO', 'FFIV', 'UI', 'EPAM', 'QRVO', 'CACI', 'NICE', 'ONTO', 'SNX', 'PSN', 'TWLO', 'DOX', 'WIX', 'DUOL', 'OLED', 'PAYC', 'DAY', 'PCTY', 'FN', 'PCOR', 'DSGX', 'OTEX', 'APPF', 'KVYO', 'CIEN', 'DBX', 'AMKR', 'AUR', 'MKSI', 'MTSI', 'WEX', 'ESTC', 'CRUS', 'ALTR', 'YMM', 'INFA', 'GTLB', 'SPSC', 'S', 'PATH', 'ARW', 'NSIT', 'G', 'CGNX', 'HCP', 'CFLT', 'CVLT', 'SMAR', 'LFUS', 'ALAB', 'CCCS', 'SAIC', 'NOVT', 'LSCC', 'U', 'NVMI', 'VRNS', 'SQSP', 'RBRK', 'BMI', 'PEGA', 'VERX', 'EXLS', 'NXT', 'BILL', 'ST', 'ZETA', 'CRDO', 'QXO', 'FOUR', 'KD', 'CWAN', 'VNT', 'ALGM', 'ACIW', 'SATS', 'CLVT', 'TENB', 'CNXC', 'TSEM', 'AVT', 'OS', 'EEFT', 'RMBS', 'LYFT', 'PI', 'BOX', 'QLYS', 'ITRI', 'RUN', 'ASTS', 'BRZE', 'QTWO', 'BDC', 'WK', 'ASGN', 'BLKB', 'LPL', 'CAMT', 'ALIT', 'LITE', 'SLAB', 'POWI', 'SANM', 'PWSC', 'FORM', 'DXC', 'ZI', 'ACLS', 'CLBT', 'FRSH', 'PLXS', 'IDCC', 'ENV', 'NCNO', 'INTA', 'DOCN','INST', 'DV', 'SITM', 'GDS', 'GBTG', 'ALKT', 'SMTC', 'DIOD', 'SYNA', 'ASAN', 'AGYS', 'ESE', 'BL', 'FROG', 'RNG', 'IPGP', 'AI', 'ALRM', 'LIF', 'PAYO', 'YOU', 'PAY', 'VSH', 'TDC', 'MQ', 'CORZ', 'RELY', 'PRFT', 'WNS', 'PLUS', 'PYCR', 'PRGS', 'OSIS', 'CALX', 'AMBA', 'FIVN', 'KLIC', 'NABL', 'CXM', 'RPD', 'JAMF', 'APPN', 'FLYW', 'VZIO', 'EVTC', 'SWI', 'AVPT', 'SIMO', 'GRND', 'ODD', 'NATL', 'EXTR', 'SPNS', 'VECO', 'SEMR', 'VSAT', 'ROG', 'TTMI', 'EVCM', 'VYX', 'PAR', 'VRNT', 'DFIN', 'CNXN', 'VIAV', 'PD', 'UPBD', 'SPT', 'SOUN', 'IBTA', 'VICR', 'MLNK', 'RAMP', 'UCTT', 'DBD', 'HLIT', 'AVDX', 'PLAB', 'KN', 'RUM', 'IONQ', 'STER', 'BHE', 'NTCT', 'CTS', 'INFN', 'SONO', 'MTTR', 'XRX', 'TWKS', 'BB', 'NOVA', 'SEDG', 'ETWO', 'ADEA', 'CSGS', 'PSFE', 'ZUO', 'WKME', 'MXL', 'COHU', 'CRCT', 'WOLF', 'SCSC', 'PDFS', 'AOSL', 'FORTY', 'TASK', 'SABR', 'ACMR', 'SGH', 'AMPL', 'PGY', 'BELFA', 'DGII', 'GDYN', 'ICHR', 'HIMX', 'ARRY', 'JKS', 'ATEN']


# In[10]:


desired_tickers = ['AAPL', 'MSFT', 'NVDA', 'AVGO', 'ORCL', 'ASML', 'ADBE', 'CRM', 'AMD', 'ACN', 'CSCO', 'TXN', 'QCOM', 'IBM', 'INTU', 'NOW', 'AMAT', 'UBER', 'ARM', 'SONY', 'PANW', 'ADI', 'ADP', 'ANET', 'KLAC', 'MU', 'LRCX', 'FI', 'SHOP', 'INTC', 'DELL', 'APH', 'SNPS', 'MSI', 'CDNS', 'PLTR', 'WDAY', 'CRWD', 'MRVL', 'NXPI', 'ROP', 'FTNT', 'ADSK', 'TTD', 'PAYX', 'TEL', 'MPWR', 'FIS', 'MCHP', 'TEAM', 'FICO', 'SQ', 'DDOG', 'CTSH', 'SNOW', 'IT', 'GLW', 'GRMN', 'HPQ', 'ON', 'APP', 'ZS', 'CDW', 'STM', 'ANSS', 'NET', 'KEYS', 'FTV', 'MSTR', 'SMCI', 'HUBS', 'HPE', 'TYL', 'BR', 'NTAP', 'FSLR', 'GDDY', 'IOT', 'WDC', 'TER', 'CPAY', 'CHKP', 'PTC', 'MDB', 'LDOS', 'ZM', 'STX', 'TDY', 'SSNC', 'VRSN', 'ZBRA', 'SWKS', 'ENTG', 'PSTG', 'ENPH', 'BSY', 'GEN', 'MANH', 'NTNX', 'AKAM', 'DT', 'AZPN', 'TOST', 'TRMB', 'LOGI', 'AFRM', 'OKTA', 'MNDY', 'FLEX', 'JNPR', 'GRAB', 'JKHY', 'JBL', 'CYBR', 'GWRE', 'DOCU', 'COHR','FFIV', 'UI', 'EPAM', 'QRVO', 'CACI', 'NICE', 'ONTO', 'SNX', 'PSN', 'TWLO', 'FFIV', 'UI', 'EPAM', 'QRVO', 'CACI', 'NICE', 'ONTO', 'SNX', 'PSN', 'TWLO', 'DOX', 'WIX', 'DUOL', 'OLED', 'PAYC', 'DAY', 'PCTY', 'FN', 'PCOR', 'DSGX', 'OTEX', 'APPF', 'KVYO', 'CIEN', 'DBX', 'AMKR', 'AUR', 'MKSI', 'MTSI', 'WEX', 'ESTC', 'CRUS', 'ALTR', 'YMM', 'INFA', 'GTLB', 'SPSC', 'S', 'PATH', 'ARW', 'NSIT', 'G', 'CGNX', 'HCP', 'CFLT', 'CVLT', 'SMAR', 'LFUS', 'ALAB', 'CCCS', 'SAIC', 'NOVT', 'LSCC', 'U', 'NVMI', 'VRNS', 'SQSP', 'RBRK', 'BMI', 'PEGA', 'VERX', 'EXLS', 'NXT', 'BILL', 'ST', 'ZETA', 'CRDO', 'QXO', 'FOUR', 'KD', 'CWAN', 'VNT', 'ALGM', 'ACIW', 'SATS', 'CLVT', 'TENB', 'CNXC', 'TSEM', 'AVT', 'OS', 'EEFT', 'RMBS', 'LYFT', 'PI', 'BOX', 'QLYS', 'ITRI', 'RUN', 'ASTS', 'BRZE', 'QTWO', 'BDC', 'WK', 'ASGN', 'BLKB', 'LPL', 'CAMT', 'ALIT', 'LITE', 'SLAB', 'POWI', 'SANM', 'PWSC', 'FORM', 'DXC', 'ZI', 'ACLS', 'CLBT', 'FRSH', 'PLXS', 'IDCC', 'ENV', 'NCNO', 'INTA', 'DOCN','INST', 'DV', 'SITM', 'GDS', 'GBTG', 'ALKT', 'SMTC', 'DIOD', 'SYNA', 'ASAN', 'AGYS', 'ESE', 'BL', 'FROG', 'RNG', 'IPGP', 'AI', 'ALRM', 'LIF', 'PAYO', 'YOU', 'PAY', 'VSH', 'TDC', 'MQ', 'CORZ', 'RELY', 'PRFT', 'WNS', 'PLUS', 'PYCR', 'PRGS', 'OSIS', 'CALX', 'AMBA', 'FIVN', 'KLIC', 'NABL', 'CXM', 'RPD', 'JAMF', 'APPN', 'FLYW', 'VZIO', 'EVTC', 'SWI', 'AVPT', 'SIMO', 'GRND', 'ODD', 'NATL', 'EXTR', 'SPNS', 'VECO', 'SEMR', 'VSAT', 'ROG', 'TTMI', 'EVCM', 'VYX', 'PAR', 'VRNT', 'DFIN', 'CNXN', 'VIAV', 'PD', 'UPBD', 'SPT', 'SOUN', 'IBTA', 'VICR', 'MLNK', 'RAMP', 'UCTT', 'DBD', 'HLIT', 'AVDX', 'PLAB', 'KN', 'RUM', 'IONQ', 'STER', 'BHE', 'NTCT', 'CTS', 'INFN', 'SONO', 'MTTR', 'XRX', 'TWKS', 'BB', 'NOVA', 'SEDG', 'ETWO', 'ADEA', 'CSGS', 'PSFE', 'ZUO', 'WKME', 'MXL', 'COHU', 'CRCT', 'WOLF', 'SCSC', 'PDFS', 'AOSL', 'FORTY', 'TASK', 'SABR', 'ACMR', 'SGH', 'AMPL', 'PGY', 'BELFA', 'DGII', 'GDYN', 'ICHR', 'HIMX', 'ARRY', 'JKS', 'ATEN']
    


# In[11]:


tickers = requests.get(
    "https://www.sec.gov/files/company_tickers.json",
    headers=headers
)
tickers_json = tickers.json()

excel_writer = pd.ExcelWriter('financial_data.xlsx', engine='openpyxl', mode='w')  # Overwrite if exists

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

# Filter desired tickers
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
    'AccountsReceivableNet', 'AccountsReceivableNetCurrent', 'TradeReceivablesHeldForSaleAmount', 'CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization',
    'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment', 'Assets', 'AssetsCurrent', 'IncomeLossFromContinuingOperationsPerBasicShare',
    'OperatingIncomeLoss', 'MarketableSecuritiesCurrent', 'CashAndCashEquivalentsAtCarryingValue', 'Cash','PropertyPlantAndEquipmentOtherAccumulatedDepreciation',
    'CommonStockSharesOutstanding', 'CostOfGoodsAndServicesSold', 'CostOfRevenue', 'Depreciation', 'InterestExpenseRelatedParty'
    'EffectiveIncomeTaxRateContinuingOperations', 'LongTermDebt', 'IncomeTaxExpenseBenefit', 'InterestExpense',
     'LeaseAndRentalExpense', 'LiabilitiesAndStockholdersEquity', 'Liabilities', 'LiabilitiesCurrent', 'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
    'NetCashProvidedByUsedInOperatingActivities', 'NetIncomeLoss', 'OperatingExpenses', 'PropertyPlantAndEquipmentGross','InterestPaid',
    'PaymentsOfDividendsCommonStock', 'PropertyPlantAndEquipmentNet', 'StockholdersEquity', 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents',
     'DebtLongtermAndShorttermCombinedAmount', 'WeightedAverageNumberOfSharesOutstandingBasic',
    'SellingGeneralAndAdministrativeExpense', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'SalesRevenueNet','Revenues',
    'LongTermDebtAndCapitalLeaseObligations','InterestExpenseDebt', 'AccountsAndOtherReceivablesNetCurrent', 'CashCashEquivalentsAndShortTermInvestments',
    'LongTermDebtNoncurrent', 'ProfitLoss','EarningsPerShareBasic','InterestAndDebtExpense',
    'LongTermNotesAndLoans','RevenueFromContractWithCustomerIncludingAssessedTax',
    'CostOfServices', 'InterestIncomeExpenseNet', 'OperatingLeaseLiabilityNoncurrent','PropertyPlantAndEquipmentOther',
    'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest','AccountsReceivableGrossCurrent',
    'PropertyPlantAndEquipmentOwnedAccumulatedDepreciation','ReceivablesNetCurrent','AccountsNotesAndLoansReceivableNetCurrent',
]


missing_us_gaap = []

for cik, data in fact_storage.items():
    if 'us-gaap' not in data.get('facts', {}):
        missing_us_gaap.append(cik)

print("CIKs missing 'us-gaap':", missing_us_gaap)

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
        
        # Check for different types of units
        for unit_type in details.get('units', {}):
            for unit in details['units'][unit_type]:
                value = unit.get('val')
                fy = unit.get('fy')  # Get fiscal year
                fp = unit.get('fp')  # Get fiscal period
                if value is not None and fy is not None and fp is not None:
                    # Determine the period end
                    if fp == 'FY':
                        formatted_period_end = f"{fy} Q4"  # Convert 'FY' to 'Q4'
                    else:
                        formatted_period_end = f"{fy} {fp}"  # Add space between fy and fp
                    # Append value, unit type, and formatted date
                    fact_values_with_dates[cik][fact].append({
                        'val': value,
                        'unit': unit_type,  # Include the unit type
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
    latest_entries = {}  # Dictionary to track the latest values by Fact, Unit, and Period End
    
    for fact, entries in facts.items():
        for entry in entries:
            period_end = entry['end']
            unit = entry['unit']
            key = (fact, unit, period_end)
            latest_entries[key] = {
                'Fact': fact,
                'Value': entry['val'],
                'Unit': unit,  # Include the unit in the Excel data
                'Period End': period_end
            }
    
    # Convert the latest_entries dictionary to a list
    data_for_sheet.extend(latest_entries.values())
    
    df = pd.DataFrame(data_for_sheet)
    
    # Fix spacing issue
    df = ensure_space_before_quarter(df)
    
    # Replace 'CostOfGoodsSold' with 'CostOfGoodsAndServicesSold'
    df['Fact'] = df['Fact'].replace('CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization', 'CostOfGoodsAndServicesSold')
    df['Fact'] = df['Fact'].replace('CostOfRevenue', 'CostOfGoodsAndServicesSold')
    df['Fact'] = df['Fact'].replace('SalesRevenueNet', 'RevenueFromContractWithCustomerExcludingAssessedTax')
    df['Fact'] = df['Fact'].replace('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax')
    df['Fact'] = df['Fact'].replace('LongTermDebtAndCapitalLeaseObligations', 'LongTermDebt')
    df['Fact'] = df['Fact'].replace('InterestExpenseDebt', 'InterestExpense')
    df['Fact'] = df['Fact'].replace('AccountsReceivableNet', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('AccountsAndOtherReceivablesNetCurrent', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('CashCashEquivalentsAndShortTermInvestments', 'CashAndCashEquivalentsAtCarryingValue')
    df['Fact'] = df['Fact'].replace('LongTermDebtNoncurrent', 'LongTermDebt')
    df['Fact'] = df['Fact'].replace('ProfitLoss', 'NetIncomeLoss')
    df['Fact'] = df['Fact'].replace('InterestAndDebtExpense', 'InterestExpense')
    df['Fact'] = df['Fact'].replace('LongTermNotesAndLoans', 'LongTermDebt')
    df['Fact'] = df['Fact'].replace('RevenueFromContractWithCustomerIncludingAssessedTax', 'RevenueFromContractWithCustomerExcludingAssessedTax')
    df['Fact'] = df['Fact'].replace('CostOfServices', 'CostOfGoodsAndServicesSold')
    df['Fact'] = df['Fact'].replace('InterestIncomeExpenseNet', 'InterestExpense')
    df['Fact'] = df['Fact'].replace('OperatingLeaseLiabilityNoncurrent', 'LongTermDebt')
    df['Fact'] = df['Fact'].replace('Cash', 'CashAndCashEquivalentsAtCarryingValue')
    df['Fact'] = df['Fact'].replace('PropertyPlantAndEquipmentOther', 'PropertyPlantAndEquipmentGross')
    df['Fact'] = df['Fact'].replace('IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest', 'OperatingIncomeLoss')
    df['Fact'] = df['Fact'].replace('AccountsReceivableGrossCurrent', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('PropertyPlantAndEquipmentOwnedAccumulatedDepreciation', 'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment')
    df['Fact'] = df['Fact'].replace('ReceivablesNetCurrent', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('AccountsNotesAndLoansReceivableNetCurrent', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'CashAndCashEquivalentsAtCarryingValue')
    df['Fact'] = df['Fact'].replace('InterestExpenseRelatedParty', 'InterestExpense')
    df['Fact'] = df['Fact'].replace('TradeReceivablesHeldForSaleAmount', 'AccountsReceivableNetCurrent')
    df['Fact'] = df['Fact'].replace('IncomeLossFromContinuingOperationsPerBasicShare', 'EarningsPerShareBasic')
    df['Fact'] = df['Fact'].replace('PropertyPlantAndEquipmentOtherAccumulatedDepreciation', 'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment')
    df['Fact'] = df['Fact'].replace('StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest', 'StockholdersEquity')
    df['Fact'] = df['Fact'].replace('InterestPaid', 'InterestExpense')
    
    # Remove duplicate rows
    df = df.drop_duplicates()
    df['Period End'] = df['Period End'].astype(str).str.strip()
    # Drop rows where 'Period End' is 0
    df = df[(df['Period End'] != '0') & (df['Period End'].str.len() >= 7)]
    
    sheet_name = f"{ticker},{cik}"
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

excel_writer.save()
print("Data has been successfully saved to 'financial_data.xlsx'")


# In[27]:


fact_storage['0000040545']['facts']['us-gaap']


# In[8]:


excel_writer = pd.ExcelWriter('financial_data.xlsx', engine='openpyxl')

for cik, facts in fact_values_with_dates.items():
    ticker = None
    for entry in filtered_entries:
        if entry['cik_str'] == cik:
            ticker = entry['ticker']
            break
    
    data_for_sheet = []
    latest_entries = {}  # Dictionary to track the latest values by Fact, Unit, and Period End
    
    for fact, entries in facts.items():
        for entry in entries:
            period_end = entry['end']
            unit = entry['unit']
            key = (fact, unit, period_end)
            latest_entries[key] = {
                'Fact': fact,
                'Value': entry['val'],
                'Unit': unit,  # Include the unit in the Excel data
                'Period End': period_end
            }
    
    # Convert the latest_entries dictionary to a list
    data_for_sheet.extend(latest_entries.values())
    
    df = pd.DataFrame(data_for_sheet)
    
    # Check if 'Period End' column exists
    if 'Period End' not in df.columns:
        print(f"Ticker {ticker} ({cik}) is missing the 'Period End' column.")
        continue  # Skip to the next ticker

    # Fix spacing issue
    df = ensure_space_before_quarter(df)
    
    # Replace 'CostOfGoodsSold' with 'CostOfGoodsAndServicesSold'
    # (same replacement code as before)

    # Remove duplicate rows
    df = df.drop_duplicates()
    df['Period End'] = df['Period End'].astype(str).str.strip()

    # Drop rows where 'Period End' is 0 or less than 7 characters
    df = df[(df['Period End'] != '0') & (df['Period End'].str.len() >= 7)]
    
    # Check again if there are any missing 'Period End' values after processing
    if df.empty or 'Period End' not in df.columns or df['Period End'].isnull().all():
        print(f"Ticker {ticker} ({cik}) is missing all valid 'Period End' values after processing.")
        continue  # Skip to the next ticker
    
    sheet_name = f"{ticker},{cik}"
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

excel_writer.save()
print("Data has been successfully saved to 'financial_data.xlsx'")


# In[ ]:




