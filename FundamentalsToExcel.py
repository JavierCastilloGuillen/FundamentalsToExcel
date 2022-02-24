import simfin as sf
from simfin.names import *
import pandas as pd
import pandas_datareader as pdr

# Set your API-key for downloading data.
# Replace YOUR_API_KEY with your actual API-key.
sf.set_api_key('Your_SIMFIN_API_key')

# Set the local directory where data-files are stored.
# The dir will be created if it does not already exist.
sf.set_data_dir('simfin_data/')

file = 'DCF.xlsx'  


stock = pd.read_excel(file, sheet_name='Price')
print(stock)
symbol= stock['Symbol'][0]
price = pdr.get_data_yahoo(symbol)['Adj Close'][-1]
data = {'Symbol': symbol,'Price':[price]}
priceData = pd.DataFrame(data=data)

wb = pd.ExcelFile(file)
with pd.ExcelWriter(wb, engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:     

    income = sf.load(dataset='income', variant='annual', market='us',
                  index=[TICKER, REPORT_DATE],
                  parse_dates=[REPORT_DATE, PUBLISH_DATE, RESTATED_DATE])
    income = income.loc[symbol].T
    income = income.reset_index().fillna(0)

    cashflow = sf.load(dataset='cashflow', variant='annual', market='us',
                  index=[TICKER, REPORT_DATE],
                  parse_dates=[REPORT_DATE, PUBLISH_DATE, RESTATED_DATE])
    cashflow = cashflow.loc[symbol].T
    cashflow = cashflow.reset_index().fillna(0)

    balance = sf.load(dataset='balance', variant='annual', market='us',
                  index=[TICKER, REPORT_DATE],
                  parse_dates=[REPORT_DATE, PUBLISH_DATE, RESTATED_DATE])
    balance = balance.loc[symbol].T
    balance = balance.reset_index().fillna(0)
    
    workBook = writer.book
    balance.to_excel(writer, sheet_name='Balance',index=False)
    income.to_excel(writer, sheet_name='Income',index=False)
    cashflow.to_excel(writer, sheet_name='Cashflow',index=False)
    priceData.to_excel(writer, sheet_name='Price', index=False)     
    writer.save()
