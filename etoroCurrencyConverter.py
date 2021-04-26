import datetime
import pandas as pd
import timedelta
from currency_converter import CurrencyConverter

c = CurrencyConverter()

df = pd.read_excel('eToroReport.xlsx', sheet_name='Closed Positions')

print(df)
format_str = '%d/%m/%Y'  # The format

dfOut = pd.DataFrame(columns=['Action', 'Dolar Profit', 'Close Date', 'Convert date', 'Euro Profit'])
for index, row in df.iterrows():
    action = row['Action']
    closeDate = row['Close Date']
    profit = row['Profit']
    x = closeDate.split()
    dtObj = datetime.datetime.strptime(x[0], format_str)

    hasConversion = False
    while hasConversion == False:
        try:
            convertedProfit = c.convert(profit, 'USD', 'EUR', date=dtObj)
            hasConversion = True
        except:
            print('Searching new Date.')
            dtObj = dtObj - datetime.timedelta(1)

    dfOut.loc[index] = [action, profit, closeDate, dtObj, convertedProfit]
    #print(dfOut)
    print(row['Action'], row['Profit'], row['Close Date'], convertedProfit)
print(dfOut)

dfOut.to_excel('processedReport.xlsx')
