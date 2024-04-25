import pandas as pd
import yfinance as yf
import os

pd.set_option('display.max_columns',None)
pd.set_option('display.max_rows',None)

def generateReport(subject):
    
    ticker = yf.Ticker(subject)

    companyOfficers = pd.DataFrame(pd.DataFrame(ticker.info)['companyOfficers'].to_list())
    companyOfficers = companyOfficers.rename(columns = {'name':"Officer's name"})
    companyOfficers.name = 'companyOfficers'

    del ticker.info['companyOfficers']

    info = pd.DataFrame(ticker.info, index = [0])
    info.name = 'info'

    incomeStatement = pd.DataFrame(ticker.financials)
    incomeStatement.columns = incomeStatement.columns.strftime('%Y/%m/%d')
    incomeStatement.name = 'incomeStatement'

    balanceSheet = pd.DataFrame(ticker.balancesheet)
    balanceSheet.columns = balanceSheet.columns.strftime('%Y/%m/%d')
    balanceSheet.name = 'balanceSheet'

    cashflow = pd.DataFrame(ticker.cashflow)
    cashflow.columns = cashflow.columns.strftime('%Y/%m/%d')
    cashflow.name = 'cashflow'

    # institutionalHolders = pd.DataFrame(ticker.institutional_holders)
    # institutionalHolders = institutionalHolders.set_index('Date Reported')
    # institutionalHolders.index = institutionalHolders.index.strftime('%Y/%m/%d')
    # institutionalHolders.name = 'institutionalHolders'

    # mutualfundHolders = pd.DataFrame(ticker.mutualfund_holders)
    # mutualfundHolders = mutualfundHolders.set_index('Date Reported')
    # mutualfundHolders.index = mutualfundHolders.index.strftime('%Y/%m/%d')
    # mutualfundHolders.name = 'mutualfundHolders'

    # majorHolders = pd.DataFrame(ticker.major_holders)
    # majorHolders.name = 'majorHolders'

    # insiderTransactions = pd.DataFrame(ticker.insider_transactions)
    # insiderTransactions = insiderTransactions.set_index('Start Date')
    # insiderTransactions.index = insiderTransactions.index.strftime('%Y/%m/%d')
    # insiderTransactions.name = 'insiderTransactions'

    actions = pd.DataFrame(ticker.actions.iloc[::-1])
    actions.index = actions.index.strftime('%Y/%m/%d')
    actions.name = 'actions'

    grades = pd.DataFrame(ticker.upgrades_downgrades)
    grades.index = grades.index.strftime('%Y/%m/%d')
    grades.name = 'grades'

    dfs = [info, companyOfficers, incomeStatement, balanceSheet, cashflow, actions, grades]

    with pd.ExcelWriter(f'C:/Users/{subject}_report.xlsx') as writer:

        for df in dfs:
            df.to_excel(writer, sheet_name = df.name, na_rep = 'NaN')
            writer.sheets[df.name].autofit()

    print('Report generated successfully!')
    os.startfile(f'C:/Users/{subject}_report.xlsx')

    return subject

print('Please type in the name:')
generateReport(input())
