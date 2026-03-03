import yfinance as yf
import xlwings as xw

def fetch_and_dump_data():
    try:
        wb = xw.Book('Valuation_model.xlsm')
        sheet_dump = wb.sheets ['Data_dump']
    except:
        print("Error!Excel file is not open or file name is wrong!")
        return

#Take ticker input:
    ticker_symbol = wb.sheets ['Dashboard'].range('Ticker_input').value
    if not ticker_symbol:
        print('Cell is empty!')
        return

    ticker = yf.Ticker(ticker_symbol)

#Clear any old figures in data dump sheet
    sheet_dump.clear_contents()
    
#Take WACC data for DCF model:
    tnx = yf.Ticker("^TNX")
    rf_rate = tnx.history(period="1d")['Close'].iloc[-1] / 100
    
    info = ticker.info
    wacc_data = [["--- WACC INPUTS ---", ""],
                 ["Market Cap", info.get('marketCap',0)],
                 ["Total Debt", info.get('totalDebt',0)],
                 ["Beta", info.get('beta',1)],
                 ["Risk-free rate", rf_rate],
                 ["Interest Expense", ticker.financials.loc['Interest Expense'].iloc[0] if 'Interest Expense' in ticker.financials.index else 0]]
    sheet_dump.range('K1').value = wacc_data


#Dump Income statement
    sheet_dump.range('A1').value = "--- INCOME STATEMENT ---"
    sheet_dump.range('A2').value = ticker.financials

#Find the next available row to dump Balance Sheet
    last_row = sheet_dump.range ('A' + str (sheet_dump.cells.last_cell.row)).end('up').row

    sheet_dump.range(f'A{last_row + 3}').value = "--- BALANCE SHEET ---"
    sheet_dump.range(f'A{last_row + 4}').value = ticker.balance_sheet

#Find the next available row to dump Cash flow statement
    last_row = sheet_dump.range('A' + str(sheet_dump.cells.last_cell.row)).end('up').row

    sheet_dump.range(f'A{last_row + 3}').value = "--- CASH FLOW ---"
    sheet_dump.range(f'A{last_row + 4}').value = ticker.cashflow

    wb.save()
    print(f"Finish data dumping for {ticker_symbol}!")

if __name__ == "__main__":
    fetch_and_dump_data()