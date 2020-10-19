import yfinance as yf

def get_quote(stock_code):
    ticker = yf.Ticker(stock_code)
    value = float(ticker.history().iloc[[-1]]["Close"])
    return value

if __name__ == "__main__":
    stock_code = "GOOG"
    stock_price = get_quote(stock_code)
    print(stock_code, stock_price)

    from excel_util import get_excel_row_data
    holdings = get_excel_row_data("D:\\holdings.xlsx", [0,1,2])
    for holding in holdings:
        print(holding[0], holding[1], get_quote(holding[1]))