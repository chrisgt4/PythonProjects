import urllib

base_url = "http://ichart.finance.yahoo.com/table.csv?s="

def make_url(stock_ticker):
    return base_url + stock_ticker

output_path = "C:\Users\CurtisC\Documents"

def make_filename(stock_ticker, index_nm):
    return output_path + "\\" + index_nm + "\\" + stock_ticker + ".csv"

def pull_stock_data(stock_ticker, index_nm):
    try:
        urllib.urlretrieve(make_url(stock_ticker), make_filename(stock_ticker, index_nm))
    except urllib.ContentTooShortError as e:
        outfile = open(make_filename(stock_ticker, index_nm), "w")
        outfile.write(e.content)
        outfile.close()

urllib.urlretrieve(base_url + "GOOG")

print "data pulled"
