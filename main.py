import yfinance as yf
import pandas as pd
import urllib.request
from datetime import datetime

tickers = pd.read_csv("tickers.csv", header = None)[0]

stock_variations = []

for ticker in tickers:
    x = yf.Ticker(ticker)
    hist = x.history(period='1d')
    hist.reset_index(drop=True)
    var = (hist['Close'].values - hist['Open'].values)/hist['Open'].values
    try:
        stock_variations.append([ticker, var[0]])
    except: 
        stock_variations.append([ticker, var])

df = pd.DataFrame(stock_variations, columns=['Ticker', 'Variation'])
df = df.sort_values(by='Variation', ascending=False)

best_stock = df.iloc[0]['Ticker']
best_var = str(round(df.iloc[0]['Variation']*100,2)) + '%'
worst_stock = df.iloc[-1]['Ticker']
worst_var = str(round(df.iloc[-1]['Variation']*100,2)) + '%'

best_news = yf.Ticker(best_stock).news[0]
best_title = best_news['title']
best_link = best_news['link']
best_publisher = best_news['publisher']

worst_news = yf.Ticker(worst_stock).news[0]
worst_title = worst_news['title']
worst_link = worst_news['link']
worst_publisher = worst_news['publisher']

today = datetime.today().strftime('%Y-%m-%d')
best_image_url = best_stock.news[0]['thumbnail']['resolutions'][0]['url']
worst_image_url = worst_stock.news[0]['thumbnail']['resolutions'][0]['url']
urllib.request.urlretrieve(best_image_url, today + " best.png") 
urllib.request.urlretrieve(worst_image_url, today + " worst.png")

# print('Best Stock Movement:')
# print(best_stock + " : " + str(round(best_var*100,2)) + '%')
# print(best_news['title'])
# print(best_news['link'])
# print(best_news['publisher'])
# print('')
# print('Worst Stock Movement:')
# print(worst_stock + " : " + str(round(worst_var*100,2)) + '%')
# print(worst_news['title'])
# print(worst_news['link'])
# print(worst_news['publisher'])