import pandas as pd
from pandasgui import show
import yfinance as yf
import os
from datetime import datetime, timedelta

stock = []
FK = []

directory = '/home/gun/Documents/Amele/RaporTarihleri/'
reports = '/home/gun/Documents/ProperReports'
sorted_reports = sorted(os.listdir(reports))
sorted_directory = sorted(os.listdir(directory))

stock_raw = os.path.splitext(sorted_directory[0])
stock.append(stock_raw[0] + '.IS')

file = pd.read_excel('/home/gun/Documents/Amele/RaporTarihleri/ACSEL.xlsx')
raw_report = pd.read_excel('/home/gun/Documents/ProperReports/ACSEL.xlsx')
raw_report.set_index('Değerler', inplace=True)
report = raw_report.T

dates = file.iloc[0].tolist()
formatted_dates = [datetime.strptime(dateStr, '%Y/%m/%d').strftime('%Y-%m-%d') for dateStr in dates[1:]]

ratio_df = pd.DataFrame(index=report.index)

for index, date in enumerate(formatted_dates):

    date_obj = datetime.strptime(date, '%Y-%m-%d')
    next_day = date_obj + timedelta(days=1)
    next_day_str = next_day.strftime('%Y-%m-%d') 

    data = yf.download(stock, start=date, end=next_day_str)['Adj Close']
    data = data.iloc[0]
    fiyat = round(data, 2)

    net_kar_yillik = report['Net Kar/Zarar Yıllık'].iloc[index]
    odenmis_sermaye = report['Ödenmiş Sermaye'].iloc[index]
    hisse_basi_kar = net_kar_yillik/odenmis_sermaye
    hisse_basi_kar = round(hisse_basi_kar, 2)

    fiyat_kazanc = round(fiyat / hisse_basi_kar, 2)
    FK.append(fiyat_kazanc)    
    
if (len(ratio_df.index) > len(FK)):
    ratio_df = ratio_df.iloc[:len(FK)]
    ratio_df['F/K'] = FK

# print(ratio_df)
show(ratio_df)
