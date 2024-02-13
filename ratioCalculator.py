import pandas as pd
import numpy as np
from pandasgui import show
import yfinance as yf
import os
from datetime import datetime, timedelta

directory = '/home/gun/Documents/Amele/RaporTarihleri/'
sorted_directory = sorted(os.listdir(directory))


for i in sorted_directory:
    print(i)
    stock = []
    FK = []
    PD_DD = []
    CARI_ORAN = []
    KALDIRAC_ORANI = []
    BRUT_KAR_MARJI_CEYREK = []
    BRUT_KAR_MARJI_YIL = []
    NET_KAR_MARJI_CEYREK = []
    NET_KAR_MARJI_YIL = []
    OZKAYNAK_KARLILIGI = []

    stock_raw = os.path.splitext(i)[0]
    stock.append(stock_raw + '.IS')

    file = pd.read_excel(f'/home/gun/Documents/Amele/RaporTarihleri/{i}')
    raw_report = pd.read_excel(f'/home/gun/Documents/ProperReports/{i}')
    raw_report.set_index('Değerler', inplace=True)
    report = raw_report.T

    dates = file.iloc[0].tolist()
    formatted_dates = [datetime.strptime(dateStr, '%Y/%m/%d').strftime('%Y-%m-%d') for dateStr in dates[1:]]

    ratio_df = pd.DataFrame(index=report.index)

    for index, date in enumerate(formatted_dates):

        date_obj = datetime.strptime(date, '%Y-%m-%d')

        # Check If It's Weekday Or Weekend If So Adjust To Last Friday
        while date_obj.weekday() >= 5:
            days_until_friday = (date_obj.weekday() - 4) % 7
            date_obj -= timedelta(days=days_until_friday)

        date_obj_str =  date_obj.strftime('%Y-%m-%d') 

        data = yf.download(stock, start=date_obj, end=None)['Close']
        fiyat = round(data.iloc[0], 2)
        print(date_obj_str)
        print(fiyat)

        net_kar_yillik = report['Net Kar/Zarar Yıllık'].iloc[index]
        odenmis_sermaye = report['Ödenmiş Sermaye'].iloc[index]
        hisse_basi_kar = np.divide(net_kar_yillik, odenmis_sermaye)
        if np.isinf(hisse_basi_kar):
            hisse_basi_kar = 0
        hisse_basi_kar = round(hisse_basi_kar, 2)
        fiyat_kazanc = np.divide(fiyat, hisse_basi_kar)
        if np.isinf(fiyat_kazanc):
            fiyat_kazanc = 0
        fiyat_kazanc = round(fiyat_kazanc, 2)
        FK.append(fiyat_kazanc)    

        piyasa_degeri = fiyat * odenmis_sermaye
        toplam_kaynaklar = report['Toplam Kaynaklar'].iloc[index]
        pd_dd = round(piyasa_degeri / toplam_kaynaklar, 2)
        PD_DD.append(pd_dd)

        cari = report['Dönen Varlıklar'].iloc[index]
        kisa_vadeli_borclar = report['Kısa Vadeli Borçlar'].iloc[index]
        cari_kisa = round((cari / kisa_vadeli_borclar), 2)
        CARI_ORAN.append(cari_kisa)

        uzun_vadeli_borclar = report['Uzun Vadeli Borçlar'].iloc[index]
        toplam_borclar = uzun_vadeli_borclar + kisa_vadeli_borclar
        kaldirac = toplam_borclar / toplam_kaynaklar
        kaldirac = round(kaldirac, 3)
        # kaldirac_yuzde = "{:.1%}".format(kaldirac)
        # KALDIRAC_ORANI.append(kaldirac_yuzde)
        KALDIRAC_ORANI.append(kaldirac)

        brut_ceyrek = report['Brüt Kar/Zarar Çeyreklik'].iloc[index]
        satis_ceyrek = report['Satış Gelirleri Çeyreklik'].iloc[index]
        brut_kar_marji_ceyrek = brut_ceyrek / satis_ceyrek
        brut_kar_marji_ceyrek = round(brut_kar_marji_ceyrek, 3)
        # brut_kar_marji_ceyrek_yuzde = "{:.1%}".format(brut_kar_marji_ceyrek)
        # BRUT_KAR_MARJI_CEYREK.append(brut_kar_marji_ceyrek_yuzde)
        BRUT_KAR_MARJI_CEYREK.append(brut_kar_marji_ceyrek)

        brut_yil = report['Brüt Kar/Zarar Yıllık'].iloc[index]
        satis_yil = report['Satış Gelirleri Yıllık'].iloc[index]
        brut_kar_marji_yil = brut_yil / satis_yil
        brut_kar_marji_yil = round(brut_kar_marji_yil, 3)
        # brut_kar_marji_yil_yuzde = "{:.1%}".format(brut_kar_marji_yil)
        # BRUT_KAR_MARJI_YIL.append(brut_kar_marji_yil_yuzde)
        BRUT_KAR_MARJI_YIL.append(brut_kar_marji_yil)
        
        net_kar_ceyrek = report['Net Kar/Zarar Çeyreklik'].iloc[index]
        net_kar_marji_ceyrek = net_kar_ceyrek / satis_ceyrek
        net_kar_marji_ceyrek = round(net_kar_marji_ceyrek, 3)
        # net_kar_marji_ceyrek_yuzde = "{:.1%}".format(net_kar_marji_ceyrek)
        # NET_KAR_MARJI_CEYREK.append(net_kar_marji_ceyrek_yuzde)
        NET_KAR_MARJI_CEYREK.append(net_kar_marji_ceyrek)

        net_kar_marji_yil = net_kar_yillik / satis_yil
        net_kar_marji_yil = round(net_kar_marji_yil, 3)
        # net_kar_marji_yil_yuzde = "{:.1%}".format(net_kar_marji_yil)
        # NET_KAR_MARJI_YIL.append(net_kar_marji_yil_yuzde)
        NET_KAR_MARJI_YIL.append(net_kar_marji_yil)

        ozsermaye_ortalama = report['Özkaynaklar (ORTALAMA)'].iloc[index]
        ozkaynak_karliligi = net_kar_yillik / ozsermaye_ortalama
        ozkaynak_karliligi = round(ozkaynak_karliligi, 3)
        # ozkaynak_karliligi_yuzde = "{:.1%}".format(ozkaynak_karliligi)
        # OZKAYNAK_KARLILIGI.append(ozkaynak_karliligi_yuzde)
        OZKAYNAK_KARLILIGI.append(ozkaynak_karliligi)

    if (len(ratio_df.index) > len(FK)):
        ratio_df = ratio_df.iloc[:len(FK)]

        ratio_df['F/K'] = FK
        ratio_df['PD/DD'] = PD_DD
        ratio_df['Cari Oran'] = CARI_ORAN
        ratio_df['Kaldıraç Oranı'] = KALDIRAC_ORANI
        ratio_df['Brüt Kar Marjı Çeyreklik'] = BRUT_KAR_MARJI_CEYREK
        ratio_df['Brüt Kar Marjı Yıllık'] = BRUT_KAR_MARJI_YIL
        ratio_df['Net Kar Marjı Çeyreklik'] = NET_KAR_MARJI_CEYREK
        ratio_df['Net Kar Marjı Yıllık'] = NET_KAR_MARJI_YIL
        ratio_df['Özkaynak Karlılığı'] = OZKAYNAK_KARLILIGI

    # show(ratio_df)
    ratio_df.to_excel('/home/gun/Documents/CalculatedRatios/{}.xlsx'.format(stock_raw), index=True)
