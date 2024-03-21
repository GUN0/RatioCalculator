#!/usr/bin/env python3

import pandas as pd
import numpy as np

# from pandasgui import show
import yfinance as yf
import os
from datetime import datetime, timedelta

directory = "/home/gun/Documents/Amele/RaporTarihleri/"
sorted_directory = sorted(os.listdir(directory))


for i in sorted_directory:
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
    ROA = []
    ROCE = []

    stock_raw = os.path.splitext(i)[0]
    stock.append(stock_raw + ".IS")

    file = pd.read_excel(f"/home/gun/Documents/Amele/RaporTarihleri/{i}")
    raw_report = pd.read_excel(f"/home/gun/Documents/ProperReports/{i}")
    raw_report.set_index("Değerler", inplace=True)
    report = raw_report.T

    dates = file.iloc[0].tolist()
    formatted_dates = [
        datetime.strptime(dateStr, "%Y/%m/%d").strftime("%Y-%m-%d")
        for dateStr in dates[1:]
    ]

    ratio_df = pd.DataFrame(index=report.index)

    for index, date in enumerate(formatted_dates):

        date_obj = datetime.strptime(date, "%Y-%m-%d")

        # Check If It's Weekday Or Weekend If So Adjust To Last Friday
        while date_obj.weekday() >= 5:
            days_until_friday = (date_obj.weekday() - 4) % 7
            date_obj -= timedelta(days=days_until_friday)

        date_obj_str = date_obj.strftime("%Y-%m-%d")

        data = yf.download(stock, start=date_obj, end=None)["Close"]
        fiyat = round(data.iloc[0], 2)

        net_kar_yillik = report["Net Kar/Zarar Yıllık"].iloc[index]
        odenmis_sermaye = report["Ödenmiş Sermaye"].iloc[index]
        net_kar_yillik = net_kar_yillik.astype(np.float64)
        odenmis_sermaye = odenmis_sermaye.astype(np.float64)
        hisse_basi_kar = np.divide(
            net_kar_yillik,
            odenmis_sermaye,
            out=np.zeros_like(net_kar_yillik),
            where=(odenmis_sermaye != 0),
        )
        hisse_basi_kar = np.round(hisse_basi_kar, 4)
        fiyat_kazanc = np.divide(
            fiyat, hisse_basi_kar, out=np.zeros_like(fiyat), where=(hisse_basi_kar != 0)
        )
        fiyat_kazanc = np.round(fiyat_kazanc, 2)
        FK.append(fiyat_kazanc)

        piyasa_degeri = fiyat * odenmis_sermaye
        toplam_kaynaklar = report["Toplam Kaynaklar"].iloc[index]
        pd_dd = np.divide(
            piyasa_degeri,
            toplam_kaynaklar,
            out=np.zeros_like(piyasa_degeri),
            where=(toplam_kaynaklar != 0),
        )
        pd_dd = np.round(pd_dd, 2)
        PD_DD.append(pd_dd)

        cari = report["Dönen Varlıklar"].iloc[index]
        kisa_vadeli_borclar = report["Kısa Vadeli Borçlar"].iloc[index]
        cari = cari.astype(np.float64)
        kisa_vadeli_borclar = kisa_vadeli_borclar.astype(np.float64)
        cari_kisa = np.divide(
            cari,
            kisa_vadeli_borclar,
            out=np.zeros_like(cari),
            where=(kisa_vadeli_borclar != 0),
        )
        cari_kisa = np.round(cari_kisa, 2)
        CARI_ORAN.append(cari_kisa)

        uzun_vadeli_borclar = report["Uzun Vadeli Borçlar"].iloc[index]
        toplam_borclar = uzun_vadeli_borclar + kisa_vadeli_borclar
        uzun_vadeli_borclar = uzun_vadeli_borclar.astype(np.float64)
        toplam_borclar = toplam_borclar.astype(np.float64)
        kaldirac = np.divide(
            toplam_borclar,
            toplam_kaynaklar,
            out=np.zeros_like(toplam_borclar),
            where=(toplam_kaynaklar != 0),
        )
        kaldirac = np.round(kaldirac, 4)
        KALDIRAC_ORANI.append(kaldirac)

        brut_ceyrek = report["Brüt Kar/Zarar Çeyreklik"].iloc[index]
        satis_ceyrek = report["Satış Gelirleri Çeyreklik"].iloc[index]
        brut_ceyrek = brut_ceyrek.astype(np.float64)
        satis_ceyrek = satis_ceyrek.astype(np.float64)
        brut_kar_marji_ceyrek = np.divide(
            brut_ceyrek,
            satis_ceyrek,
            out=np.zeros_like(brut_ceyrek),
            where=(satis_ceyrek != 0),
        )
        brut_kar_marji_ceyrek = np.round(brut_kar_marji_ceyrek, 4)
        BRUT_KAR_MARJI_CEYREK.append(brut_kar_marji_ceyrek)

        brut_yil = report["Brüt Kar/Zarar Yıllık"].iloc[index]
        satis_yil = report["Satış Gelirleri Yıllık"].iloc[index]
        brut_yil = brut_yil.astype(np.float64)
        satis_yil = satis_yil.astype(np.float64)
        brut_kar_marji_yil = np.divide(
            brut_yil, satis_yil, out=np.zeros_like(brut_yil), where=(satis_yil != 0)
        )
        brut_kar_marji_yil = np.round(brut_kar_marji_yil, 4)
        BRUT_KAR_MARJI_YIL.append(brut_kar_marji_yil)

        net_kar_ceyrek = report["Net Kar/Zarar Çeyreklik"].iloc[index]
        net_kar_ceyrek = net_kar_ceyrek.astype(np.float64)
        net_kar_marji_ceyrek = np.divide(
            net_kar_ceyrek,
            satis_ceyrek,
            out=np.zeros_like(net_kar_ceyrek),
            where=(satis_ceyrek != 0),
        )
        net_kar_marji_ceyrek = np.round(net_kar_marji_ceyrek, 4)
        NET_KAR_MARJI_CEYREK.append(net_kar_marji_ceyrek)

        net_kar_marji_yil = np.divide(
            net_kar_yillik,
            satis_yil,
            out=np.zeros_like(net_kar_yillik),
            where=(satis_yil != 0),
        )
        net_kar_marji_yil = np.round(net_kar_marji_yil, 4)
        NET_KAR_MARJI_YIL.append(net_kar_marji_yil)

        ozkaynak_ortalama = report["Özkaynaklar (ORTALAMA)"].iloc[index]
        ozkaynak_ortalama = ozkaynak_ortalama.astype(np.float64)
        ozkaynak_karliligi = np.divide(
            net_kar_yillik,
            ozkaynak_ortalama,
            out=np.zeros_like(net_kar_yillik),
            where=(ozkaynak_ortalama != 0),
        )
        ozkaynak_karliligi = np.round(ozkaynak_karliligi, 4)
        OZKAYNAK_KARLILIGI.append(ozkaynak_karliligi)

        ortalama_kaynaklar = report["Ortalama Kaynaklar"].iloc[index]
        ortalama_kaynaklar = ortalama_kaynaklar.astype(np.float64)
        roa = np.divide(
            net_kar_yillik,
            ortalama_kaynaklar,
            out=np.zeros_like(net_kar_yillik),
            where=(ortalama_kaynaklar != 0),
        )
        roa = np.round(roa, 5)
        ROA.append(roa)

        ar_ge_giderleri = report["Araştırma ve Geliştirme Giderleri (-)"].iloc[index]
        ar_ge_giderleri = ar_ge_giderleri.astype(np.float64)
        pazarlama_satis_dagitim_giderleri = report[
            "Pazarlama, Satış ve Dağıtım Giderleri (-)"
        ].iloc[index]
        pazarlama_satis_dagitim_giderleri = pazarlama_satis_dagitim_giderleri.astype(
            np.float64
        )
        genel_yonetim_giderleri = report["Genel Yönetim Giderleri (-)"].iloc[index]
        genel_yonetim_giderleri = genel_yonetim_giderleri.astype(np.float64)
        diger_faliyet_giderleri = report["Diğer Faaliyet Giderleri (-)"].iloc[index]
        diger_faliyet_giderleri = diger_faliyet_giderleri.astype(np.float64)
        diger_faliyet_gelirleri = report["Diğer Faaliyet Gelirleri"].iloc[index]
        diger_faliyet_gelirleri = diger_faliyet_gelirleri.astype(np.float64)
        ebit = (
            brut_yil
            + ar_ge_giderleri
            + pazarlama_satis_dagitim_giderleri
            + genel_yonetim_giderleri
            + diger_faliyet_giderleri
            + diger_faliyet_gelirleri
        )
        toplam_varliklar = report["Toplam Varlıklar"].iloc[index]
        toplam_varliklar = toplam_varliklar.astype(np.float64)
        capital_employed = toplam_varliklar - kisa_vadeli_borclar
        roce = np.divide(
            ebit,
            capital_employed,
            out=np.zeros_like(ebit),
            where=(capital_employed != 0),
        )
        roce = np.round(roce, 5)
        ROCE.append(roce)

    if len(ratio_df.index) > len(FK):
        ratio_df = ratio_df.iloc[: len(FK)]

        ratio_df["F/K"] = FK
        ratio_df["PD/DD"] = PD_DD
        ratio_df["Cari Oran"] = CARI_ORAN
        ratio_df["Kaldıraç Oranı"] = KALDIRAC_ORANI
        ratio_df["Brüt Kar Marjı Çeyreklik"] = BRUT_KAR_MARJI_CEYREK
        ratio_df["Brüt Kar Marjı Yıllık"] = BRUT_KAR_MARJI_YIL
        ratio_df["Net Kar Marjı Çeyreklik"] = NET_KAR_MARJI_CEYREK
        ratio_df["Net Kar Marjı Yıllık"] = NET_KAR_MARJI_YIL
        ratio_df["Özkaynak Karlılığı"] = OZKAYNAK_KARLILIGI
        ratio_df["ROA"] = ROA
        ratio_df["ROCE"] = ROCE

    # show(ratio_df)
    ratio_df.to_excel(
        "/home/gun/Documents/CalculatedRatios/{}.xlsx".format(stock_raw), index=True
    )
