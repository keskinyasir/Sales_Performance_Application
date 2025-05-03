import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import plotly.express as px
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split, RandomizedSearchCV
from prophet import Prophet
from prophet.plot import plot_plotly
from pptx import Presentation
import json

# --- Sayfa Yapılandırma ---
st.set_page_config(
    page_title="Satış Analizi Dashboard",
    page_icon="Logo.jpg",
    layout="wide"
)

# --- Sidebar: Logo ve Veri Yükleme ---
st.sidebar.image("Logo.jpg", width=150)
st.sidebar.title("📂 Veri Yükle")
excel_file = st.sidebar.file_uploader("Excel dosyasını yükleyin", type=["xlsx","xls"])

if excel_file:
    # --- Veri Yükleme ---
    df_sales = pd.read_excel(excel_file, sheet_name="SATIŞ")
    df_cross = pd.read_excel(excel_file, sheet_name="ÇAPRAZ SATIŞ")
    df_demo = pd.read_excel(excel_file, sheet_name="ILCE DEMOGRAFI")

    # Tarih formatı
    df_sales['YEARMONTH'] = pd.to_datetime(df_sales['YEARMONTH'].astype(str), format='%Y%m')
    df_cross['AY'] = pd.to_datetime(df_cross['AY'].astype(str), format='%Y%m')

    # Başlık
    st.title("Satış Analizi Dashboard")
    st.markdown("2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # Sekmeler
    tab1, tab2, tab3 = st.tabs(["Veri Önizleme","EDA & Görselleştirme","Tahmin & Rapor"])

    # --- Tab1: Veri Önizleme ---
    with tab1:
        st.subheader("Ürün1 Satış Verisi")
        st.dataframe(df_sales.head(5))
        st.write(df_sales.describe())
        st.subheader("Ürün2 Çapraz Satış Verisi")
        st.dataframe(df_cross.head(5))
        st.write(df_cross.describe())
        st.subheader("Demografi Verisi (İlçe Bazlı)")
        st.dataframe(df_demo.head(5))
        st.write(df_demo.describe())

    # --- Tab2: EDA & Görselleştirme ---
    with tab2:
        st.subheader("Aylık Satış Zaman Serisi - Adet")
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET']].sum().reset_index()
        fig1 = px.line(df_time, x='YEARMONTH', y='URUNADET', title='Aylık Satış Adedi')
        st.plotly_chart(fig1, use_container_width=True)

        st.subheader("Aylık Satış Zaman Serisi - Hacim")
        df_time2 = df_sales.groupby('YEARMONTH')[['URUNHACIM']].sum().reset_index()
        fig2 = px.line(df_time2, x='YEARMONTH', y='URUNHACIM', title='Aylık Satış Hacmi')
        st.plotly_chart(fig2, use_container_width=True)

        st.subheader("Şube Performansı Haritası (İl Bazında)")
        st.info("Önceki kodla oluşturulan harita burada gösterilecek.")

    # --- Tab3: Tahmin & Rapor ---
    with tab3:
        st.subheader("2023 Ocak Tahmini – Kanal Bazlı Rastgele Orman")
        if st.button("Tahmini Hesapla"):
            # Dealer kodu-kanal eşleştirme
            dealer_channel = df_cross[['DEALER_CODE','KANAL']].drop_duplicates()
            df_sales_map = df_sales.merge(dealer_channel, on='DEALER_CODE', how='left')
            channels = ['DIJITAL','FIZIKSEL']
            results = []

            # Demografi ortalama regresörleri
            demo_mean = df_demo.mean(numeric_only=True)

            for kanal in channels:
                # Ana ürün verileri
                main = (df_sales_map[df_sales_map['KANAL']==kanal]
                        .groupby('YEARMONTH')[['URUNADET','URUNHACIM',
                                                'ABONE_YAS_0_3AY','ABONE_YAS_4_12AY',
                                                'ABONE_YAS_1_3YAS','ABONE_YAS_3_YAS']]
                        .sum().reset_index())
                main['ds'] = main['YEARMONTH']

                # Çapraz ürün verileri
                cross = (df_cross[df_cross['KANAL']==kanal]
                         .groupby('AY')[['ÇAPRAZ ÜRÜN ADET','5GUNIPTAL','6-45GUNIPTAL']]
                         .sum().reset_index())
                cross['ds'] = cross['AY']

                # Birleştirme
                df_fc = pd.merge(main, cross[['ds','ÇAPRAZ ÜRÜN ADET','5GUNIPTAL','6-45GUNIPTAL']],
                                 on='ds', how='left').fillna(0)
                # Demografi regresör ekle
                for col, val in demo_mean.items():
                    df_fc[col] = val

                # Özellik ve hedef
                feats_main = ['URUNHACIM','ABONE_YAS_0_3AY','ABONE_YAS_4_12AY',
                              'ABONE_YAS_1_3YAS','ABONE_YAS_3_YAS','ÇAPRAZ ÜRÜN ADET',
                              '5GUNIPTAL','6-45GUNIPTAL'] + list(demo_mean.index)
                X = df_fc[feats_main]
                y_main = df_fc['URUNADET']

                # Eğitim / test ayırımı
                if len(X) < 3:
                    pred_main = y_main.iloc[-1]
                    pred_cross = df_fc['ÇAPRAZ ÜRÜN ADET'].iloc[-1]
                else:
                    X_train, X_val, y_train, y_val = train_test_split(X, y_main, test_size=0.2, random_state=42)
                    # Model ve hiperparametre aralığı
                    model = RandomForestRegressor(random_state=42)
                    param_dist = {
                        'n_estimators': [50,100,200],
                        'max_depth': [None,5,10],
                        'min_samples_split': [2,5,10]
                    }
                    search = RandomizedSearchCV(model, param_distributions=param_dist,
                                                n_iter=5, cv=3, random_state=42)
                    search.fit(X_train, y_train)
                    best = search.best_estimator_
                    # Tahmin ve önem
                    X_next = X.iloc[[-1]]
                    pred_main = best.predict(X_next)[0]
                    importances = pd.Series(best.feature_importances_, index=feats_main)
                    st.write(f"**{kanal} Ana Ürün Özellik Önemleri:**")
                    st.bar_chart(importances.sort_values(ascending=False))

                    # Çapraz ürün tahminleme
                    y_cross = df_fc['ÇAPRAZ ÜRÜN ADET']
                    Xc_train, Xc_val, yc_train, yc_val = train_test_split(X, y_cross, test_size=0.2, random_state=42)
                    search2 = RandomizedSearchCV(RandomForestRegressor(random_state=42),
                                                 param_distributions=param_dist,
                                                 n_iter=5, cv=3, random_state=42)
                    search2.fit(Xc_train, yc_train)
                    best2 = search2.best_estimator_
                    pred_cross = best2.predict(X_next)[0]
                    imp2 = pd.Series(best2.feature_importances_, index=feats_main)
                    st.write(f"**{kanal} Çapraz Ürün Özellik Önemleri:**")
                    st.bar_chart(imp2.sort_values(ascending=False))

                results.append({
                    'Kanal': kanal,
                    'Ana Ürün Öngörü': round(pred_main),
                    'Çapraz Ürün Öngörü': round(pred_cross)
                })

            df_res = pd.DataFrame(results)
            st.table(df_res)

        # Rapor Indirme
        st.markdown("---")
        st.subheader("PowerPoint Raporu Oluştur")
        if st.button("Raporu İndir"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Ocak 2023 Kanal Bazlı & RF Tahmin"
            prs.save("rf_forecast_report.pptx")
            with open("rf_forecast_report.pptx","rb") as f:
                st.download_button("Raporu İndir", f, file_name="rf_forecast_report.pptx")
