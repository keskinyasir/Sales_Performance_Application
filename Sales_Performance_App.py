import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
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
excel_file = st.sidebar.file_uploader("Excel dosyasını yükleyin", type=["xlsx", "xls"])

if excel_file:
    # Veri Yükleme
    df_sales = pd.read_excel(excel_file, sheet_name="SATIŞ")
    df_cross = pd.read_excel(excel_file, sheet_name="ÇAPRAZ SATIŞ")
    df_demo = pd.read_excel(excel_file, sheet_name="ILCE DEMOGRAFI")

    # Veri Önişleme
    # Tarih formatı düzeltmesi
    df_sales['YEARMONTH'] = pd.to_datetime(df_sales['YEARMONTH'].astype(str), format='%Y%m')
    df_cross['AY'] = pd.to_datetime(df_cross['AY'].astype(str), format='%Y%m')
    # Başlık
    st.title("Satış Analizi Dashboard")
    st.markdown("Bu uygulama 2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # Sekmeler oluştur
    tab1, tab2, tab3 = st.tabs(["Veri Önizleme", "EDA & Görselleştirme", "Tahmin & Rapor"])

    # -------- VERİ ÖNİZLEME --------
    with tab1:
        st.subheader("1. Ürün1 Satış Verisi")
        st.dataframe(df_sales.head(10))
        st.write(df_sales.describe())

        st.subheader("2. Ürün2 Çapraz Satış Verisi")
        st.dataframe(df_cross.head(10))
        st.write(df_cross.describe())

        st.subheader("3. Demografi Verisi")
        st.dataframe(df_demo.head(10))
        st.write(df_demo.describe())

    # -------- EDA & GÖRSELLEŞTİRME --------
    with tab2:
        # Zaman Serisi
        st.subheader("Aylık Toplam Satış Adedi")
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET', 'URUNHACIM']].sum().reset_index()
        fig_time = px.line(
            df_time,
            x='YEARMONTH',
            y=['URUNADET'],
            title='Zaman Serisi Analizi - Adet'
        )
        st.plotly_chart(fig_time, use_container_width=True)

        st.subheader("Aylık Toplam Satış Hacmi")
        fig_time2 = px.line(
            df_time,
            x='YEARMONTH',
            y=['URUNHACIM'],
            title='Zaman Serisi Analizi - Hacim'
        )
        st.plotly_chart(fig_time2, use_container_width=True)


        # İl Bazlı Harita (Scatter Mapbox)
        st.subheader("Şube Performansı Haritası (İl Bazında)")
        # Dealer -> City ilişkilendirme
        dealer_city = df_cross[['DEALER_CODE', 'CITY']].drop_duplicates()
        dealer_city = dealer_city.dropna()
        dealer_city['DEALER_CODE'] = dealer_city['DEALER_CODE'] // 10
        df_sales_map = pd.merge(df_sales,dealer_city, on='DEALER_CODE', how='inner')
        df_city_sales = df_sales_map.groupby('CITY')['URUNADET'].sum().reset_index()

        # GeoJSON yükleme
        try:
            with open('tr-cities.json', 'r', encoding='utf-8') as f:
                geojson_data = json.load(f)
        except FileNotFoundError:
            st.error("GeoJSON dosyası bulunamadı. 'tr-cities.json' ekleyin.")
            st.stop()

        # Her ilin centroid koordinatını hesapla
        centroids = {}
        for feat in geojson_data['features']:
            name = feat['properties']['name']
            coords = []
            geom = feat['geometry']
            if geom['type'] == 'Polygon':
                rings = geom['coordinates']
            else:
                rings = [ring for poly in geom['coordinates'] for ring in poly]
            for ring in rings:
                coords.extend(ring)
            lons = [c[0] for c in coords]
            lats = [c[1] for c in coords]
            if lats and lons:
                centroids[name] = {'lat': sum(lats)/len(lats), 'lon': sum(lons)/len(lons)}

        # Koordinatları data frame'e ekle
        df_city_sales['lat'] = df_city_sales['CITY'].map(lambda x: centroids.get(x, {}).get('lat', 38))
        df_city_sales['lon'] = df_city_sales['CITY'].map(lambda x: centroids.get(x, {}).get('lon', 35))

        # Scatter Mapbox
        fig_map = px.scatter_mapbox(
            df_city_sales,
            lat='lat',
            lon='lon',
            size='URUNADET',
            color='URUNADET',
            hover_name='CITY',
            hover_data={'URUNADET': True},
            size_max=40,
            zoom=5,
            mapbox_style='open-street-map',
            title='İl Bazında Toplam Ürün1 Satış Adedi'
        )
        st.plotly_chart(fig_map, use_container_width=True)

    # -------- TAHMİN & RAPOR --------
    with tab3:
        st.subheader("2023 Ocak Tahmini – Kanal Bazlı ve Regresörlü Model")
        channels = ['DIJITAL','FIZIKSEL']
        if st.button("Tahmini Hesapla"):
            results = []
            dealer_channel = df_cross[['DEALER_CODE','KANAL']].drop_duplicates()
            df_sales_map = df_sales.merge(dealer_channel, on='DEALER_CODE', how='left')

            for kanal in channels:
                # Ana ürün aylık veri
                main = (
                    df_sales_map[df_sales_map['KANAL']==kanal]
                    .groupby('YEARMONTH')[['URUNADET','URUNHACIM',
                                            'ABONE_YAS_0_3AY','ABONE_YAS_4_12AY',
                                            'ABONE_YAS_1_3YAS','ABONE_YAS_3_YAS']]
                    .sum().reset_index()
                )
                main.columns = ['periode','URUNADET','URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS']
                # Çapraz ürün aylık veri
                cross = (
                    df_cross[df_cross['KANAL']==kanal]
                    .groupby('AY')[['ÇAPRAZ ÜRÜN ADET','5GUNIPTAL','6-45GUNIPTAL']]
                    .sum().reset_index()
                )
                cross.columns = ['periode','CROSS_SALES','CANCEL_5D','CANCEL_6_45D']

                # Tarih dönüştürme
                main['ds'] = pd.to_datetime(main['periode'], format='%Y%m')
                cross['ds'] = pd.to_datetime(cross['periode'], format='%Y%m')
                # Birleştirme
                df_fc = pd.merge(
                    main, cross[['ds','CROSS_SALES','CANCEL_5D','CANCEL_6_45D']],
                    on='ds', how='left'
                )
                # Demografi regresörleri
                demo_mean = df_demo.mean(numeric_only=True)
                for col in demo_mean.index:
                    df_fc[col] = demo_mean[col]

                # Hazırlık
                df_fc.rename(columns={'URUNADET':'y'}, inplace=True)
                regressors = [
                    'URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS',
                    'CROSS_SALES','CANCEL_5D','CANCEL_6_45D'
                ] + list(demo_mean.index)

                # Eğer yeterli veri yoksa son değeri kullan
                if df_fc['y'].dropna().shape[0] < 2:
                    pred_main = df_fc['y'].iloc[-1]
                else:
                    m = Prophet()
                    for reg in regressors:
                        m.add_regressor(reg)
                    m.fit(df_fc[['ds','y'] + regressors])
                    future = m.make_future_dataframe(periods=1, freq='M')
                    last = df_fc.iloc[-1]
                    for reg in regressors:
                        future[reg] = last[reg]
                    forecast = m.predict(future)
                    pred_main = forecast.loc[
                        forecast['ds']==pd.to_datetime('2023-01-31'),'yhat'
                    ].iloc[0]

                results.append({'Kanal':kanal, 'Tahmin (Adet)': round(pred_main)})

            df_res = pd.DataFrame(results)
            st.table(df_res)



        # Rapor indirme
        st.markdown("---")
        st.subheader("Rapor İndir")
        if st.button("PowerPoint Oluştur ve İndir"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Satış Analizi Özet"
            prs.save("sales_report.pptx")
            with open("sales_report.pptx", "rb") as f:
                st.download_button("PowerPoint İndir", f, file_name="sales_analysis.pptx")
