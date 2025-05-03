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
        st.subheader("Tahminleme (Prophet + Tüm Regresörler)")
        if st.button("Tahmini Hesapla"):
            # 1) Aylık satış verilerini hazırlama
            monthly_sales = df_sales.groupby('YEARMONTH').agg({
                'URUNADET':'sum', 'URUNHACIM':'sum',
                'ABONE_YAS_0_3AY':'sum','ABONE_YAS_4_12AY':'sum',
                'ABONE_YAS_1_3YAS':'sum','ABONE_YAS_3_YAS':'sum'
            }).reset_index()
            monthly_sales.columns = ['periode','y','URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS']
            monthly_sales['ds'] = pd.to_datetime(monthly_sales['periode'], format='%Y%m')

            # 2) Çapraz satış verilerini aylık ekleme
            cross_month = df_cross.groupby('AY').agg({
                'ÇAPRAZ ÜRÜN ADET':'sum', '5GUNIPTAL':'sum','6-45GUNIPTAL':'sum'
            }).reset_index()
            cross_month.columns = ['periode_cross','cross_sales','cancel_5d','cancel_6_45d']
            cross_month['ds'] = pd.to_datetime(cross_month['periode_cross'], format='%Y%m')

            # 3) Demografi verilerini il bazında ortalama alarak ekleme (genel sabit regresör)
            demo_agg = df_demo.groupby('IL').mean(numeric_only=True)
            demo_global = demo_agg.mean().to_dict()

            # 4) Ana tabloyu birleştirme
            df_fc = pd.merge(monthly_sales, cross_month[['ds','cross_sales','cancel_5d','cancel_6_45d']], on='ds', how='left')
            # Demografi: tüm satırlara sabit değer ekle
            for k,v in demo_global.items():
                df_fc[k] = v

            # Regresör listesi
            regressors = ['URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS','cross_sales','cancel_5d','cancel_6_45d'] + list(demo_global.keys())

            # 5) Prophet modeli oluşturma ve regresör ekleme
            m = Prophet()
            for reg in regressors:
                m.add_regressor(reg)
            m.fit(df_fc[['ds','y'] + regressors])

            # 6) Gelecek dataframe oluşturma ve regresör değerlerini son bilinenle doldurma
            future = m.make_future_dataframe(periods=1, freq='M')
            last_row = df_fc.iloc[-1]
            for reg in regressors:
                future[reg] = last_row[reg]

            # 7) Tahmin ve görselleştirme
            forecast = m.predict(future)
            fig_pred = plot_plotly(m, forecast)
            st.plotly_chart(fig_pred, use_container_width=True)

            # 8) Tahmin sonucunu göster
            next_pred = forecast.loc[forecast['ds']==forecast['ds'].max(),'yhat'].iloc[0]
            st.write(f"2023 Ocak öngörülen satış adedi: {next_pred:.0f}")



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
