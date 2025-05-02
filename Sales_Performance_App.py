import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from prophet import Prophet
from prophet.plot import plot_plotly
from pptx import Presentation
import json
import folium
import streamlit.components.v1 as components

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


    # Veri önişleme
    # Convert to datetime
    df_sales['YEARMONTH'] = pd.to_datetime(df_sales['YEARMONTH'].astype(str), format='%Y%m')

    # Başlık
    st.title("Satış Analizi Dashboard")
    st.markdown("Bu uygulama 2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # --- Sekmeler ---
    tab1, tab2, tab3 = st.tabs(["Veri Önizleme", "EDA & Görselleştirme", "Tahmin & Rapor"])

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

    with tab2:
        st.subheader("Aylık Toplam Satış Adedi")
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET', 'URUNHACIM']].sum().reset_index()
        fig_time = px.line(
            df_time,
            x='YEARMONTH',
            y=['URUNADET'],
            title='Zaman Serisi Analizi - Adet'
)
        st.plotly_chart(fig_time, use_container_width=True)

        st.subheader("Aylık Toplam Satış Tutarı")
        fig_time2 = px.line(
            df_time,
            x='YEARMONTH',
            y=['URUNHACIM'],
            title='Zaman Serisi Analizi - Hacim'
)

        
        st.plotly_chart(fig_time2, use_container_width=True)

        st.subheader("Şube Performansı Haritası (Folium ile)")
        dealer_city = df_cross[['DEALER_CODE', 'CITY']].drop_duplicates()
        df_sales_map = df_sales.merge(dealer_city, on='DEALER_CODE', how='left')
        df_city_sales = df_sales_map.groupby('CITY')['URUNADET'].sum().reset_index()

        try:
            with open('tr-cities.json', 'r', encoding='utf-8') as f:
                geojson_data = json.load(f)
        except FileNotFoundError:
            st.error("GeoJSON dosyası bulunamadı. 'tr-cities.json' ekleyin.")
            st.stop()

        m = folium.Map(location=[38, 35], zoom_start=5, tiles='cartodbpositron')
        folium.Choropleth(
            geo_data=geojson_data,
            data=df_city_sales,
            columns=['CITY', 'URUNADET'],
            key_on='feature.properties.name',
            fill_color='YlOrRd',
            fill_opacity=0.7,
            line_opacity=0.2,
            legend_name='Satış Adedi'
        ).add_to(m)

        for _, row in df_city_sales.iterrows():
            # Dummy coordinates: need real lat/lon mapping
            # Here using geocoded lookup could be added
            folium.Marker(
                location=[38, 35],
                popup=f"{row['CITY']}: {int(row['URUNADET'])} adet"
            ).add_to(m)

        # Folium haritasını Streamlit'e embed et
        map_html = m._repr_html_()
        components.html(map_html, height=600)

    with tab3:
        st.subheader("2023 Ocak Tahmini")
        product = st.selectbox("Tahmin için ürün seçin", ["Ürün1", "Ürün2"])
        channel = st.multiselect("Kanal seçin", df_cross['KANAL'].unique())
        if st.button("Tahmini Hesapla"):
            if product == "Ürün1":
                df_fc = df_sales.groupby('YEARMONTH')['URUNADET'].sum().reset_index()
                df_fc.columns = ['ds', 'y']
            else:
                df_fc = df_cross.groupby('AY')['ÇAPRAZ ÜRÜN ADET'].sum().reset_index()
                df_fc.columns = ['ds', 'y']

            df_fc['ds'] = pd.to_datetime(df_fc['ds'], format='%Y%m')
            m_prophet = Prophet(yearly_seasonality=True)
            m_prophet.fit(df_fc)
            future = m_prophet.make_future_dataframe(periods=1, freq='M')
            forecast = m_prophet.predict(future)

            fig_fc = plot_plotly(m_prophet, forecast)
            st.plotly_chart(fig_fc)


            # 2023 yılının Ocak ayını filtrele
            jan_forecast = forecast[(forecast['ds'].dt.year == 2023) & (forecast['ds'].dt.month == 1)]

            if not jan_forecast.empty:
                jan_pred = jan_forecast['yhat'].values[0]
                st.write(f"2023-01 için öngörülen değer: {jan_pred:.2f}")
            else:
                st.warning("Ocak 2023 tahmini bulunamadı.")

        st.markdown("---")
        st.subheader("Rapor İndir")
        if st.button("PowerPoint Oluştur ve İndir"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Satış Analizi Özet"
            prs.save("sales_report.pptx")
            with open("sales_report.pptx", "rb") as f:
                st.download_button("PowerPoint İndir", f, file_name="sales_analysis.pptx")
