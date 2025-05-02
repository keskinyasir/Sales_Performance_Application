import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from prophet import Prophet
from prophet.plot import plot_plotly
from pptx import Presentation
import json
import folium
from streamlit_folium import st_folium

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

    # Başlık
    st.title("Satış Analizi Dashboard")
    st.markdown("Bu uygulama 2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # --- Sekmeler ---
    tab1, tab2, tab3 = st.tabs(["Veri Önizleme", "EDA & Görselleştirme", "Tahmin & Rapor"])

    # ----- Sekme 1: Veri Önizleme -----
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

    # ----- Sekme 2: EDA & Görselleştirme -----
    with tab2:
        st.subheader("Aylık Toplam Satış Adedi ve Tutarı")
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET', 'URUNHACIM']].sum().reset_index()
        fig_time = px.line(
            df_time,
            x='YEARMONTH',
            y=['URUNADET', 'URUNHACIM'],
            title='Zaman Serisi Analizi',
            labels={'value': 'Değer', 'variable': 'Metrik'}
        )
        st.plotly_chart(fig_time, use_container_width=True)

        st.subheader("Şube Performansı Haritası (Folium ile)")
        # Şube ile şehir eşleştirme
        dealer_city = df_cross[['DEALER_CODE', 'CITY']].drop_duplicates()
        df_sales_map = df_sales.merge(dealer_city, on='DEALER_CODE', how='left')
        df_city_sales = df_sales_map.groupby('CITY')['URUNADET'].sum().reset_index()

        # GeoJSON yükleme
        try:
            with open('tr-cities.json', 'r', encoding='utf-8') as f:
                geojson_data = json.load(f)
        except FileNotFoundError:
            st.error("GeoJSON dosyası bulunamadı. 'tr-cities.json' dosyasını ekleyin.")
            st.stop()

        # Folium harita oluşturma
        m = folium.Map(location=[38, 35], zoom_start=5, tiles='cartodbpositron')
        choropleth = folium.Choropleth(
            geo_data=geojson_data,
            data=df_city_sales,
            columns=['CITY', 'URUNADET'],
            key_on='feature.properties.name',
            fill_color='YlOrRd',
            fill_opacity=0.7,
            line_opacity=0.2,
            legend_name='Satış Adedi'
        ).add_to(m)

        # Tooltip ekleme
        folium.GeoJson(
            geojson_data,
            name='Satış Bilgisi',
            style_function=lambda x: {'fillColor': 'transparent', 'color': 'transparent'},
            tooltip=folium.GeoJsonTooltip(
                fields=['name'],
                aliases=['İl:'],
                labels=True,
                sticky=False,
                toLocaleString=True,
                style=('background-color: white; color: #333333; font-family: arial; font-size: 12px; padding: 5px;')
            )
        ).add_to(m)

        # Satış adetlerini popup olarak ekleme
        for _, row in df_city_sales.iterrows():
            folium.Marker(
                location=[df_demo[df_demo['IL']==row['CITY']]['ORTALAMA_HANE_GELIRI_(TL_/_AY)'].mean() or 38,
                          df_demo[df_demo['IL']==row['CITY']]['TOPLAM_CALISAN_(KISI)'].mean() or 35],
                popup=f"{row['CITY']}: {int(row['URUNADET'])} adet",
                icon=folium.DivIcon(html=f"<div style='font-size: 12px; color: black;'>{int(row['URUNADET'])}</div>")
            ).add_to(m)

        st_folium(m, width=700, height=500)

    # ----- Sekme 3: Tahmin & Rapor -----
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

            jan_pred = forecast.loc[forecast['ds'] == pd.to_datetime('2023-01-31'), 'yhat'].values[0]
            st.write(f"2023-01 için öngörülen değer: {jan_pred:.2f}")

        st.markdown("---")
        st.subheader("Rapor İndir")
        if st.button("PowerPoint Oluştur ve İndir"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Satış Analizi Özet"
            prs.save("sales_report.pptx")
            with open("sales_report.pptx", "rb") as f:
                st.download_button("PowerPoint İndir", f, file_name="sales_analysis.pptx")
