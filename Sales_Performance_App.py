import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
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

        st.subheader("Şube Performansı Haritası (Plotly Mapbox)")
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

        # Choropleth layer
        fig_map = px.choropleth_mapbox(
            df_city_sales,
            geojson=geojson_data,
            locations='CITY',
            featureidkey='properties.name',
            color='URUNADET',
            color_continuous_scale='Viridis',
            mapbox_style='carto-positron',
            zoom=5,
            center={'lat': 38, 'lon': 35},
            opacity=0.6,
            labels={'URUNADET': 'Satış Adedi'}
        )

        # İl merkez koordinatlarının hesaplanması
        centroids = {}
        for feat in geojson_data['features']:
            name = feat['properties']['name']
            geom = feat['geometry']
            coords = []
            if geom['type'] == 'Polygon':
                for ring in geom['coordinates']:
                    coords.extend(ring)
            elif geom['type'] == 'MultiPolygon':
                for part in geom['coordinates']:
                    for ring in part:
                        coords.extend(ring)
            lons = [c[0] for c in coords]
            lats = [c[1] for c in coords]
            if lats and lons:
                centroids[name] = (sum(lats)/len(lats), sum(lons)/len(lons))

        # Satış adetlerini text olarak ekleme
        lats, lons, texts = [], [], []
        for _, row in df_city_sales.iterrows():
            city = row['CITY']
            lat, lon = centroids.get(city, (38, 35))
            lats.append(lat)
            lons.append(lon)
            texts.append(str(int(row['URUNADET'])))
        fig_map.add_trace(go.Scattermapbox(
            lat=lats,
            lon=lons,
            mode='text',
            text=texts,
            textfont=dict(size=12, color='black'),
            showlegend=False
        ))
        fig_map.update_layout(margin={'r':0,'t':30,'l':0,'b':0})
        st.plotly_chart(fig_map, use_container_width=True, config={'scrollZoom': True})

    with tab3:
        st.subheader("2023 Ocak Tahmini")
        product = st.selectbox("Tahmin için ürün seçin", ["Ürün1", "Ürün2"])
        channel = st.multiselect("Kanal seçin", df_cross['KANAL'].unique())
        if st.button("Tahmini Hesapla"):
            # Forecast veri hazırlığı
            if product == "Ürün1":
                df_fc = df_sales.groupby('YEARMONTH')['URUNADET'].sum().reset_index()
                df_fc.columns = ['ds', 'y']
            else:
                df_fc = df_cross.groupby('AY')['ÇAPRAZ ÜRÜN ADET'].sum().reset_index()
                df_fc.columns = ['ds', 'y']
            df_fc['ds'] = pd.to_datetime(df_fc['ds'], format='%Y%m')
            m = Prophet(yearly_seasonality=True, weekly_seasonality=False)
            m.fit(df_fc)
            future = m.make_future_dataframe(periods=1, freq='M')
            forecast = m.predict(future)
            fig_fc = plot_plotly(m, forecast)
            st.plotly_chart(fig_fc)
            jan_pred = forecast.loc[forecast['ds']==pd.to_datetime('2023-01-31'), 'yhat'].values[0]
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
