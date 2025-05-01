import streamlit as st
import pandas as pd
from prophet import Prophet
from prophet.plot import plot_plotly
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Satış Analizi Dashboard", layout="wide")

# --- Sidebar ---
st.sidebar.title("📂 Veri Yükle")
excel_file = st.sidebar.file_uploader("Excel dosyasını yükleyin", type=["xlsx", "xls"])

if excel_file:
    # Veri yükleme
    df_sales = pd.read_excel(excel_file, sheet_name=0)
    df_cross = pd.read_excel(excel_file, sheet_name=1)
    df_demo = pd.read_excel(excel_file, sheet_name=2)

    # --- Ana Sayfa ---
    st.title("Satış Analizi Uygulaması")
    st.markdown("Bu uygulama 2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # --- Veri Önizleme ---
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
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET','URUNHACIM']].sum().reset_index()
        fig1 = px.line(df_time, x='YEARMONTH', y=['URUNADET','URUNHACIM'], title='Zaman Serisi Analizi')
        st.plotly_chart(fig1, use_container_width=True)

        st.subheader("Şube Performansı Haritası")
        # placeholder: şehir demografi ile birleşip harita gösterimi
        st.info("Harita gösterimi eklemek için demografi verisiyle birleşim yapılacak.")

    with tab3:
        st.subheader("2023 Ocak Tahmini")
        product = st.selectbox("Tahmin için ürün seçin", ["Ürün1", "Ürün2"])
        channel = st.multiselect("Kanal seçin", df_cross['KANAL'].unique())
        if st.button("Tahmini Hesapla"):
            # Forecast için veri hazırlığı
            if product == 'Ürün1':
                df_fc = df_sales.groupby('YEARMONTH')['URUNADET'].sum().reset_index()
                df_fc.columns = ['ds','y']
            else:
                df_fc = df_cross.groupby('AY')['ÇAPRAZ ÜRÜN ADET'].sum().reset_index()
                df_fc.columns = ['ds','y']
            df_fc['ds'] = pd.to_datetime(df_fc['ds'], format='%Y%m')
            m = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False)
            m.fit(df_fc)
            future = m.make_future_dataframe(periods=1, freq='M')
            forecast = m.predict(future)
            fig_fc = plot_plotly(m, forecast)
            st.plotly_chart(fig_fc)
            st.write(f"2023-01 için öngörülen değer: {forecast.loc[forecast['ds']==pd.to_datetime('2023-01-31'),'yhat'].values[0]:.2f}")

        st.markdown("---")
        st.subheader("Rapor İndir")
        if st.button("PowerPoint Oluştur ve İndir"):
            prs = Presentation()
            # Örnek slide
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title = slide.shapes.title
            title.text = "Satış Analizi Özet"
            prs.save("sales_report.pptx")
            with open("sales_report.pptx", "rb") as f:
                st.download_button("PowerPoint İndir", f, file_name="sales_analysis.pptx")
