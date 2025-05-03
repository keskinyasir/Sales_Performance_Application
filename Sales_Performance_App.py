import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from prophet import Prophet
from prophet.plot import plot_plotly
from pptx import Presentation
from pptx.util import Inches
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
    # Veri Yükleme
    df_sales = pd.read_excel(excel_file, sheet_name="SATIŞ")
    df_cross = pd.read_excel(excel_file, sheet_name="ÇAPRAZ SATIŞ")
    df_demo = pd.read_excel(excel_file, sheet_name="ILCE DEMOGRAFI")

    # Tarih formatı
    df_sales['YEARMONTH'] = pd.to_datetime(df_sales['YEARMONTH'].astype(str), format='%Y%m')
    df_cross['AY'] = pd.to_datetime(df_cross['AY'].astype(str), format='%Y%m')

    # Başlık
    st.title("Satış Analizi Dashboard")
    st.markdown("Bu uygulama 2021-2022 dönemine ait satış, çapraz satış ve demografi verilerini analiz eder.")

    # Sekmeler
    tab1, tab2, tab3 = st.tabs(["Veri Önizleme","EDA & Görselleştirme","Tahmin & Rapor"])

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
        st.subheader("Aylık Toplam Satış Adedi")
        df_time = df_sales.groupby('YEARMONTH')[['URUNADET','URUNHACIM']].sum().reset_index()
        fig_time = px.line(df_time, x='YEARMONTH', y=['URUNADET'], title='Zaman Serisi Analizi - Adet')
        st.plotly_chart(fig_time, use_container_width=True)

        st.subheader("Aylık Toplam Satış Hacmi")
        fig_time2 = px.line(df_time, x='YEARMONTH', y=['URUNHACIM'], title='Zaman Serisi Analizi - Hacim')
        st.plotly_chart(fig_time2, use_container_width=True)

        st.subheader("Şube Performansı Haritası (İl Bazında)")
        st.info("Harita burada yer alacak.")

    # -------- TAHMİN & RAPOR --------
    with tab3:
        st.subheader("Tahminleme (Prophet + Tüm Regresörler)")
        # Forecast butonunda hesaplama ve saklama
        if st.button("Tahmini Hesapla"):
            # 1) Aylık satış
            monthly_sales = df_sales.groupby('YEARMONTH').agg({
                'URUNADET':'sum','URUNHACIM':'sum',
                'ABONE_YAS_0_3AY':'sum','ABONE_YAS_4_12AY':'sum',
                'ABONE_YAS_1_3YAS':'sum','ABONE_YAS_3_YAS':'sum'
            }).reset_index()
            monthly_sales.columns = ['periode','y','URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS']
            monthly_sales['ds'] = pd.to_datetime(monthly_sales['periode'], format='%Y%m')

            # 2) Çapraz satış
            cross_month = df_cross.groupby('AY').agg({
                'ÇAPRAZ ÜRÜN ADET':'sum','5GUNIPTAL':'sum','6-45GUNIPTAL':'sum'
            }).reset_index()
            cross_month.columns = ['periode_cross','cross_sales','cancel_5d','cancel_6_45d']
            cross_month['ds'] = pd.to_datetime(cross_month['periode_cross'], format='%Y%m')

            # 3) Demografi sabit
            demo_mean = df_demo.groupby('IL').mean(numeric_only=True).mean().to_dict()

            # 4) Birleştir ve regresör ekle
            df_fc = pd.merge(monthly_sales,
                             cross_month[['ds','cross_sales','cancel_5d','cancel_6_45d']],
                             on='ds', how='left')
            for k,v in demo_mean.items(): df_fc[k] = v

            regressors = ['URUNHACIM','A0_3','A4_12','A1_3','A3_PLUS',
                          'cross_sales','cancel_5d','cancel_6_45d'] + list(demo_mean.keys())

            # Prophet modeli
            m = Prophet()
            for reg in regressors: m.add_regressor(reg)
            m.fit(df_fc[['ds','y'] + regressors])

            # Gelecek tahmin
            future = m.make_future_dataframe(periods=1, freq='M')
            last = df_fc.iloc[-1]
            for reg in regressors: future[reg] = last[reg]
            forecast = m.predict(future)

            # Sonuçları göster ve kaydet
            st.plotly_chart(plot_plotly(m, forecast), use_container_width=True)
            next_pred = forecast.loc[forecast['ds']==forecast['ds'].max(),'yhat'].iloc[0]
            st.write(f"2023 Ocak öngörülen satış adedi: {next_pred:.0f}")

            # State'e kaydet
            st.session_state['forecast_df'] = forecast

        # ------------ Rapor İndir ------------
        st.markdown("---")
        st.subheader("Rapor İndir")
        if st.button("PowerPoint Oluştur ve İndir"):
            if 'forecast_df' not in st.session_state:
                st.error("Önce 'Tahmini Hesapla' butonuna tıklayın.")
            else:
                prs = Presentation()
                # Slide 1: Ürün1 Satış
                slide1 = prs.slides.add_slide(prs.slide_layouts[5])
                slide1.shapes.title.text = "Ürün1 Satış Verisi"
                tbl1 = df_sales.head(3)
                rows1, cols1 = tbl1.shape
                table1 = slide1.shapes.add_table(rows1+1, cols1, Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
                for i, col in enumerate(tbl1.columns): table1.cell(0,i).text = str(col)
                for r_idx, (_, row) in enumerate(tbl1.iterrows(), start=1):
                    for c_idx, val in enumerate(row): table1.cell(r_idx,c_idx).text = str(val)

                # Slide 2: Ürün2 Çapraz Satış
                slide2 = prs.slides.add_slide(prs.slide_layouts[5])
                slide2.shapes.title.text = "Ürün2 Çapraz Satış Verisi"
                tbl2 = df_cross.head(3)
                rows2, cols2 = tbl2.shape
                table2 = slide2.shapes.add_table(rows2+1, cols2, Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
                for i, col in enumerate(tbl2.columns): table2.cell(0,i).text = str(col)
                for r_idx, (_, row) in enumerate(tbl2.iterrows(), start=1):
                    for c_idx, val in enumerate(row): table2.cell(r_idx,c_idx).text = str(val)

                # Slide 3: Demografi
                slide3 = prs.slides.add_slide(prs.slide_layouts[5])
                slide3.shapes.title.text = "Demografi Verisi"
                tbl3 = df_demo.head(3)
                rows3, cols3 = tbl3.shape
                table3 = slide3.shapes.add_table(rows3+1, cols3, Inches(0.5), Inches(1.5), Inches(9), Inches(3)).table
                for i, col in enumerate(tbl3.columns): table3.cell(0,i).text = str(col)
                for r_idx, (_, row) in enumerate(tbl3.iterrows(), start=1):
                    for c_idx, val in enumerate(row): table3.cell(r_idx,c_idx).text = str(val)

                # Slide 4: Tahmin
                slide4 = prs.slides.add_slide(prs.slide_layouts[5])
                slide4.shapes.title.text = "2023 Ocak Öngörü"
                forecast = st.session_state['forecast_df']
                df_res = forecast[['ds','yhat']].tail(1)
                rows4, cols4 = df_res.shape
                table4 = slide4.shapes.add_table(rows4+1, cols4, Inches(0.5), Inches(1.5), Inches(6), Inches(2)).table
                for i, col in enumerate(df_res.columns): table4.cell(0,i).text = str(col)
                for r_idx, (_, row) in enumerate(df_res.iterrows(), start=1):
                    for c_idx, val in enumerate(row): table4.cell(r_idx,c_idx).text = str(val)

                # Kaydet ve indir
                prs.save("sales_report.pptx")
                with open("sales_report.pptx","rb") as f:
                    st.download_button("PowerPoint İndir", f, file_name="sales_analysis.pptx")
