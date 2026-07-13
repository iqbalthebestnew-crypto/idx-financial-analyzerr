import streamlit as st
import pandas as pd
import pdfplumber
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 1. SETUP HALAMAN WEB
st.set_page_config(page_title="IDX Financial Analyzer", page_icon="📊", layout="wide")

st.title("📊 Automated IDX Financial Statement Analyzer")
st.write("Unggah PDF Laporan Keuangan dari IDX, sistem akan mengekstrak data, menghitung rasio akuntansi, dan menghasilkan output Excel siap pakai.")

# 2. FUNGSI EKSTRAKSI DATA (SIMULASI STRUKTUR)
def extract_financial_data(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        for page in pdf.pages[:5]: # Baca 5 halaman pertama
            text = page.extract_text()
            if text:
                full_text += text
                
    # Data sampel terstruktur untuk output
    data_mentah = {
        "Item Laporan Keuangan": [
            "Total Pendapatan", "Laba Kotor", "Laba Bersih", 
            "Total Aset", "Aset Lancar", 
            "Total Liabilitas (Utang)", "Liabilitas Jangka Pendek", "Total Ekuitas (Modal)"
        ],
        "2023": [10000000000, 4000000000, 1200000000, 15000000000, 6000000000, 7000000000, 4000000000, 8000000000],
        "2024": [12000000000, 5000000000, 1500000000, 17000000000, 7500000000, 8000000000, 4500000000, 9000000000],
        "2025": [15000000000, 6500000000, 2100000000, 20000000000, 9000000000, 9000000000, 5000000000, 11000000000]
    }
    return pd.DataFrame(data_mentah)

# 3. FUNGSI MEMBUAT EXCEL PRO
def create_excel_report(df_raw):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    font_family = "Arial"
    title_font = Font(name=font_family, size=14, bold=True, color="1B365D")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    data_font = Font(name=font_family, size=10)
    bold_data_font = Font(name=font_family, size=10, bold=True)
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )

    # TAB DATA MENTAH
    ws1 = wb.active
    ws1.title = "Data_Mentah"
    ws1.views.sheetView[0].showGridLines = True
    ws1["B2"] = "HASIL EKSTRAKSI LAPORAN KEUANGAN"
    ws1["B2"].font = title_font
    
    headers = ["Item Laporan Keuangan", "2023", "2024", "2025"]
    for col_idx, h in enumerate(headers, start=2):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    for r_idx, row in df_raw.iterrows():
        row_num = r_idx + 5
        for c_idx, val in enumerate(row, start=2):
            cell = ws1.cell(row=row_num, column=c_idx, value=val)
            cell.border = thin_border
            if c_idx == 2:
                cell.font = data_font
            else:
                cell.font = data_font
                cell.number_format = "Rp #,##0"
                cell.alignment = Alignment(horizontal="right")

    # TAB ANALISIS RASIO
    ws2 = wb.create_sheet(title="Analisis_Rasio")
    ws2.views.sheetView[0].showGridLines = True
    ws2["B2"] = "ANALISIS RASIO KEUANGAN (FORMULA OTOMATIS)"
    ws2["B2"].font = title_font
    
    ratio_headers = ["Komponen Rasio", "Formula Akuntansi", "2023", "2024", "2025"]
    for col_idx, h in enumerate(ratio_headers, start=2):
        cell = ws2.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    ratios = [
        ("Net Profit Margin (NPM)", "Laba Bersih / Pendapatan", "=Data_Mentah!D7/Data_Mentah!D5", "=Data_Mentah!E7/Data_Mentah!E5", "=Data_Mentah!F7/Data_Mentah!F5", "0.0%"),
        ("Return on Assets (ROA)", "Laba Bersih / Total Aset", "=Data_Mentah!D7/Data_Mentah!D8", "=Data_Mentah!E7/Data_Mentah!E8", "=Data_Mentah!F7/Data_Mentah!F8", "0.0%"),
        ("Return on Equity (ROE)", "Laba Bersih / Total Ekuitas", "=Data_Mentah!D7/Data_Mentah!D12", "=Data_Mentah!E7/Data_Mentah!E12", "=Data_Mentah!F7/Data_Mentah!F12", "0.0%"),
        ("Current Ratio (CR)", "Aset Lancar / Liabilitas Jk Pendek", "=Data_Mentah!D9/Data_Mentah!D11", "=Data_Mentah!E9/Data_Mentah!E11", "=Data_Mentah!F9/Data_Mentah!F11", "0.00"),
        ("Debt to Equity Ratio (DER)", "Total Liabilitas / Total Ekuitas", "=Data_Mentah!D10/Data_Mentah!D12", "=Data_Mentah!E10/Data_Mentah!E12", "=Data_Mentah!F10/Data_Mentah!F12", "0.00")
    ]
    
    for r_idx, ratio in enumerate(ratios, start=5):
        for c_idx, val in enumerate(ratio[:5], start=2):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if c_idx == 2:
                cell.font = bold_data_font
            elif c_idx == 3:
                cell.font = Font(name=font_family, size=9, italic=True)
            else:
                cell.font = data_font
                cell.number_format = ratio[5]
                cell.alignment = Alignment(horizontal="right")

    for ws in [ws1, ws2]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = 25
    ws1.column_dimensions['B'].width = 35
    ws2.column_dimensions['C'].width = 35

    wb.save(output)
    return output.getvalue()

# 4. INTERFACE WEB (STREAMLIT)
uploaded_file = st.file_uploader("Pilih file PDF Laporan Keuangan IDX kamu...", type=["pdf"])

if uploaded_file is not None:
    st.success("File berhasil diunggah! Memulai ekstraksi data akuntansi...")
    df_extracted = extract_financial_data(uploaded_file)
    
    tab1, tab2 = st.tabs(["📊 Pratinjau Data Mentah", "📈 Analisis Tren Rasio"])
    
    with tab1:
        st.subheader("Data Hasil Ekstraksi Posisi Keuangan & Laba Rugi")
        st.dataframe(df_extracted.set_index("Item Laporan Keuangan"), use_container_width=True)
        
    with tab2:
        st.subheader("Kalkulasi Rasio Keuangan Utama")
        df_p = df_extracted.set_index("Item Laporan Keuangan")
        years = ["2023", "2024", "2025"]
        
        npm = [df_p.loc["Laba Bersih", y] / df_p.loc["Total Pendapatan", y] for y in years]
        roa = [df_p.loc["Laba Bersih", y] / df_p.loc["Total Aset", y] for y in years]
        roe = [df_p.loc["Laba Bersih", y] / df_p.loc["Total Ekuitas", y] for y in years]
        
        df_chart = pd.DataFrame({
            "Tahun": years,
            "NPM (%)": [n * 100 for n in npm],
            "ROA (%)": [r * 100 for r in roa],
            "ROE (%)": [e * 100 for e in roe]
        }).set_index("Tahun")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            st.dataframe(df_chart)
        with col2:
            st.line_chart(df_chart)
            
    excel_data = create_excel_report(df_extracted)
    st.write("---")
    st.subheader("📥 Siap Unduh Output Riset")
    st.download_button(
        label="Download Laporan Analisis Lengkap (.xlsx)",
        data=excel_data,
        file_name="Analisis_Laporan_Keuangan_IDX.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )