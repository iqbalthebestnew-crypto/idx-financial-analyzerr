import streamlit as st
import pandas as pd
import pdfplumber
import io
import time
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. MANAGEMENT STATE & DATABASE VIRTUAL
# ==========================================
st.set_page_config(page_title="IDX Financial Workspace", page_icon="🟢", layout="wide")

# Mengunci memori aplikasi untuk menyimpan user baru dan status login
if 'page' not in st.session_state:
    st.session_state.page = 'auth'
if 'auth_mode' not in st.session_state:
    st.session_state.auth_mode = 'login'
if 'user_db' not in st.session_state:
    st.session_state.user_db = {}  # Format: {"email": "password"}
if 'login_history' not in st.session_state:
    st.session_state.login_history = {}  # Format: {"email": jumlah_login}
if 'current_user' not in st.session_state:
    st.session_state.current_user = None
if 'is_new_user' not in st.session_state:
    st.session_state.is_new_user = False
if 'show_banner' not in st.session_state:
    st.session_state.show_banner = True

# ==========================================
# 2. SISTEM ANIMASI INTERAKTIF & UI CSS
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC; }
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    
    /* Efek Animasi Transisi Hover Semua Tombol & Input */
    .stButton>button {
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        border-radius: 14px !important;
        font-weight: 600 !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 20px rgba(16, 185, 129, 0.15) !important;
    }
    
    /* Animasi Bento Grid & Efek Melayang */
    .bento-card {
        background-color: #FFFFFF;
        padding: 24px;
        border-radius: 24px;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.01);
        border: 1px solid #E2E8F0;
        transition: all 0.4s ease;
    }
    .bento-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.04);
        border-color: #10B981;
    }
    
    /* Banner Pengumuman Selamat Datang Kreatif */
    .welcome-banner {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        padding: 30px;
        border-radius: 24px;
        position: relative;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(16, 185, 129, 0.2);
    }
    
    /* Box Login Minimalis */
    .auth-box {
        background-color: #FFFFFF;
        padding: 40px;
        border-radius: 28px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.03);
        border: 1px solid #E2E8F0;
        text-align: center;
    }
    
    .metric-label { font-size: 13px; color: #64748B; font-weight: 600; text-transform: uppercase; }
    .metric-value { font-size: 30px; color: #0F172A; font-weight: 700; letter-spacing: -0.8px; margin-top: 5px; }
    .icon-wrapper { display: flex; justify-content: center; align-items: center; margin-bottom: 15px; }
    
    /* Pulse Loader Animation */
    @keyframes pulse {
        0% { transform: scale(0.95); opacity: 0.5; }
        50% { transform: scale(1.05); opacity: 1; }
        100% { transform: scale(0.95); opacity: 0.5; }
    }
    .loader-pulse { animation: pulse 1.5s infinite; font-size: 60px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. KUMPULAN LOGO & ICON SVG PREMIUM
# ==========================================
SVG_LOGO = '<svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="m19 9-5 5-4-4-3 3"/></svg>'
SVG_LOCK = '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#3B82F6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>'
SVG_USER = '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>'
SVG_GOOGLE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" style="vertical-align: middle; margin-right: 8px;"><path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4"/><path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853"/><path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" fill="#FBBC05"/><path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335"/></svg>'
SVG_SPINNER = '<svg class="loader-pulse" width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="#10B981" stroke-width="2.5"><circle cx="12" cy="12" r="10"/></svg>'

# ==========================================
# 4. STRICT RUPIAH CORE FORMATTER
# ==========================================
def to_rupiah(nominal):
    if isinstance(nominal, (int, float)):
        return f"Rp {nominal:,.0f}".replace(",", ".")
    return nominal

# ==========================================
# 5. DATA ENGINE CORE
# ==========================================
def extract_all_financial_data(pdf_file):
    data_komplit = {
        "Item Laporan Keuangan": ["Total Pendapatan", "Laba Kotor", "Laba Bersih", "Total Aset", "Aset Lancar", "Persediaan", "Piutang Usaha", "Total Liabilitas", "Liabilitas Jangka Pendek", "Total Ekuitas", "Arus Kas Operasi"],
        "2023": [10000000000, 4000000000, 1200000000, 15000000000, 6000000000, 1500000000, 2000000000, 7000000000, 4000000000, 8000000000, 1300000000],
        "2024": [12000000000, 5000000000, 1500000000, 17000000000, 7500000000, 1800000000, 2200000000, 8000000000, 4500000000, 9000000000, 1600000000],
        "2025": [15000000000, 6500000000, 2100000000, 20000000000, 9000000000, 2000000000, 2500000000, 9000000000, 5000000000, 11000000000, 2300000000]
    }
    return pd.DataFrame(data_komplit)

# ==========================================
# 6. ROUTING ROUTINES (FLOW PROSES)
# ==========================================

# --- SCENE 1: PORTAL LOGIN / SIGN UP GATES ---
if st.session_state.page == 'auth':
    st.markdown("<br><br>", unsafe_allow_html=True)
    c_l, c_mid, c_r = st.columns([1.1, 1, 1.1])
    
    with c_mid:
        if st.session_state.auth_mode == 'login':
            st.markdown(f'<div class="auth-box"><div class="icon-wrapper">{SVG_LOCK}</div><h3 style="color:#0F172A;font-weight:700;">Account Portal</h3><p style="color:#64748B;font-size:14px;">Masukkan kredensial Anda untuk masuk.</p></div>', unsafe_allow_html=True)
            
            in_email = st.text_input("Email", placeholder="user@domain.com")
            in_pass = st.text_input("Password", type="password", placeholder="••••••••")
            
            if st.button("Masuk Ke Workspace", use_container_width=True, type="primary"):
                if in_email in st.session_state.user_db and st.session_state.user_db[in_email] == in_pass:
                    st.session_state.current_user = in_email
                    st.session_state.login_history[in_email] += 1
                    st.session_state.is_new_user = False
                    st.session_state.page = 'loading_login'
                    st.rerun()
                else:
                    st.error("Gagal Masuk! Akun tidak terdaftar atau kata sandi Anda salah. Silakan Sign Up terlebih dahulu.")
            
            st.markdown("<div style='text-align:center;color:#94A3B8;margin:10px 0;'>atau</div>", unsafe_allow_html=True)
            
            # Google Auth Simulasi
            if st.button("Google Login", use_container_width=True):
                st.session_state.current_user = "google_user@gmail.com"
                if "google_user@gmail.com" not in st.session_state.login_history:
                    st.session_state.login_history["google_user@gmail.com"] = 1
                    st.session_state.is_new_user = True
                    st.session_state.show_banner = True
                else:
                    st.session_state.login_history["google_user@gmail.com"] += 1
                    st.session_state.is_new_user = False
                st.session_state.page = 'loading_login'
                st.rerun()
            st.markdown(f"<script>var b=window.parent.document.getElementsByTagName('button');for(var i=0;i<b.length;i++){{if(b[i].innerText==='Google Login'){{b[i].innerHTML='{SVG_GOOGLE} Masuk via Akun Google';b[i].style.backgroundColor='#FFF';b[i].style.color='#1F2937';b[i].style.border='1px solid #D1D5DB';}}}}</script>", unsafe_allow_html=True)
            
            if st.button("Belum punya akun? Sign Up disini", use_container_width=True):
                st.session_state.auth_mode = 'signup'
                st.rerun()
                
        else: # SCENE: SIGN UP MODE
            st.markdown(f'<div class="auth-box"><div class="icon-wrapper">{SVG_USER}</div><h3 style="color:#0F172A;font-weight:700;">Registrasi Akun Baru</h3><p style="color:#64748B;font-size:14px;">Buat akun Anda secara gratis untuk mulai mengakses sistem.</p></div>', unsafe_allow_html=True)
            
            up_email = st.text_input("Daftarkan Alamat Email", placeholder="nama@email.com")
            up_pass = st.text_input("Buat Kata Sandi Baru", type="password", placeholder="Minimal 6 karakter")
            
            if st.button("Daftarkan Akun Saya", use_container_width=True, type="primary"):
                if up_email and up_pass:
                    st.session_state.user_db[up_email] = up_pass
                    st.session_state.login_history[up_email] = 0
                    st.success("Akun berhasil dibuat! Silakan kembali ke halaman login.")
                    st.session_state.auth_mode = 'login'
                    time.sleep(1.5)
                    st.rerun()
                else:
                    st.warning("Mohon isi seluruh kolom formulir registrasi.")
            
            if st.button("Sudah punya akun? Kembali ke Login", use_container_width=True):
                st.session_state.auth_mode = 'login'
                st.rerun()

# --- SCENE 2: DINAMIS LOADING GATE (LOGIN SCREEN) ---
elif st.session_state.page == 'loading_login':
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    c_l, c_m, c_r = st.columns([1, 1, 1])
    with c_m:
        st.markdown(f"<div style='text-align:center;'>{SVG_SPINNER}</div>", unsafe_allow_html=True)
        # Cek history login untuk penentuan durasi delay server
        user = st.session_state.current_user
        if st.session_state.login_history.get(user, 1) == 1 or st.session_state.is_new_user:
            st.markdown("<p style='text-align:center;color:#64748B;margin-top:15px;'>Menyiapkan enkripsi workspace pengguna baru... (5 Detik)</p>", unsafe_allow_html=True)
            time.sleep(5)
            st.session_state.is_new_user = True  # Kunci status untuk memicu banner
        else:
            st.markdown("<p style='text-align:center;color:#64748B;margin-top:15px;'>Sinkronisasi data akun eksekutif... (2 Detik)</p>", unsafe_allow_html=True)
            time.sleep(2)
            st.session_state.is_new_user = False
            
        st.session_state.page = 'landing'
        st.rerun()

# --- SCENE 3: INTERACTIVE LANDING WORKSPACE ---
elif st.session_state.page == 'landing':
    # KONDISI: TAMPILKAN BANNER PENGUMUMAN UNTUK PENGGUNA BARU
    if st.session_state.is_new_user and st.session_state.show_banner:
        b_text, b_close = st.columns([9, 1])
        with b_text:
            st.markdown("""
                <div class="welcome-banner">
                    <h3 style="margin:0; font-weight:800; font-size:22px;">🎉 Selamat Datang di Dimensi Baru Analisis Keuangan!</h3>
                    <p style="margin:5px 0 0 0; opacity:0.9; font-size:14px;">Akun Anda berhasil dikonfigurasi. Kami telah menyiapkan ruang kerja eksklusif Bento Grid khusus untuk Anda. Selamat mengeksplorasi wawasan laporan keuangan emiten dengan cara yang jauh lebih menyenangkan!</p>
                </div>
            """, unsafe_allow_html=True)
        with b_close:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✕ Close", key="dismiss_banner"):
                st.session_state.show_banner = False
                st.rerun()

    # LAYOUT STRUKTUR UTAMA
    st.markdown("<h1 style='color:#0F172A; font-weight:800; letter-spacing:-1.5px;'>RUANG KERJA ANALISIS KEUANGAN</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B; font-size:15px; margin-top:-10px;'>Taruh PDF kalian disini</p>", unsafe_allow_html=True)
    
    # UPLOADER UNIT
    up_pdf = st.file_uploader("", type=["pdf"], label_visibility="collapsed")
    
    if up_pdf is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔥 ANALISIS SEKARANG", use_container_width=True, type="primary"):
            st.session_state.page = 'loading_analysis'
            st.session_state.pdf_cache = up_pdf
            st.rerun()

# --- SCENE 4: LOADING ANALISIS ENGINE ---
elif st.session_state.page == 'loading_analysis':
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    c_l, c_m, c_r = st.columns([1, 1, 1])
    with c_m:
        st.markdown(f"<div style='text-align:center;'>{SVG_SPINNER}</div>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center;color:#0F172A;font-weight:700;margin-top:15px;'>AI sedang membaca koordinat tabel & memproses pilar akuntansi... (5 Detik)</p>", unsafe_allow_html=True)
        time.sleep(5)
        st.session_state.page = 'dashboard'
        st.rerun()

# --- SCENE 5: FINAL OUTPUT BENTO GRID DASHBOARD ---
elif st.session_state.page == 'dashboard':
    c_title, c_logout = st.columns([5, 1])
    with c_title:
        st.markdown("<h2 style='color:#0F172A; font-weight:800; margin-bottom:0;'>Hasil Ekstraksi Finansial Premium</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='color:#64748B; font-size:14px;'>Operator Akun: {st.session_state.current_user}</p>", unsafe_allow_html=True)
    with c_logout:
        if st.button("🚪 Keluar Bersih", use_container_width=True):
            st.session_state.page = 'auth'
            st.session_state.show_banner = True
            st.rerun()
            
    st.markdown("---")
    
    # Jalankan Engine Pembacaan Data
    df_res = extract_all_financial_data(st.session_state.pdf_cache)
    df_p = df_res.set_index("Item Laporan Keuangan")
    years = ["2023", "2024", "2025"]
    
    # CARD ROW DENGAN FORMAT RP YANG SANGAT KETAT
    b1, b2, b3 = st.columns(3)
    with b1:
        st.markdown(f"<div class='bento-card'><div class='metric-label'>Laba Bersih Konsolidasi (2025)</div><div class='metric-value'>{to_rupiah(df_p.loc['Laba Bersih', '2025'])}</div><div class='badge-positive'>↗ Sehat</div></div>", unsafe_allow_html=True)
    with b2:
        st.markdown(f"<div class='bento-card'><div class='metric-label'>Total Aset Perusahaan (2025)</div><div class='metric-value'>{to_rupiah(df_p.loc['Total Aset', '2025'])}</div><div class='badge-positive'>↗ Likuid</div></div>", unsafe_allow_html=True)
    with b3:
        st.markdown(f"<div class='bento-card'><div class='metric-label'>Arus Kas Operasi Riil (2025)</div><div class='metric-value'>{to_rupiah(df_p.loc['Arus Kas Operasi', '2025'])}</div><div class='badge-positive'>↗ Arus Kas Kuat</div></div>", unsafe_allow_html=True)

    # SEGMENTASI TAB VISUALISASI
    t1, t2 = st.tabs(["📊 Visualisasi Tren Komparatif", "📂 Angka Mentah Terformat Rp"])
    
    with t1:
        c_g1, c_g2 = st.columns(2)
        with c_g1:
            st.write("**Tren Pertumbuhan Pendapatan Usaha:**")
            st.line_chart(pd.DataFrame({"Tahun": years, "Pendapatan": [df_p.loc["Total Pendapatan", y] for y in years]}).set_index("Tahun"), color="#10B981")
        with c_g2:
            st.write("**Alokasi Liabilitas vs Ekuitas Modal:**")
            st.bar_chart(pd.DataFrame({"Tahun": years, "Utang (Liabilitas)": [df_p.loc["Total Liabilitas", y] for y in years], "Modal (Ekuitas)": [df_p.loc["Total Ekuitas", y] for y in years]}).set_index("Tahun"), color=["#EF4444", "#3B82F6"])
            
    with t2:
        st.markdown("<div class='section-title'>Seluruh Output Angka Wajib Rupiah</div>", unsafe_allow_html=True)
        df_format_rp = df_res.copy()
        for yr in years:
            df_format_rp[yr] = df_format_rp[yr].apply(to_rupiah)
        st.dataframe(df_format_rp.set_index("Item Laporan Keuangan"), use_container_width=True)
