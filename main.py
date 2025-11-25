import streamlit as st
import sqlite3
import hashlib
import binascii
import os
import pandas as pd
from datetime import datetime
import openpyxl
from io import BytesIO
import base64
from pathlib import Path

st.set_page_config(
    page_title="MooLance Accounting",
    page_icon="🐄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DB_PATH = "app.db"
EXCEL_DB_PATH = "jurnal_umum.xlsx"
BUKU_BESAR_DB_PATH = "buku_besar.xlsx" 
NERACA_SALDO_DB_PATH = "neraca_saldo.xlsx" 
AJP_DB_PATH = "jurnal_penyesuaian.xlsx"
PERSEDIAAN_DB_PATH = "kartu_persediaan.xlsx"
LABA_RUGI_DB_PATH = "laporan_laba_rugi.xlsx"
PERUBAHAN_MODAL_DB_PATH = "laporan_perubahan_modal.xlsx"
POSISI_KEUANGAN_DB_PATH = "laporan_posisi_keuangan.xlsx"

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'username' not in st.session_state:
    st.session_state['username'] = None
if 'page' not in st.session_state:
    st.session_state['page'] = 'login'
if 'jurnal_view' not in st.session_state:
    st.session_state['jurnal_view'] = 'input'
if 'ajp_view' not in st.session_state:
    st.session_state['ajp_view'] = 'input'
if 'persediaan_view' not in st.session_state:  
    st.session_state['persediaan_view'] = 'input'

def load_css():
    st.markdown("""
        <style>
        /* Import Font */
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

        /* Global Reset */
        * { font-family: 'Poppins', sans-serif; }

        /* Background Gradient Utama - Warna #C8DBBE */
        .stApp {
            background: #C8DBBE;
            background-attachment: fixed;
        }

        /* Sembunyikan elemen bawaan Streamlit */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .block-container { padding-top: 1rem; padding-bottom: 2rem; }

        /* --- CUSTOM HEADER --- */
        .custom-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: rgba(255, 255, 255, 0.15);
            backdrop-filter: blur(10px);
            padding: 15px 30px;
            border-radius: 15px;
            margin-bottom: 20px;
            border: 1px solid rgba(255,255,255,0.2);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .header-logo {
            font-size: 24px;
            font-weight: 700;
            color: #333;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .header-user {
            color: #333;
            font-size: 14px;
        }

        /* --- STYLING TABS (MENU HORIZONTAL) --- */
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            background-color: rgba(255,255,255,0.3);
            padding: 10px;
            border-radius: 15px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: transparent;
            border-radius: 10px;
            color: #666;
            font-weight: 500;
            padding: 0 20px; 
        }
        .stTabs [data-baseweb="tab"]:hover {
            background-color: rgba(255,255,255,0.5);
            color: #333;
        }
        .stTabs [aria-selected="true"] {
            background-color: #C8DBBE !important;
            color: #333 !important;
            font-weight: 700 !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        }

        /* --- CONTENT CARDS --- */
        .content-card {
            background: white;
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.15);
            margin-top: 10px;
        }
        
        /* Header Title inside Content */
        h2, h3 { color: #2d3748; }
        p { color: #4a5568; }

        /* --- FORM INPUTS --- */
        .stTextInput input, .stSelectbox div[data-baseweb="select"] {
            color: #333 !important;
            background-color: #f7fafc;
            border-radius: 8px;
        }
        
        /* Login Box Specifics */
        .login-container {
            max-width: 400px;
            margin: 50px auto;
            background: white;
            padding: 40px;
            border-radius: 20px;
            box-shadow: 0 20px 50px rgba(0,0,0,0.2);
            text-align: center;
        }
        
        /* Table Styling */
        div[data-testid="stDataFrame"] {
            background: white;
            padding: 10px;
            border-radius: 10px;
            border: 1px solid #e2e8f0;
        }
        
        /* Button Styling */
        .stButton button {
            border-radius: 8px;
        }
        .delete-btn {
            background-color: #FF6B6B !important;
            color: white !important;
            border: none !important;
        }
        
        /* Buku Besar Table Styling - Warna #C8DBBE */
        .buku-besar-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
        }
        .buku-besar-table th {
            background-color: #C8DBBE;
            color: #333;
            padding: 12px;
            text-align: center;
            border: 1px solid #ddd;
            font-weight: 600;
        }
        .buku-besar-table td {
            padding: 10px;
            border: 1px solid #ddd;
            text-align: center;
        }
        .buku-besar-table .tanggal { width: 15%; }
        .buku-besar-table .keterangan { width: 35%; text-align: left; }
        .buku-besar-table .debit { width: 15%; text-align: right; }
        .buku-besar-table .kredit { width: 15%; text-align: right; }
        .buku-besar-table .saldo-debit { width: 10%; text-align: right; }
        .buku-besar-table .saldo-kredit { width: 10%; text-align: right; }
        .debit-amount { color: #27AE60; font-weight: 600; }
        .kredit-amount { color: #E74C3C; font-weight: 600; }
        .saldo-positive { color: #27AE60; font-weight: 600; }
        .saldo-negative { color: #E74C3C; font-weight: 600; }
        
        /* Jurnal Table Styling - Warna #C8DBBE */
        .jurnal-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
            border: 1px solid #ddd;
        }
        .jurnal-table th {
            background-color: #C8DBBE;
            color: #333;
            padding: 12px;
            text-align: center;
            border: 1px solid #ddd;
            font-weight: 600;
        }
        .jurnal-table td {
            padding: 12px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
        .jurnal-table .tanggal {
            width: 15%;
            text-align: center;
            font-weight: 500;
        }
        .jurnal-table .akun {
            width: 45%;
        }
        .jurnal-table .debit {
            width: 20%;
            text-align: right;
        }
        .jurnal-table .kredit {
            width: 20%;
            text-align: right;
        }
        .jurnal-table .debit-amount {
            color: #27AE60;
            font-weight: 600;
        }
        .jurnal-table .kredit-amount {
            color: #E74C3C;
            font-weight: 600;
        }
        .akun-debit {
            font-weight: 600;
            margin-bottom: 5px;
        }
        .akun-kredit {
            margin-left: 20px;
            color: #555;
            margin-bottom: 5px;
        }
        .keterangan {
            font-style: italic;
            color: #666;
            font-size: 12px;
            margin-top: 8px;
            padding-top: 8px;
            border-top: 1px dashed #eee;
        }

        /* Warna primary untuk button dan highlight - Warna #C8DBBE */
        .stButton > button:first-child {
            background-color: #C8DBBE;
            color: #333;
            border: 1px solid #B0C9A5;
        }
        
        .stButton > button:first-child:hover {
            background-color: #B0C9A5;
            color: #333;
        }
        
        /* Login title color */
        .login-container h2 {
            color: #333 !important;
        }

        /* Metric cards styling */
        [data-testid="stMetricValue"] {
            color: #333;
        }

        /* Success message styling */
        .stAlert > div:first-child {
            background-color: #F0FFF0;
            border-color: #27AE60;
        }

        /* Error message styling */
        .stAlert > div:first-child[data-testid="stAlertError"] {
            background-color: #FFF0F0;
            border-color: #E74C3C;
        }

        /* Text color adjustments for better contrast */
        .stText, .stMarkdown, .stWrite {
            color: #333;
        }
        </style>
    """, unsafe_allow_html=True)
    
def render_header():
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        # Logo besar di tengah
        try:
            logo_base64 = img_to_base64("logo.png")
            logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="80" style="border-radius: 15px; display: block; margin: 0 auto;">'
        except:
            logo_html = '<div style="font-size: 4rem; text-align: center;">🐄</div>'
        
        st.markdown(f"""
            <div style="text-align: center; margin-bottom: 10px;">
                {logo_html}
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
            <div style="text-align: center;">
                <div style="font-size: 28px; font-weight: 700; color: #333;">MooLance</div>
                <div style="font-size: 14px; color: #666; margin-top: -5px;">Accounting System</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
            <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 10px; width: 100%;">
        """, unsafe_allow_html=True)
        
        st.markdown(f"<div style='color: #333; text-align: right;'>Hi, <b>{st.session_state['username']}</b></div>", unsafe_allow_html=True)
        st.markdown("<div style='display: flex; justify-content: flex-end; width: 100%;'>", unsafe_allow_html=True)
        
        if st.button("🚪 Logout", key="top_logout"):
            st.session_state['logged_in'] = False
            st.session_state['username'] = None
            st.rerun()
        
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
def main_dashboard():
    render_header()
    
    tab_jurnal, tab_buku_besar, tab_neraca, tab_ajp, tab_laporan, tab_stok = st.tabs([
        "📒 Jurnal Umum", 
        "📖 Buku Besar", 
        "⚖️ Neraca Saldo", 
        "✏️ Jurnal Penyesuaian", 
        "📊 Laporan Keuangan", 
        "📦 Kartu Persediaan"
    ])

    with tab_jurnal:
        jurnal_umum_tab()

    with tab_buku_besar:
        display_buku_besar()

    with tab_neraca:
        neraca_saldo_tab()

    with tab_ajp:
        jurnal_penyesuaian_tab()

    with tab_laporan:
        laporan_keuangan_tab()  

    with tab_stok:
        kartu_persediaan_tab()

def jurnal_umum_tab():
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Input Jurnal Baru", use_container_width=True, 
                    type="primary" if st.session_state.get('jurnal_view', 'input') == 'input' else "secondary"):
            st.session_state['jurnal_view'] = 'input'
            st.rerun()
    
    with col2:
        if st.button("📋 Lihat Jurnal Umum", use_container_width=True,
                    type="primary" if st.session_state.get('jurnal_view', 'input') == 'view' else "secondary"):
            st.session_state['jurnal_view'] = 'view'
            st.rerun()
    
    st.markdown("---")
    
    if st.session_state.get('jurnal_view', 'input') == 'input':
        show_jurnal_input()
    else:
        show_jurnal_view()

def show_jurnal_input(): 
    with st.form("jurnal_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            tanggal = st.date_input("Tanggal Transaksi", value=datetime.now())
            akun_debit = st.selectbox("Akun Debit", get_akun_list(), key="debit_select")
            nominal_debit = st.number_input("Nominal Debit (Rp)", min_value=0, step=1000, key="debit_nominal")
        
        with col2:
            keterangan = st.text_input("Keterangan Transaksi")
            akun_kredit = st.selectbox("Akun Kredit", get_akun_list(), key="kredit_select")
            nominal_kredit = st.number_input("Nominal Kredit (Rp)", min_value=0, step=1000, key="kredit_nominal")
        
        is_balanced = nominal_debit == nominal_kredit
        
        submitted = st.form_submit_button("Simpan Jurnal", type="primary", use_container_width=True)
        
        if submitted:
            if not is_balanced:
                st.error("❌ Jurnal tidak balance! Debit dan Kredit harus sama.")
            elif not keterangan.strip():
                st.error("❌ Keterangan transaksi harus diisi.")
            else:
                success = save_jurnal_data(
                    tanggal, akun_debit, akun_kredit, 
                    nominal_debit, nominal_kredit, keterangan, 
                    st.session_state['username']
                )
                if success:
                    st.success("✅ Jurnal berhasil disimpan!")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 📊 Data Jurnal")
    
    df = get_jurnal_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data jurnal. Silakan input transaksi terlebih dahulu.")
    else:
        # Tampilkan tabel data dengan tombol hapus
        for idx, jurnal in df.sort_values('tanggal').iterrows():
            col1, col2, col3, col4, col5 = st.columns([2, 3, 2, 2, 1])
            
            with col1:
                st.markdown(f"<div style='color: black;'>{jurnal['tanggal']}</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"<div style='color: black; font-weight: bold;'>{jurnal['akun_debit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; margin-left: 20px;'>{jurnal['akun_kredit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; font-style: italic; font-size: 0.8em;'>{jurnal['keterangan']}</div>", unsafe_allow_html=True)
            
            with col3:
                if jurnal['nominal_debit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_debit']:,.0f}</div>", unsafe_allow_html=True)
            
            with col4:
                if jurnal['nominal_kredit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_kredit']:,.0f}</div>", unsafe_allow_html=True)
            
            with col5:
                is_persediaan_jurnal = any(jurnal['keterangan'].startswith(prefix) for prefix in ['[PERSEDIAAN]', '[HPP]'])
                
                if is_persediaan_jurnal:
                    st.button("🚫", key=f"disabled_{jurnal['id']}_{idx}", 
                             help="Jurnal dari persediaan - hapus dari menu Kartu Persediaan", disabled=True)
                else:
                    if st.button("🗑️", key=f"delete_{jurnal['id']}_{idx}", help="Hapus jurnal ini"):
                        if delete_jurnal_data(jurnal['id']):
                            st.rerun()
            
            st.markdown("---")

def show_jurnal_view():
    df = get_jurnal_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data jurnal. Silakan input transaksi terlebih dahulu.")
        return
    
    with st.container():
        col1, col2, col3, col4 = st.columns([2, 4, 2, 2])
        with col1:
            st.markdown("<div style='color: black; font-weight: bold;'>Tanggal</div>", unsafe_allow_html=True)
        with col2:
            st.markdown("<div style='color: black; font-weight: bold;'>Akun</div>", unsafe_allow_html=True)
        with col3:
            st.markdown("<div style='color: black; font-weight: bold;'>Debit</div>", unsafe_allow_html=True)
        with col4:
            st.markdown("<div style='color: black; font-weight: bold;'>Kredit</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        for _, jurnal in df.sort_values('tanggal').iterrows():
            col1, col2, col3, col4 = st.columns([2, 4, 2, 2])
            
            with col1:
                st.markdown(f"<div style='color: black;'>{jurnal['tanggal']}</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"<div style='color: black;'>{jurnal['akun_debit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; margin-left: 30px;'>{jurnal['akun_kredit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='font-style: italic; color: black; font-size: 0.8em; margin-top: 5px;'>{jurnal['keterangan']}</div>", unsafe_allow_html=True)
            
            with col3:
                if jurnal['nominal_debit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_debit']:,.0f}</div>", unsafe_allow_html=True)
                    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
            
            with col4:
                if jurnal['nominal_kredit'] > 0:
                    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_kredit']:,.0f}</div>", unsafe_allow_html=True)
            
            st.markdown("---")

def display_buku_besar():
    df_buku_besar = get_buku_besar_data()
    
    if len(df_buku_besar) == 0:
        st.info("📝 Belum ada data transaksi untuk ditampilkan.")
        return

    for akun in df_buku_besar['akun'].unique():
        akun_data = df_buku_besar[df_buku_besar['akun'] == akun]
        
        st.markdown(f"### **{akun}**")
        
        display_single_account_table(akun_data, akun)
        st.markdown("---")

def display_single_account_table(df, akun_nama):
    display_df = df[['tanggal', 'keterangan', 'debit', 'kredit', 'saldo_debit', 'saldo_kredit']].copy()
    
    for col in ['debit', 'kredit', 'saldo_debit', 'saldo_kredit']:
        display_df[col] = display_df[col].apply(lambda x: f"Rp {x:,.0f}" if x > 0 else "")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            'tanggal': st.column_config.TextColumn('Tanggal'),
            'keterangan': st.column_config.TextColumn('Keterangan'),
            'debit': st.column_config.TextColumn('Debit'),
            'kredit': st.column_config.TextColumn('Kredit'),
            'saldo_debit': st.column_config.TextColumn('Saldo Debit'),
            'saldo_kredit': st.column_config.TextColumn('Saldo Kredit')
        }
    )

def neraca_saldo_tab():
    generate_neraca_saldo_from_buku_besar()
    
    df_neraca = get_neraca_saldo_data()
    
    if len(df_neraca) == 0:
        st.info("📝 Belum ada data neraca saldo. Pastikan sudah ada transaksi di buku besar.")
        
    # Calculate totals
    total_debit = df_neraca['debit'].sum()
    total_kredit = df_neraca['kredit'].sum()
    
    # Display totals
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Debit", f"Rp {total_debit:,.0f}")
    with col2:
        st.metric("Total Kredit", f"Rp {total_kredit:,.0f}")
    
    # Check if balanced
    is_balanced = total_debit == total_kredit
    if is_balanced:
        st.success("✅ Neraca Saldo Balance!")
    else:
        st.error(f"❌ Neraca Saldo Tidak Balance! Selisih: Rp {abs(total_debit - total_kredit):,.0f}")
    
    st.markdown("---")
    
    display_neraca_saldo_table(df_neraca)

def display_neraca_saldo_table(df):
    display_df = df[['akun', 'debit', 'kredit']].copy()
    
    display_df['debit'] = display_df['debit'].apply(lambda x: f"Rp {x:,.0f}" if x > 0 else "")
    display_df['kredit'] = display_df['kredit'].apply(lambda x: f"Rp {x:,.0f}" if x > 0 else "")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            'akun': st.column_config.TextColumn('Akun', width="large"),
            'debit': st.column_config.TextColumn('Debit', width="medium"),
            'kredit': st.column_config.TextColumn('Kredit', width="medium")
        }
    )

def jurnal_penyesuaian_tab():
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Input Jurnal Penyesuaian", use_container_width=True, 
                    type="primary" if st.session_state.get('ajp_view', 'input') == 'input' else "secondary"):
            st.session_state['ajp_view'] = 'input'
            st.rerun()
    
    with col2:
        if st.button("📋 Lihat Jurnal Penyesuaian", use_container_width=True,
                    type="primary" if st.session_state.get('ajp_view', 'input') == 'view' else "secondary"):
            st.session_state['ajp_view'] = 'view'
            st.rerun()
    
    st.markdown("---")
    
    if st.session_state.get('ajp_view', 'input') == 'input':
        show_ajp_input()
    else:
        show_ajp_view()

def show_ajp_input(): 
    with st.form("ajp_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            tanggal = st.date_input("Tanggal Penyesuaian", value=datetime.now())
            akun_debit = st.selectbox("Akun Debit", get_akun_list(), key="ajp_debit_select")
            nominal_debit = st.number_input("Nominal Debit (Rp)", min_value=0, step=1000, key="ajp_debit_nominal")
        
        with col2:
            keterangan = st.text_input("Keterangan Penyesuaian", placeholder="Contoh: Penyusutan peralatan, Beban yang masih harus dibayar, dll.")
            akun_kredit = st.selectbox("Akun Kredit", get_akun_list(), key="ajp_kredit_select")
            nominal_kredit = st.number_input("Nominal Kredit (Rp)", min_value=0, step=1000, key="ajp_kredit_nominal")
        
        is_balanced = nominal_debit == nominal_kredit
        
        submitted = st.form_submit_button("Simpan Jurnal Penyesuaian", type="primary", use_container_width=True)
        
        if submitted:
            if not is_balanced:
                st.error("❌ Jurnal tidak balance! Debit dan Kredit harus sama.")
            elif not keterangan.strip():
                st.error("❌ Keterangan penyesuaian harus diisi.")
            elif akun_debit == akun_kredit:
                st.error("❌ Akun debit dan kredit tidak boleh sama.")
            else:
                success = save_ajp_data(
                    tanggal, akun_debit, akun_kredit, 
                    nominal_debit, nominal_kredit, keterangan, 
                    st.session_state['username']
                )
                if success:
                    st.success("✅ Jurnal Penyesuaian berhasil disimpan! Neraca saldo telah di-update.")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 📊 Data Jurnal Penyesuaian")
    
    df = get_ajp_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data jurnal penyesuaian.")
    else:
        for idx, jurnal in df.sort_values('tanggal').iterrows():
            col1, col2, col3, col4, col5 = st.columns([2, 3, 2, 2, 1])
            
            with col1:
                st.markdown(f"<div style='color: black;'>{jurnal['tanggal']}</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"<div style='color: black; font-weight: bold;'>{jurnal['akun_debit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; margin-left: 20px;'>{jurnal['akun_kredit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; font-style: italic; font-size: 0.8em;'>{jurnal['keterangan']}</div>", unsafe_allow_html=True)
            
            with col3:
                if jurnal['nominal_debit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_debit']:,.0f}</div>", unsafe_allow_html=True)
            
            with col4:
                if jurnal['nominal_kredit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_kredit']:,.0f}</div>", unsafe_allow_html=True)
            
            with col5:
                # Tombol hapus untuk setiap entri
                if st.button("🗑️", key=f"delete_ajp_{jurnal['id']}_{idx}", help="Hapus jurnal penyesuaian ini"):
                    if delete_ajp_data(jurnal['id']):
                        st.rerun()
            
            st.markdown("---")

def show_ajp_view():
    df = get_ajp_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data jurnal penyesuaian.")
        return
    
    with st.container():
        col1, col2, col3, col4 = st.columns([2, 4, 2, 2])
        with col1:
            st.markdown("<div style='color: black; font-weight: bold;'>Tanggal</div>", unsafe_allow_html=True)
        with col2:
            st.markdown("<div style='color: black; font-weight: bold;'>Akun</div>", unsafe_allow_html=True)
        with col3:
            st.markdown("<div style='color: black; font-weight: bold;'>Debit</div>", unsafe_allow_html=True)
        with col4:
            st.markdown("<div style='color: black; font-weight: bold;'>Kredit</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        for _, jurnal in df.sort_values('tanggal').iterrows():
            col1, col2, col3, col4 = st.columns([2, 4, 2, 2])
            
            with col1:
                st.markdown(f"<div style='color: black;'>{jurnal['tanggal']}</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"<div style='color: black;'>{jurnal['akun_debit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; margin-left: 30px;'>{jurnal['akun_kredit']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='font-style: italic; color: black; font-size: 0.8em; margin-top: 5px;'>{jurnal['keterangan']}</div>", unsafe_allow_html=True)
            
            with col3:
                if jurnal['nominal_debit'] > 0:
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_debit']:,.0f}</div>", unsafe_allow_html=True)
                    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
            
            with col4:
                if jurnal['nominal_kredit'] > 0:
                    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
                    st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {jurnal['nominal_kredit']:,.0f}</div>", unsafe_allow_html=True)
            
            st.markdown("---")

def kartu_persediaan_tab():
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Input Transaksi Persediaan", use_container_width=True, 
                    type="primary" if st.session_state.get('persediaan_view', 'input') == 'input' else "secondary"):
            st.session_state['persediaan_view'] = 'input'
            st.rerun()
    
    with col2:
        if st.button("📋 Lihat Kartu Persediaan", use_container_width=True,
                    type="primary" if st.session_state.get('persediaan_view', 'input') == 'view' else "secondary"):
            st.session_state['persediaan_view'] = 'view'
            st.rerun()
    
    st.markdown("---")
    
    if st.session_state.get('persediaan_view', 'input') == 'input':
        show_persediaan_input()
    else:
        show_kartu_persediaan()

def show_persediaan_input(): 
    with st.form("persediaan_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            tanggal = st.date_input("Tanggal Transaksi*", value=datetime.now())
            jenis = st.selectbox("Jenis Transaksi*", ["pembelian", "penjualan"], 
                               format_func=lambda x: "Pembelian" if x == "pembelian" else "Penjualan")
            quantity = st.number_input("Quantity*", min_value=1, step=1)
            harga = st.number_input("Harga per Unit (Rp)*", min_value=0, step=1000)
        
        with col2:
            keterangan = st.text_input("Keterangan*", placeholder="Contoh: Pembelian barang A, Penjualan barang B, dll.")
            
            if jenis == "pembelian":
                akun_tambahan = st.selectbox("Akun Pembayaran*", ["Kas", "Utang Usaha"],
                                           help="Pilih akun untuk pembayaran pembelian")
            else: 
                akun_tambahan = st.selectbox("Akun Penerimaan*", ["Kas", "Piutang Usaha"],
                                            help="Pilih akun untuk penerimaan penjualan")
            
            total = quantity * harga
            st.metric("Total Transaksi", f"Rp {total:,.0f}")
        
        submitted = st.form_submit_button("Simpan Transaksi Persediaan", type="primary", use_container_width=True)
        
        if submitted:
            if not keterangan.strip():
                st.error("❌ Keterangan transaksi harus diisi.")
            else:
                success = save_persediaan_data(
                    tanggal, keterangan, jenis, akun_tambahan, 
                    quantity, harga, total, st.session_state['username']
                )
                if success:
                    st.success("✅ Transaksi Persediaan berhasil disimpan! Jurnal umum dan neraca saldo telah di-update.")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 📊 Data Transaksi Persediaan")
    
    df = get_persediaan_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data transaksi persediaan.")
    else:
        for idx, transaksi in df.sort_values('tanggal').iterrows():
            col1, col2, col3, col4, col5 = st.columns([2, 3, 2, 2, 1])
            
            with col1:
                st.markdown(f"<div style='color: black;'>{transaksi['tanggal']}</div>", unsafe_allow_html=True)
            
            with col2:
                jenis_text = "🛒 Pembelian" if transaksi['jenis'] == 'pembelian' else "💰 Penjualan"
                st.markdown(f"<div style='color: black; font-weight: bold;'>{jenis_text}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black;'>{transaksi['keterangan']}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; font-size: 0.8em;'>{transaksi['quantity']} unit @ Rp {transaksi['harga']:,.0f}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color: black; font-size: 0.8em;'>Akun: {transaksi['akun_tambahan']}</div>", unsafe_allow_html=True)
            
            with col3:
                if transaksi['jenis'] == 'pembelian':
                    st.markdown(f"<div style='color: #27AE60; font-weight: bold;'>+{transaksi['quantity']} unit</div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<div style='color: #E74C3C; font-weight: bold;'>-{transaksi['quantity']} unit</div>", unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"<div style='color: black; font-weight: bold;'>Rp {transaksi['total']:,.0f}</div>", unsafe_allow_html=True)
            
            with col5:
                if st.button("🗑️", key=f"delete_persediaan_{transaksi['id']}_{idx}", help="Hapus transaksi ini"):
                    if delete_persediaan_data(transaksi['id']):
                        st.rerun()
            
            st.markdown("---")

def show_kartu_persediaan():
    df = get_kartu_persediaan_data()
    
    if len(df) == 0:
        st.info("📝 Belum ada data kartu persediaan.")
        return
    
    saldo_akhir = df.iloc[-1]
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Saldo Quantity", f"{saldo_akhir['balance_qty']:,.0f} unit")
    with col2:
        st.metric("Harga Rata-rata", f"Rp {saldo_akhir['balance_price']:,.0f}")
    with col3:
        st.metric("Total Nilai Persediaan", f"Rp {saldo_akhir['balance_total']:,.0f}")
    
    st.markdown("---")
    
    display_df = df.copy()
    
    currency_cols = ['in_price', 'in_total', 'out_price', 'out_total', 'balance_price', 'balance_total']
    for col in currency_cols:
        display_df[col] = display_df[col].apply(lambda x: f"Rp {x:,.0f}" if x > 0 else "")
    
    qty_cols = ['in_qty', 'out_qty', 'balance_qty']
    for col in qty_cols:
        display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}" if x > 0 else "")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            'tanggal': st.column_config.TextColumn('Tanggal', width="small"),
            'keterangan': st.column_config.TextColumn('Keterangan', width="medium"),
            'in_qty': st.column_config.TextColumn('IN Qty', width="small"),
            'in_price': st.column_config.TextColumn('IN Price', width="small"),
            'in_total': st.column_config.TextColumn('IN Total', width="small"),
            'out_qty': st.column_config.TextColumn('OUT Qty', width="small"),
            'out_price': st.column_config.TextColumn('OUT Price', width="small"),
            'out_total': st.column_config.TextColumn('OUT Total', width="small"),
            'balance_qty': st.column_config.TextColumn('Balance Qty', width="small"),
            'balance_price': st.column_config.TextColumn('Balance Price', width="small"),
            'balance_total': st.column_config.TextColumn('Balance Total', width="small")
        }
    )

def laporan_keuangan_tab():
    tab1, tab2, tab3 = st.tabs([
        "💰 Laporan Laba Rugi", 
        "📈 Laporan Perubahan Modal", 
        "🏢 Laporan Posisi Keuangan"
    ])
    
    with tab1:
        show_laporan_laba_rugi()
    
    with tab2:
        show_laporan_perubahan_modal()
    
    with tab3:
        show_laporan_posisi_keuangan()

def show_laporan_laba_rugi(): 
    try:
        df_laba_rugi, total_pendapatan, total_beban, laba_rugi_bersih, laba_kotor, total_hpp = get_laporan_laba_rugi()
        
        save_laba_rugi_to_excel(total_pendapatan, total_hpp, total_beban, laba_kotor, laba_rugi_bersih)
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col3:
            if st.button("📥 Download Excel", use_container_width=True, type="primary"):
                download_laporan_laba_rugi(df_laba_rugi)
        
        # Key Metrics
        st.markdown("#### 📈 Ringkasan Performa")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="💰 Pendapatan Usaha",
                value=f"Rp {total_pendapatan:,.0f}",
                delta=None
            )
        
        with col2:
            st.metric(
                label="📦 Laba Kotor", 
                value=f"Rp {laba_kotor:,.0f}",
                delta=None
            )
        
        with col3:
            if laba_rugi_bersih >= 0:
                st.metric(
                    label="📈 Laba Bersih",
                    value=f"Rp {laba_rugi_bersih:,.0f}",
                    delta="Laba"
                )
            else:
                st.metric(
                    label="📉 Rugi Bersih",
                    value=f"Rp {abs(laba_rugi_bersih):,.0f}",
                    delta="Rugi",
                    delta_color="inverse"
                )
        
        with col4:
            margin_ratio = (laba_rugi_bersih / total_pendapatan * 100) if total_pendapatan > 0 else 0
            st.metric(
                label="🎯 Margin Laba",
                value=f"{margin_ratio:+.1f}%",
                delta=None
            )
        
        st.markdown("### Laporan Laba Rugi")
        
        st.markdown("###### PENDAPATAN USAHA")
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.write("Pendapatan Penjualan")
        with col2:
            st.write(f"**Rp {total_pendapatan:,.0f}**")
        with col3:
            st.write("")
        
        st.markdown(f"**Total Pendapatan Usaha: Rp {total_pendapatan:,.0f}**")
        st.divider()
        
        st.markdown("###### HARGA POKOK PENJUALAN")
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.write("Harga Pokok Penjualan")
        with col2:
            st.write("")
        with col3:
            st.write(f"**Rp {total_hpp:,.0f}**")
        
        st.markdown(f"**Total HPP: Rp {total_hpp:,.0f}**")
        st.divider()
        
        st.markdown("###### LABA KOTOR")
        st.success(f"**Rp {laba_kotor:,.0f}**")
        st.divider()
        
        st.markdown("###### BEBAN OPERASIONAL")
        
        df_neraca = get_neraca_saldo_data()
        akun_beban = [
            "Beban Gaji", "Beban Sewa", "Beban Listrik", "Beban Air", 
            "Beban Telepon", "Beban Transportasi", "Beban Penyusutan Kendaraan",
            "Beban Pakan Konsentrat", "Beban Obat dan Vitamin", "Beban Penyusutan Kandang",
            "Beban Pakan Katul", "Beban Pakan Ampas Tahu", "Beban Pakan Jerami",
            "Beban Perawatan Kandang"  
        ]
        
        ada_beban = False
        for akun in akun_beban:
            akun_data = df_neraca[df_neraca['akun'] == akun]
            if len(akun_data) > 0 and akun_data.iloc[0]['debit'] > 0:
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(akun)
                with col2:
                    st.write("")
                with col3:
                    st.write(f"Rp {akun_data.iloc[0]['debit']:,.0f}")
                ada_beban = True
        
        if not ada_beban:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.write("Tidak ada beban operasional")
            with col2:
                st.write("")
            with col3:
                st.write("Rp 0")
        
        st.markdown(f"**Total Beban Operasional: Rp {total_beban:,.0f}**")
        st.divider()
        
        st.markdown("###### HASIL BERSIH")
        if laba_rugi_bersih >= 0:
            st.success(f"**LABA BERSIH: Rp {laba_rugi_bersih:,.0f}**")
        else:
            st.error(f"**RUGI BERSIH: Rp {abs(laba_rugi_bersih):,.0f}**")
        
    except Exception as e:
        st.error(f"Error generating laba rugi report: {e}")
        st.info("Silakan refresh halaman atau pastikan sudah ada data transaksi")

def show_laporan_perubahan_modal():
    try:
       
        df_perubahan_modal, modal_awal, laba_rugi_bersih, prive, modal_akhir = get_laporan_perubahan_modal()
        
        save_perubahan_modal_to_excel(df_perubahan_modal, modal_awal, laba_rugi_bersih, prive, modal_akhir)
            
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col3:
            if st.button("📥 Download Excel", use_container_width=True, type="primary", key="download_modal"):
                download_laporan_perubahan_modal(df_perubahan_modal)
        
        st.markdown("#### 📊 Ringkasan Ekuitas")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="💰 Modal Awal",
                value=f"Rp {modal_awal:,.0f}",
                delta=None
            )
        
        with col2:
            if laba_rugi_bersih >= 0:
                st.metric(
                    label="📈 Laba Bersih",
                    value=f"Rp {laba_rugi_bersih:,.0f}",
                    delta=None
                )
            else:
                st.metric(
                    label="📉 Rugi Bersih", 
                    value=f"Rp {abs(laba_rugi_bersih):,.0f}",
                    delta=None
                )
        
        with col3:
            st.metric(
                label="💸 Prive",
                value=f"Rp {prive:,.0f}",
                delta=None
            )
        
        with col4:
            perubahan = modal_akhir - modal_awal
            st.metric(
                label="🎯 Modal Akhir",
                value=f"Rp {modal_akhir:,.0f}",
                delta=f"Rp {perubahan:,.0f}" if perubahan != 0 else None,
                delta_color="normal" if perubahan >= 0 else "inverse"
            )
        
        st.markdown("### Laporan Perubahan Modal")
        
        st.markdown("###### Modal Awal")
        st.info(f"**Rp {modal_awal:,.0f}**")
        st.divider()
        
        st.markdown("###### Laba/Rugi Bersih")
        if laba_rugi_bersih >= 0:
            st.success(f"**Rp {laba_rugi_bersih:,.0f}**")
            st.caption("Penambahan modal dari kegiatan operasional")
        else:
            st.error(f"**Rp {abs(laba_rugi_bersih):,.0f}**")
            st.caption("Pengurangan modal dari kegiatan operasional")
        st.divider()
        
        st.markdown("###### Prive")
        if prive > 0:
            st.warning(f"**Rp {prive:,.0f}**")
            st.caption("Pengambilan pribadi oleh pemilik")
        else:
            st.info("**Rp 0**")
            st.caption("Tidak ada pengambilan prive")
        st.divider()
        
        # Modal Akhir
        st.markdown("###### Modal Akhir")
        if modal_akhir > modal_awal:
            st.success(f"**Rp {modal_akhir:,.0f}**")
        elif modal_akhir < modal_awal:
            st.error(f"**Rp {modal_akhir:,.0f}**")
        else:
            st.info(f"**Rp {modal_akhir:,.0f}**")
        
    except Exception as e:
        st.error(f"Error generating laporan perubahan modal: {e}")
        st.info("Silakan refresh halaman atau pastikan sudah ada data transaksi")

def show_laporan_posisi_keuangan():
    try:
        df_posisi_keuangan, total_aktiva, total_kewajiban, total_modal, total_pasiva = get_laporan_posisi_keuangan()
        
        save_posisi_keuangan_to_excel(total_aktiva, total_kewajiban, total_modal, total_pasiva)
        
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col3:
            if st.button("📥 Download Excel", use_container_width=True, type="primary", key="download_posisi"):
                download_laporan_posisi_keuangan(df_posisi_keuangan)
        
        st.markdown("#### ⚖️ Ringkasan Neraca")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="💰 Total Aktiva",
                value=f"Rp {total_aktiva:,.0f}",
                delta=None
            )
        
        with col2:
            st.metric(
                label="📋 Total Kewajiban", 
                value=f"Rp {total_kewajiban:,.0f}",
                delta=None
            )
        
        with col3:
            st.metric(
                label="🏢 Total Modal",
                value=f"Rp {total_modal:,.0f}",
                delta=None
            )
        
        with col4:
            is_balanced = total_aktiva == total_pasiva
            st.metric(
                label="✅ Keseimbangan",
                value="BALANCE" if is_balanced else "TIDAK BALANCE",
                delta=None,
                delta_color="normal" if is_balanced else "off"
            )
        
        st.markdown("### Laporan Posisi Keuangan")
        
        col_aktiva, col_pasiva = st.columns(2)
        
        with col_aktiva:
            st.markdown("#### AKTIVA")
            st.markdown("---")
            
            # SEMUA ASET DIGABUNG
            st.markdown("##### Aset")
            df_neraca = get_neraca_saldo_data()
            semua_aset = ["Kas", "Bank", "Piutang Usaha", "Persediaan Sapi", "Perlengkapan",
                        "Peralatan", "Kandang", "Tanah", "Akumulasi Penyusutan Kendaraan",
                        "Akumulasi Penyusutan Kandang"]
            total_semua_aset = 0
            ada_aset = False
            for akun in semua_aset:
                akun_data = df_neraca[df_neraca['akun'] == akun]
                if len(akun_data) > 0 and akun_data.iloc[0]['debit'] > 0:
                    jumlah = akun_data.iloc[0]['debit']
                    total_semua_aset += jumlah
                    col1, col2 = st.columns([3, 2])
                    with col1:
                        st.write(akun)
                    with col2:
                        st.write(f"**Rp {jumlah:,.0f}**")
                    ada_aset = True
            
            if not ada_aset:
                st.info("Tidak ada aset")
            else:
                st.markdown(f"**Total Aset: Rp {total_semua_aset:,.0f}**")
            
            st.markdown("---")
            st.success(f"##### 🎯 TOTAL AKTIVA: Rp {total_aktiva:,.0f}")
        
        with col_pasiva:
            st.markdown("#### PASIVA")
            st.markdown("---")
            
            st.markdown("##### Kewajiban")
            kewajiban_akun = ["Utang Usaha", "Utang Bank"]
            total_kewajiban_detail = 0
            
            ada_kewajiban = False
            for akun in kewajiban_akun:
                akun_data = df_neraca[df_neraca['akun'] == akun]
                if len(akun_data) > 0 and akun_data.iloc[0]['kredit'] > 0:
                    jumlah = akun_data.iloc[0]['kredit']
                    total_kewajiban_detail += jumlah
                    col1, col2 = st.columns([3, 2])
                    with col1:
                        st.write(akun)
                    with col2:
                        st.write(f"**Rp {jumlah:,.0f}**")
                    ada_kewajiban = True
            
            if not ada_kewajiban:
                st.info("Tidak ada kewajiban")
            else:
                st.markdown(f"**Total Kewajiban: Rp {total_kewajiban_detail:,.0f}**")
            
            st.markdown("---")
            
            st.markdown("##### Ekuitas")
            
            modal_akun_data = df_neraca[df_neraca['akun'] == "Modal Pemilik"]
            modal_pemilik = modal_akun_data.iloc[0]['kredit'] if len(modal_akun_data) > 0 else 0
            
            if modal_pemilik > 0:
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.write("Modal Pemilik")
                with col2:
                    st.write(f"**Rp {modal_pemilik:,.0f}**")
            
            laba_rugi_bersih = calculate_laba_rugi_bersih(df_neraca)
            if laba_rugi_bersih != 0:
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.write("Laba/Rugi Bersih" if laba_rugi_bersih >= 0 else "Rugi Bersih")
                with col2:
                    if laba_rugi_bersih >= 0:
                        st.write(f"**Rp {laba_rugi_bersih:,.0f}**")
                    else:
                        st.write(f"**(Rp {abs(laba_rugi_bersih):,.0f})**")
            
            prive_akun_data = df_neraca[df_neraca['akun'] == "Prive"]
            prive = prive_akun_data.iloc[0]['debit'] if len(prive_akun_data) > 0 else 0
            
            if prive > 0:
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.write("Prive")
                with col2:
                    st.write(f"**(Rp {prive:,.0f})**")
            
            st.markdown(f"**Total Ekuitas: Rp {total_modal:,.0f}**")
            
            st.markdown("---")
            st.success(f"##### 🎯 TOTAL PASIVA: Rp {total_pasiva:,.0f}")
             
    except Exception as e:
        st.error(f"Error generating laporan posisi keuangan: {e}")
        st.info("Silakan refresh halaman atau pastikan sudah ada data transaksi")
    
def img_to_base64(img_path):
    """Convert image to base64 string"""
    with open(img_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def login_page():
    st.markdown("""<br><br>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        try:
            logo_base64 = img_to_base64("logo.png")
            logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="130" style="border-radius: 10px;">'
        except:
            logo_html = '<div style="font-size: 3.5rem; color: #667eea;">📊</div>'
        
        st.markdown(f"""
            <div class="login-container">
                <div style="display: flex; align-items: center; justify-content: center; margin-bottom: 20px;">
                    {logo_html}
                </div>
                <h2 style="color: #667eea; margin-bottom: 5px;">MooLance</h2>
                <p style="margin-bottom: 25px; color: #666;">Sistem Akuntansi Keuangan</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("👤 Username")
            password = st.text_input("🔒 Password", type="password")
            st.markdown("<br>", unsafe_allow_html=True)
            submitted = st.form_submit_button("Login", type="primary", use_container_width=True)
            
            if submitted:
                if username and password:
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = username
                    st.rerun()
                else:
                    st.error("Login Gagal")

        if st.button("📝 Daftar Akun Baru", use_container_width=True):
            st.session_state['page'] = 'register'
            st.rerun()

def register_page():
    st.markdown("""<br><br>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        try:
            logo_base64 = img_to_base64("logo.png")
            logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="130" style="border-radius: 10px;">'
        except:
            logo_html = '<div style="font-size: 3.5rem; color: #667eea;">📊</div>'
        
        st.markdown(f"""
            <div class="login-container">
                <div style="display: flex; align-items: center; justify-content: center; margin-bottom: 20px;">
                    {logo_html}
                </div>
                <h2 style="color: #667eea; margin-bottom: 5px;">MooLance</h2>
                <p style="margin-bottom: 25px; color: #666;">Sistem Akuntansi Keuangan</p>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("reg_form"):
            username = st.text_input("👤 Username Baru")
            password = st.text_input("🔒 Password", type="password")
            confirm = st.text_input("🔒 Konfirmasi Password", type="password")
            submitted = st.form_submit_button("Daftar", type="primary", use_container_width=True)
            
            if submitted:
                if password == confirm and len(password) >= 6:
                    try:
                        conn = get_db_connection()
                        pw_hash = hash_password(password)
                        conn.execute("INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
                                    (username, pw_hash, datetime.now()))
                        conn.commit()
                        conn.close()
                        st.success("Berhasil! Silakan Login.")
                        st.session_state['page'] = 'login'
                        st.rerun()
                    except:
                        st.error("Username sudah ada.")
                else:
                    st.warning("Password harus minimal 6 karakter dan sama dengan konfirmasi.")
        
        if st.button("⬅️ Kembali ke Login", use_container_width=True):
            st.session_state['page'] = 'login'
            st.rerun()

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP NOT NULL
        )
    """)
    conn.commit()
    conn.close()

def hash_password(password: str) -> str:
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100_000)
    return binascii.hexlify(salt).decode() + '$' + binascii.hexlify(dk).decode()

def verify_password(stored: str, provided_password: str) -> bool:
    try:
        salt_hex, hash_hex = stored.split("$")
    except ValueError:
        return False
    salt = binascii.unhexlify(salt_hex)
    dk = hashlib.pbkdf2_hmac('sha256', provided_password.encode('utf-8'), salt, 100_000)
    return binascii.hexlify(dk).decode() == hash_hex

def init_excel_db():
    try:
        pd.read_excel(EXCEL_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'tanggal', 'akun_debit', 'akun_kredit', 
            'nominal_debit', 'nominal_kredit', 'keterangan', 
            'created_at', 'created_by'
        ])
        df.to_excel(EXCEL_DB_PATH, index=False)

def get_jurnal_data():
    try:
        df = pd.read_excel(EXCEL_DB_PATH)
        df['tanggal'] = pd.to_datetime(df['tanggal']).dt.strftime('%Y-%m-%d')
        return df
    except FileNotFoundError:
        init_excel_db()
        return pd.DataFrame()

def get_akun_list():
    return [
        "Kas", "Bank", "Piutang Usaha", "Persediaan Sapi", "Perlengkapan",
        "Peralatan", "Kandang", "Tanah", "Kendaraan", "Akumulasi Penyusutan Kendaraan", 
        "Akumulasi Penyusutan Kandang", "Utang Usaha", "Utang Bank",
        "Modal Pemilik", "Prive", "Pendapatan Jasa", "Pendapatan Penjualan",
        "Beban Gaji", "Beban Sewa", "Beban Listrik", "Beban Air", 
        "Beban Telepon", "Beban Transportasi", "Beban Penyusutan Kendaraan", "Beban Penyusutan Kandang",
        "Harga Pokok Penjualan", "Beban Pakan Konsentrat", "Beban Obat dan Vitamin",
        "Beban Pakan Katul", "Beban Pakan Ampas Tahu", "Beban Pakan Jerami",
        "Beban Perawatan Kandang"  
    ]
    

def init_buku_besar_db():
    try:
        pd.read_excel(BUKU_BESAR_DB_PATH)
    except FileNotFoundError:
        with pd.ExcelWriter(BUKU_BESAR_DB_PATH, engine='openpyxl') as writer:
            for akun in get_akun_list():
                df = pd.DataFrame(columns=[
                    'id', 'tanggal', 'akun', 'keterangan', 'debit', 'kredit', 
                    'saldo_debit', 'saldo_kredit', 'created_at'
                ])
                df.to_excel(writer, sheet_name=akun[:31], index=False) 

def get_buku_besar_from_db(akun=None):
    try:
        if akun:
            df = pd.read_excel(BUKU_BESAR_DB_PATH, sheet_name=akun[:31])
            if 'tanggal' in df.columns and len(df) > 0:
                df['tanggal'] = pd.to_datetime(df['tanggal']).dt.strftime('%Y-%m-%d')
            return df
        else:
            all_data = []
            xl = pd.ExcelFile(BUKU_BESAR_DB_PATH)
            
            for sheet_name in xl.sheet_names:
                df_sheet = pd.read_excel(BUKU_BESAR_DB_PATH, sheet_name=sheet_name)
                if len(df_sheet) > 0:
                    if 'tanggal' in df_sheet.columns:
                        df_sheet['tanggal'] = pd.to_datetime(df_sheet['tanggal']).dt.strftime('%Y-%m-%d')
                    all_data.append(df_sheet)
            
            if all_data:
                return pd.concat(all_data, ignore_index=True)
            else:
                return pd.DataFrame()
                
    except FileNotFoundError:
        init_buku_besar_db()
        return pd.DataFrame()

def get_buku_besar_for_akun(akun):

    return get_buku_besar_from_db(akun)

def save_to_buku_besar_db(tanggal, akun, keterangan, debit, kredit, saldo_debit, saldo_kredit):
    try:
        df_akun = get_buku_besar_from_db(akun)
        
        new_id = len(df_akun) + 1 if len(df_akun) > 0 else 1
        
        new_entry = {
            'id': new_id,
            'tanggal': tanggal,
            'akun': akun,
            'keterangan': keterangan,
            'debit': debit,
            'kredit': kredit,
            'saldo_debit': saldo_debit,
            'saldo_kredit': saldo_kredit,
            'created_at': datetime.now()
        }
        
        df_akun = pd.concat([df_akun, pd.DataFrame([new_entry])], ignore_index=True)
        
        update_buku_besar_sheet(akun, df_akun)
        return True
        
    except Exception as e:
        st.error(f"Error saving to buku besar: {e}")
        return False

def update_buku_besar_sheet(akun, df_akun):
    try:
        xl = pd.ExcelFile(BUKU_BESAR_DB_PATH)
        sheet_data = {}
        
        for sheet_name in xl.sheet_names:
            if sheet_name != akun[:31]: 
                sheet_data[sheet_name] = pd.read_excel(BUKU_BESAR_DB_PATH, sheet_name=sheet_name)
        
        sheet_data[akun[:31]] = df_akun
        
        with pd.ExcelWriter(BUKU_BESAR_DB_PATH, engine='openpyxl') as writer:
            for sheet_name, data in sheet_data.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)
                
    except Exception as e:
        st.error(f"Error updating buku besar sheet: {e}")

def clear_buku_besar_db():
    try:
        with pd.ExcelWriter(BUKU_BESAR_DB_PATH, engine='openpyxl') as writer:
            for akun in get_akun_list():
                df = pd.DataFrame(columns=[
                    'id', 'tanggal', 'akun', 'keterangan', 'debit', 'kredit', 
                    'saldo_debit', 'saldo_kredit', 'created_at'
                ])
                df.to_excel(writer, sheet_name=akun[:31], index=False)
        return True
    except Exception as e:
        st.error(f"Error clearing buku besar: {e}")
        return False

def clear_akun_from_buku_besar(akun):
    try:
        df_empty = pd.DataFrame(columns=[
            'id', 'tanggal', 'akun', 'keterangan', 'debit', 'kredit', 
            'saldo_debit', 'saldo_kredit', 'created_at'
        ])
        
        update_buku_besar_sheet(akun, df_empty)
        
    except Exception as e:
        st.error(f"Error clearing akun from buku besar: {e}")

def generate_buku_besar_from_jurnal():
    
    jurnal_data = get_jurnal_data()
    ajp_data = get_ajp_data()
    
    clear_buku_besar_db()
    
    semua_jurnal = pd.concat([jurnal_data, ajp_data], ignore_index=True)
    
    if len(semua_jurnal) == 0:
        return pd.DataFrame()
    
    for _, jurnal in semua_jurnal.iterrows():
        tanggal = jurnal['tanggal']
        keterangan = jurnal['keterangan']
        
        save_to_buku_besar_db(
            tanggal,
            jurnal['akun_debit'],
            keterangan,
            jurnal['nominal_debit'],
            0,  # kredit = 0 untuk akun debit
            0,  # saldo_debit sementara 0
            0   # saldo_kredit sementara 0
        )

        save_to_buku_besar_db(
            tanggal,
            jurnal['akun_kredit'],
            keterangan,
            0,  # debit = 0 untuk akun kredit
            jurnal['nominal_kredit'],
            0,  # saldo_debit sementara 0
            0   # saldo_kredit sementara 0
        )
    
    for akun in get_akun_list():
        akun_data = get_buku_besar_from_db(akun)
        
        if len(akun_data) == 0:
            continue
            
        akun_data = akun_data.sort_values('tanggal')
        
        saldo_debit = 0
        saldo_kredit = 0
        
        clear_akun_from_buku_besar(akun)
        
        for _, row in akun_data.iterrows():
            debit = row['debit']
            kredit = row['kredit']
        
            if debit > 0:
                saldo_debit += debit
            if kredit > 0:
                saldo_kredit += kredit
            
            if saldo_debit > saldo_kredit:
                final_saldo_debit = saldo_debit - saldo_kredit
                final_saldo_kredit = 0
            else:
                final_saldo_debit = 0
                final_saldo_kredit = saldo_kredit - saldo_debit
            
            save_to_buku_besar_db(
                row['tanggal'],
                akun,
                row['keterangan'],
                debit,
                kredit,
                final_saldo_debit,
                final_saldo_kredit
            )
    
    return get_buku_besar_from_db()

def get_buku_besar_data(akun_terpilih=None):
    if akun_terpilih:
        return get_buku_besar_from_db(akun_terpilih)
    else:
        return get_buku_besar_from_db()

def save_jurnal_data(tanggal, akun_debit, akun_kredit, nominal_debit, nominal_kredit, keterangan, username):
    df = get_jurnal_data()
    
    new_id = len(df) + 1 if len(df) > 0 else 1
    
    new_entry = {
        'id': new_id,
        'tanggal': tanggal,
        'akun_debit': akun_debit,
        'akun_kredit': akun_kredit,
        'nominal_debit': nominal_debit,
        'nominal_kredit': nominal_kredit,
        'keterangan': keterangan,
        'created_at': datetime.now(),
        'created_by': username
    }
    
    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    df.to_excel(EXCEL_DB_PATH, index=False)
    
    generate_buku_besar_from_jurnal()
    generate_neraca_saldo_from_buku_besar()
    
    return True

def delete_jurnal_data(jurnal_id):
    try:
        df = get_jurnal_data()
        if len(df) == 0:
            return False
        
        if jurnal_id not in df['id'].values:
            st.error(f"Jurnal ID {jurnal_id} tidak ditemukan")
            return False
        
        df = df[df['id'] != jurnal_id]
        
        df = df.reset_index(drop=True)
        df['id'] = df.index + 1
        
        df.to_excel(EXCEL_DB_PATH, index=False)
        
        generate_buku_besar_from_jurnal()
        generate_neraca_saldo_from_buku_besar()
        
        return True
        
    except Exception as e:
        st.error(f"Error menghapus jurnal: {e}")
        return False

def init_neraca_saldo_db():
    try:
        pd.read_excel(NERACA_SALDO_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'akun', 'debit', 'kredit', 'created_at'
        ])
        df.to_excel(NERACA_SALDO_DB_PATH, index=False)

def get_neraca_saldo_from_db():
    try:
        df = pd.read_excel(NERACA_SALDO_DB_PATH)
        return df
    except FileNotFoundError:
        init_neraca_saldo_db()
        return pd.DataFrame()

def save_neraca_saldo_to_db(neraca_data):
    try:
        df = pd.DataFrame(columns=['id', 'akun', 'debit', 'kredit', 'created_at'])
        
        for i, (akun, debit, kredit) in enumerate(neraca_data, 1):
            new_entry = {
                'id': i,
                'akun': akun,
                'debit': debit,
                'kredit': kredit,
                'created_at': datetime.now()
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        
        df.to_excel(NERACA_SALDO_DB_PATH, index=False)
        return True
    except Exception as e:
        st.error(f"Error saving neraca saldo: {e}")
        return False

def generate_neraca_saldo_from_buku_besar():
    try:
        df_buku_besar = get_buku_besar_from_db()
        
        if len(df_buku_besar) == 0:
            save_neraca_saldo_to_db([])
            return []
        
        neraca_data = []
        
        semua_akun = get_akun_list()
        
        for akun in semua_akun:
            akun_data = df_buku_besar[df_buku_besar['akun'] == akun]
            
            if len(akun_data) == 0:

                continue
                
            akun_data_sorted = akun_data.sort_values('tanggal')
            
            if len(akun_data_sorted) == 0:
                continue
                
            last_entry = akun_data_sorted.iloc[-1]
            
            saldo_debit = float(last_entry['saldo_debit']) if pd.notna(last_entry['saldo_debit']) and last_entry['saldo_debit'] != 0 else 0
            saldo_kredit = float(last_entry['saldo_kredit']) if pd.notna(last_entry['saldo_kredit']) and last_entry['saldo_kredit'] != 0 else 0
            
            if saldo_debit > 0:
                debit = saldo_debit
                kredit = 0
            elif saldo_kredit > 0:
                debit = 0
                kredit = saldo_kredit
            else:
                continue
            
            if debit > 0 or kredit > 0:
                neraca_data.append((akun, debit, kredit))
                
        save_neraca_saldo_to_db(neraca_data)
        
        return neraca_data
        
    except Exception as e:
        st.error(f"Error generating neraca saldo: {e}")
        return []

def get_neraca_saldo_data():
    df = get_neraca_saldo_from_db()
    
    if len(df) == 0:
        generate_neraca_saldo_from_buku_besar()
        df = get_neraca_saldo_from_db()
    
    return df

def get_buku_besar_data(akun_terpilih=None):
    jurnal_data = get_jurnal_data()
    
    if len(jurnal_data) == 0:
        return pd.DataFrame()
    
    if akun_terpilih:
        jurnal_data = jurnal_data[
            (jurnal_data['akun_debit'] == akun_terpilih) | 
            (jurnal_data['akun_kredit'] == akun_terpilih)
        ]
    
    if len(jurnal_data) == 0:
        return pd.DataFrame()
    
    buku_besar_entries = []
    
    for _, jurnal in jurnal_data.iterrows():
        tanggal = jurnal['tanggal']
        keterangan = jurnal['keterangan']
    
        if not akun_terpilih or jurnal['akun_debit'] == akun_terpilih:
            buku_besar_entries.append({
                'tanggal': tanggal,
                'keterangan': keterangan,
                'akun': jurnal['akun_debit'],
                'debit': jurnal['nominal_debit'],
                'kredit': 0,
                'saldo_debit': 0,
                'saldo_kredit': 0
            })
        
        if not akun_terpilih or jurnal['akun_kredit'] == akun_terpilih:
            buku_besar_entries.append({
                'tanggal': tanggal,
                'keterangan': keterangan,
                'akun': jurnal['akun_kredit'],
                'debit': 0,
                'kredit': jurnal['nominal_kredit'],
                'saldo_debit': 0,
                'saldo_kredit': 0
            })
    
    if not buku_besar_entries:
        return pd.DataFrame()
    
    df_buku_besar = pd.DataFrame(buku_besar_entries)
    df_buku_besar['tanggal'] = pd.to_datetime(df_buku_besar['tanggal'])
    df_buku_besar = df_buku_besar.sort_values(['akun', 'tanggal'])
    
    result_data = []
    
    for akun in df_buku_besar['akun'].unique():
        akun_data = df_buku_besar[df_buku_besar['akun'] == akun].copy()
        saldo = 0
        saldo_posisi = None  
        
        for idx, row in akun_data.iterrows():
            debit = row['debit']
            kredit = row['kredit']
            
            if saldo_posisi is None:
                if debit > 0:
                    saldo_posisi = "debit"
                    saldo = debit
                elif kredit > 0:
                    saldo_posisi = "kredit" 
                    saldo = kredit
            else:
                if saldo_posisi == "debit":
                    saldo = saldo + debit - kredit
                else:
                    saldo = saldo + kredit - debit
            
            if saldo_posisi == "debit":
                saldo_debit = saldo
                saldo_kredit = 0
            else:
                saldo_debit = 0
                saldo_kredit = saldo
            
            result_data.append({
                'tanggal': row['tanggal'].strftime('%Y-%m-%d'),
                'keterangan': row['keterangan'],
                'debit': debit,
                'kredit': kredit,
                'saldo_debit': saldo_debit,
                'saldo_kredit': saldo_kredit,
                'akun': akun
            })
    
    result_df = pd.DataFrame(result_data)
    
    result_df['tanggal'] = pd.to_datetime(result_df['tanggal'])
    result_df = result_df.sort_values('tanggal')
    result_df['tanggal'] = result_df['tanggal'].dt.strftime('%Y-%m-%d')
    
    return result_df

def init_ajp_db():
    try:
        pd.read_excel(AJP_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'tanggal', 'akun_debit', 'akun_kredit', 
            'nominal_debit', 'nominal_kredit', 'keterangan', 
            'created_at', 'created_by'
        ])
        df.to_excel(AJP_DB_PATH, index=False)

def get_ajp_data():
    try:
        df = pd.read_excel(AJP_DB_PATH)
        df['tanggal'] = pd.to_datetime(df['tanggal']).dt.strftime('%Y-%m-%d')
        return df
    except FileNotFoundError:
        init_ajp_db()
        return pd.DataFrame()

def save_ajp_data(tanggal, akun_debit, akun_kredit, nominal_debit, nominal_kredit, keterangan, username):
    df = get_ajp_data()
    
    new_id = len(df) + 1 if len(df) > 0 else 1
    
    new_entry = {
        'id': new_id,
        'tanggal': tanggal,
        'akun_debit': akun_debit,
        'akun_kredit': akun_kredit,
        'nominal_debit': nominal_debit,
        'nominal_kredit': nominal_kredit,
        'keterangan': keterangan,
        'created_at': datetime.now(),
        'created_by': username
    }
    
    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    df.to_excel(AJP_DB_PATH, index=False)
    
    generate_buku_besar_from_jurnal()
    generate_neraca_saldo_from_buku_besar()
    
    return True

def delete_ajp_data(ajp_id):
    try:
        df = get_ajp_data()
        if len(df) == 0:
            return False
        
        if ajp_id not in df['id'].values:
            st.error(f"Jurnal Penyesuaian ID {ajp_id} tidak ditemukan")
            return False
        
        df = df[df['id'] != ajp_id]
        
        df = df.reset_index(drop=True)
        df['id'] = df.index + 1
        
        df.to_excel(AJP_DB_PATH, index=False)
        
        generate_buku_besar_from_jurnal()
        generate_neraca_saldo_from_buku_besar()
        
        st.success("✅ Jurnal Penyesuaian berhasil dihapus dan semua laporan telah di-update!")
        return True
        
    except Exception as e:
        st.error(f"Error menghapus jurnal penyesuaian: {e}")
        return False
    
def init_persediaan_db():
    try:
        pd.read_excel(PERSEDIAAN_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'tanggal', 'keterangan', 'jenis', 'akun_tambahan', 
            'quantity', 'harga', 'total', 'created_at', 'created_by'
        ])
        df.to_excel(PERSEDIAAN_DB_PATH, index=False)

def get_persediaan_data():
    try:
        df = pd.read_excel(PERSEDIAAN_DB_PATH)
        df['tanggal'] = pd.to_datetime(df['tanggal']).dt.strftime('%Y-%m-%d')
        return df
    except FileNotFoundError:
        init_persediaan_db()
        return pd.DataFrame()

def save_persediaan_data(tanggal, keterangan, jenis, akun_tambahan, quantity, harga, total, username):
    df = get_persediaan_data()
    
    new_id = len(df) + 1 if len(df) > 0 else 1
    
    new_entry = {
        'id': new_id,
        'tanggal': tanggal,
        'keterangan': keterangan,
        'jenis': jenis,  # 'pembelian' atau 'penjualan'
        'akun_tambahan': akun_tambahan,  # 'Kas', 'Utang', 'Piutang'
        'quantity': quantity,
        'harga': harga,
        'total': total,
        'created_at': datetime.now(),
        'created_by': username
    }
    
    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    df.to_excel(PERSEDIAAN_DB_PATH, index=False)
    
    create_jurnal_from_persediaan(tanggal, keterangan, jenis, akun_tambahan, quantity, harga, total, username)
    
    return True

def delete_persediaan_data(persediaan_id):
    try:
        df_persediaan = get_persediaan_data()
        if len(df_persediaan) == 0:
            return False
        
        if persediaan_id not in df_persediaan['id'].values:
            st.error(f"Data Persediaan ID {persediaan_id} tidak ditemukan")
            return False
        
        transaksi = df_persediaan[df_persediaan['id'] == persediaan_id].iloc[0]
        tanggal = transaksi['tanggal']
        keterangan = transaksi['keterangan']
        jenis = transaksi['jenis']
        quantity = transaksi['quantity']
        harga = transaksi['harga']
        
        df_persediaan = df_persediaan[df_persediaan['id'] != persediaan_id]
    
        df_persediaan = df_persediaan.reset_index(drop=True)
        df_persediaan['id'] = df_persediaan.index + 1
        
        df_persediaan.to_excel(PERSEDIAAN_DB_PATH, index=False)
        
        delete_jurnal_related_to_persediaan(tanggal, keterangan, jenis, quantity, harga)
        
        generate_buku_besar_from_jurnal()
        generate_neraca_saldo_from_buku_besar()
        
        st.success("✅ Data Persediaan dan jurnal terkait berhasil dihapus! Semua laporan telah di-update.")
        return True
        
    except Exception as e:
        st.error(f"Error menghapus data persediaan: {e}")
        return False

def delete_jurnal_related_to_persediaan(tanggal, keterangan, jenis, quantity, harga):
    try:
        df_jurnal = get_jurnal_data()
        if len(df_jurnal) == 0:
            return

        patterns = []
        
        if jenis == 'pembelian':
            patterns.append(f"[PERSEDIAAN] {keterangan} - {quantity} unit")
        elif jenis == 'penjualan':
            patterns.append(f"[PERSEDIAAN] {keterangan} - {quantity} unit")
            patterns.append(f"[HPP] {keterangan} - {quantity} unit")
        
        jurnal_to_delete = []
        for pattern in patterns:
            for idx, jurnal in df_jurnal.iterrows():
                if pattern in jurnal['keterangan']:
                    jurnal_to_delete.append(jurnal['id'])
        
        if jurnal_to_delete:
            df_jurnal = df_jurnal[~df_jurnal['id'].isin(jurnal_to_delete)]
            
            df_jurnal = df_jurnal.reset_index(drop=True)
            df_jurnal['id'] = df_jurnal.index + 1
            
            df_jurnal.to_excel(EXCEL_DB_PATH, index=False)
            
            st.info(f"🗑️ Menghapus {len(jurnal_to_delete)} jurnal terkait")
        
    except Exception as e:
        st.error(f"Error menghapus jurnal terkait: {e}")

def create_jurnal_from_persediaan(tanggal, keterangan, jenis, akun_tambahan, quantity, harga, total, username):
    
    keterangan_persediaan = f"[PERSEDIAAN] {keterangan} - {quantity} unit"
    
    if jenis == 'pembelian':
        save_jurnal_data(
            tanggal=tanggal,
            akun_debit="Persediaan Sapi",
            akun_kredit=akun_tambahan,
            nominal_debit=total,
            nominal_kredit=total,
            keterangan=keterangan_persediaan,
            username=username
        )
    
    elif jenis == 'penjualan':
        hpp_per_unit = calculate_hpp_average()
        total_hpp = hpp_per_unit * quantity
        
        save_jurnal_data(
            tanggal=tanggal,
            akun_debit=akun_tambahan,
            akun_kredit="Pendapatan Penjualan",
            nominal_debit=total,
            nominal_kredit=total,
            keterangan=keterangan_persediaan,
            username=username
        )
    
        save_jurnal_data(
            tanggal=tanggal,
            akun_debit="Harga Pokok Penjualan",
            akun_kredit="Persediaan Sapi",
            nominal_debit=total_hpp,
            nominal_kredit=total_hpp,
            keterangan=f"[HPP] {keterangan} - {quantity} unit",
            username=username
        )
        
def calculate_hpp_average():
    df = get_kartu_persediaan_data()
    
    if len(df) == 0:
        return 0
    
    last_balance = df.iloc[-1]
    return last_balance.get('balance_price', 0)

def get_kartu_persediaan_data():
    df = get_persediaan_data()
    
    if len(df) == 0:
        return pd.DataFrame()
    
    df = df.sort_values('tanggal')
    
    kartu_data = []
    saldo_quantity = 0
    saldo_total = 0
    saldo_harga_rata2 = 0
    
    total_out_value = 0
    
    for _, transaksi in df.iterrows():
        tanggal = transaksi['tanggal']
        keterangan = transaksi['keterangan']
        jenis = transaksi['jenis']
        quantity = transaksi['quantity']
        harga = transaksi['harga']
        total_transaksi = transaksi['total']
        
        if jenis == 'pembelian':
            saldo_quantity += quantity
            saldo_total += total_transaksi
            saldo_harga_rata2 = saldo_total / saldo_quantity if saldo_quantity > 0 else 0
            
            kartu_data.append({
                'tanggal': tanggal,
                'keterangan': keterangan,
                'in_qty': quantity,
                'in_price': harga,
                'in_total': total_transaksi,
                'out_qty': 0,
                'out_price': 0,
                'out_total': 0,
                'balance_qty': saldo_quantity,
                'balance_price': saldo_harga_rata2,
                'balance_total': saldo_total,
                'total_out_accumulated': total_out_value  # Total OUT sampai transaksi ini
            })
            
        elif jenis == 'penjualan':
            hpp_per_unit = saldo_harga_rata2
            hpp_total = hpp_per_unit * quantity
            
            total_out_value += hpp_total
            
            saldo_quantity -= quantity
            saldo_total -= hpp_total
            saldo_harga_rata2 = saldo_total / saldo_quantity if saldo_quantity > 0 else 0
            
            kartu_data.append({
                'tanggal': tanggal,
                'keterangan': keterangan,
                'in_qty': 0,
                'in_price': 0,
                'in_total': 0,
                'out_qty': quantity,
                'out_price': hpp_per_unit,  # HPP per unit
                'out_total': hpp_total,     # Total HPP untuk penjualan ini
                'balance_qty': saldo_quantity,
                'balance_price': saldo_harga_rata2,
                'balance_total': saldo_total,
                'total_out_accumulated': total_out_value  # Total OUT sampai transaksi ini
            })
    
    result_df = pd.DataFrame(kartu_data)
    return result_df

def get_total_beban_hpp():
    df = get_kartu_persediaan_data()
    
    if len(df) == 0:
        return 0
    
    last_transaction = df.iloc[-1]
    return last_transaction.get('total_out_accumulated', 0)

def init_laporan_keuangan_db():
    try:
        pd.read_excel(LABA_RUGI_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'periode', 'akun', 'kategori', 'jumlah', 'created_at'
        ])
        df.to_excel(LABA_RUGI_DB_PATH, index=False)
    
    try:
        pd.read_excel(PERUBAHAN_MODAL_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'periode', 'keterangan', 'penambahan', 'pengurangan', 'jumlah_modal', 'created_at'
        ])
        df.to_excel(PERUBAHAN_MODAL_DB_PATH, index=False)
    
    try:
        pd.read_excel(POSISI_KEUANGAN_DB_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            'id', 'periode', 'kategori', 'akun', 'jumlah', 'created_at'
        ])
        df.to_excel(POSISI_KEUANGAN_DB_PATH, index=False)

def get_laba_rugi_from_db():
    try:
        df = pd.read_excel(LABA_RUGI_DB_PATH)
        return df
    except FileNotFoundError:
        init_laporan_keuangan_db()
        return pd.DataFrame()

def save_perubahan_modal_to_excel(modal_awal, laba_rugi_bersih, prive, modal_akhir):
    try:
        df = pd.DataFrame(columns=['id', 'periode', 'keterangan', 'penambahan', 'pengurangan', 'jumlah_modal', 'created_at'])
        
        periode = datetime.now().strftime('%Y-%m')
        
        perubahan_modal_data = []
        
        perubahan_modal_data.append({
            'keterangan': 'Modal Awal',
            'penambahan': 0,
            'pengurangan': 0,
            'jumlah_modal': modal_awal
        })
        
        if laba_rugi_bersih >= 0:
            perubahan_modal_data.append({
                'keterangan': 'Laba Bersih',
                'penambahan': laba_rugi_bersih,
                'pengurangan': 0,
                'jumlah_modal': laba_rugi_bersih
            })
        else:
            perubahan_modal_data.append({
                'keterangan': 'Rugi Bersih',
                'penambahan': 0,
                'pengurangan': abs(laba_rugi_bersih),
                'jumlah_modal': -abs(laba_rugi_bersih)
            })
        
        perubahan_modal_data.append({
            'keterangan': 'Prive',
            'penambahan': 0,
            'pengurangan': prive,
            'jumlah_modal': -prive
        })
        
        perubahan_modal_data.append({
            'keterangan': 'Modal Akhir',
            'penambahan': 0,
            'pengurangan': 0,
            'jumlah_modal': modal_akhir
        })
        
        for i, item in enumerate(perubahan_modal_data, 1):
            new_entry = {
                'id': i,
                'periode': periode,
                'keterangan': item['keterangan'],
                'penambahan': item['penambahan'],
                'pengurangan': item['pengurangan'],
                'jumlah_modal': item['jumlah_modal'],
                'created_at': datetime.now()
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        
        df.to_excel(PERUBAHAN_MODAL_DB_PATH, index=False)
        return True
        
    except Exception as e:
        st.error(f"Error saving laporan perubahan modal: {e}")
        return False
    
def get_laporan_laba_rugi():
    return generate_laba_rugi_from_neraca()

def generate_laba_rugi_from_neraca():
    df_neraca = get_neraca_saldo_data()
    
    if len(df_neraca) == 0:
        return create_empty_laba_rugi()
    
    total_pendapatan = calculate_total_pendapatan(df_neraca)
    total_hpp = calculate_total_hpp(df_neraca)
    total_beban = calculate_total_beban(df_neraca)
    laba_kotor = total_pendapatan - total_hpp
    laba_rugi_bersih = laba_kotor - total_beban
    
    df_laba_rugi = create_laba_rugi_dataframe(total_pendapatan, total_hpp, total_beban, laba_kotor, laba_rugi_bersih, df_neraca)
    
    return df_laba_rugi, total_pendapatan, total_beban, laba_rugi_bersih, laba_kotor, total_hpp

def calculate_total_pendapatan(df_neraca):
    akun_pendapatan = "Pendapatan Penjualan"
    akun_data = df_neraca[df_neraca['akun'] == akun_pendapatan]
    if len(akun_data) > 0:
        return akun_data.iloc[0]['kredit']
    return 0

def calculate_total_hpp(df_neraca):
    akun_hpp = "Harga Pokok Penjualan"
    akun_data = df_neraca[df_neraca['akun'] == akun_hpp]
    if len(akun_data) > 0:
        return akun_data.iloc[0]['debit']
    return 0

def calculate_total_beban(df_neraca):
    akun_beban = [
        "Beban Gaji", "Beban Sewa", "Beban Listrik", "Beban Air", 
        "Beban Telepon", "Beban Transportasi", "Beban Penyusutan Kendaraan",
        "Beban Pakan Konsentrat", "Beban Obat dan Vitamin", "Beban Penyusutan Kandang",
        "Beban Pakan Katul", "Beban Pakan Ampas Tahu", "Beban Pakan Jerami",
        "Beban Perawatan Kandang"  
]
    
    total_beban = 0
    for akun in akun_beban:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0:
            total_beban += akun_data.iloc[0]['debit']
    
    return total_beban

def create_laba_rugi_dataframe(total_pendapatan, total_hpp, total_beban, laba_kotor, laba_rugi_bersih, df_neraca):
    laba_rugi_data = []
    
    laba_rugi_data.append(("**PENDAPATAN USAHA**", "", ""))
    if total_pendapatan > 0:
        laba_rugi_data.append(("Pendapatan Penjualan", f"Rp {total_pendapatan:,.0f}", ""))
    else:
        laba_rugi_data.append(("Pendapatan Penjualan", "Rp 0", ""))
    laba_rugi_data.append(("**Total Pendapatan Usaha**", f"**Rp {total_pendapatan:,.0f}**", ""))
    laba_rugi_data.append(("", "", ""))
    
    laba_rugi_data.append(("**HARGA POKOK PENJUALAN**", "", ""))
    if total_hpp > 0:
        laba_rugi_data.append(("Harga Pokok Penjualan", "", f"Rp {total_hpp:,.0f}"))
    else:
        laba_rugi_data.append(("Harga Pokok Penjualan", "", "Rp 0"))
    laba_rugi_data.append(("**Total Harga Pokok Penjualan**", "", f"**Rp {total_hpp:,.0f}**"))
    laba_rugi_data.append(("", "", ""))
    
    laba_rugi_data.append(("**LABA KOTOR**", f"**Rp {laba_kotor:,.0f}**", ""))
    laba_rugi_data.append(("", "", ""))
    
    laba_rugi_data.append(("**BEBAN OPERASIONAL**", "", ""))
    
    akun_beban = [
       "Beban Gaji", "Beban Sewa", "Beban Listrik", "Beban Air", 
        "Beban Telepon", "Beban Transportasi", "Beban Penyusutan Kendaraan",
        "Beban Pakan Konsentrat", "Beban Obat dan Vitamin", "Beban Penyusutan Kandang",
        "Beban Pakan Katul", "Beban Pakan Ampas Tahu", "Beban Pakan Jerami",
        "Beban Perawatan Kandang"  
    ]
    
    ada_beban = False
    for akun in akun_beban:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0 and akun_data.iloc[0]['debit'] > 0:
            laba_rugi_data.append((akun, "", f"Rp {akun_data.iloc[0]['debit']:,.0f}"))
            ada_beban = True
    
    if not ada_beban:
        laba_rugi_data.append(("Tidak ada beban operasional", "", "Rp 0"))
    
    laba_rugi_data.append(("**Total Beban Operasional**", "", f"**Rp {total_beban:,.0f}**"))
    laba_rugi_data.append(("", "", ""))
    
    if laba_rugi_bersih >= 0:
        laba_rugi_data.append(("**LABA BERSIH**", f"**Rp {laba_rugi_bersih:,.0f}**", ""))
    else:
        laba_rugi_data.append(("**RUGI BERSIH**", "", f"**(Rp {abs(laba_rugi_bersih):,.0f})**"))
    
    return pd.DataFrame(laba_rugi_data, columns=['Akun', 'Pendapatan', 'Beban'])

def create_empty_laba_rugi():
    """Buat laporan laba rugi kosong"""
    empty_data = [
        ("**PENDAPATAN USAHA**", "", ""),
        ("Pendapatan Penjualan", "Rp 0", ""),
        ("**Total Pendapatan Usaha**", "**Rp 0**", ""),
        ("", "", ""),
        ("**HARGA POKOK PENJUALAN**", "", ""),
        ("Harga Pokok Penjualan", "", "Rp 0"),
        ("**Total Harga Pokok Penjualan**", "", "**Rp 0**"),
        ("", "", ""),
        ("**LABA KOTOR**", "**Rp 0**", ""),
        ("", "", ""),
        ("**BEBAN OPERASIONAL**", "", ""),
        ("Tidak ada beban operasional", "", "Rp 0"),
        ("**Total Beban Operasional**", "", "**Rp 0**"),
        ("", "", ""),
        ("**LABA BERSIH**", "**Rp 0**", "")
    ]
    
    df_empty = pd.DataFrame(empty_data, columns=['Akun', 'Pendapatan', 'Beban'])
    return df_empty, 0, 0, 0, 0, 0

def save_laba_rugi_to_excel(total_pendapatan, total_hpp, total_beban, laba_kotor, laba_rugi_bersih):
    try:
        df = pd.DataFrame(columns=['id', 'periode', 'akun', 'kategori', 'jumlah', 'created_at'])
        
        periode = datetime.now().strftime('%Y-%m')
        
        laba_rugi_data = []
        
        laba_rugi_data.append({
            'akun': 'Pendapatan Penjualan',
            'kategori': 'pendapatan',
            'jumlah': total_pendapatan
        })
        
        laba_rugi_data.append({
            'akun': 'Harga Pokok Penjualan', 
            'kategori': 'beban',
            'jumlah': total_hpp
        })
        
        laba_rugi_data.append({
            'akun': 'Laba Kotor',
            'kategori': 'laba_kotor',
            'jumlah': laba_kotor
        })
        
        laba_rugi_data.append({
            'akun': 'Total Beban Operasional',
            'kategori': 'beban_operasional',
            'jumlah': total_beban
        })
        
        laba_rugi_data.append({
            'akun': 'Laba/Rugi Bersih',
            'kategori': 'laba_rugi_bersih',
            'jumlah': laba_rugi_bersih
        })
        
        for i, item in enumerate(laba_rugi_data, 1):
            new_entry = {
                'id': i,
                'periode': periode,
                'akun': item['akun'],
                'kategori': item['kategori'],
                'jumlah': item['jumlah'],
                'created_at': datetime.now()
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        
        df.to_excel(LABA_RUGI_DB_PATH, index=False)
        return True
        
    except Exception as e:
        st.error(f"Error saving laporan laba rugi: {e}")
        return False

def download_laporan_laba_rugi(df_laba_rugi):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_laba_rugi.to_excel(writer, sheet_name='Laporan Laba Rugi', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Laporan Laba Rugi']
            
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 20
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Laporan Laba Rugi (.xlsx)",
            data=output,
            file_name=f"Laporan_Laba_Rugi_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Error generating download: {e}")

def get_laporan_perubahan_modal():
    return generate_perubahan_modal_from_neraca()

def generate_perubahan_modal_from_neraca():
    df_neraca = get_neraca_saldo_data()
    
    if len(df_neraca) == 0:
        return create_empty_perubahan_modal()
    
    modal_awal = calculate_modal_awal(df_neraca)
    laba_rugi_bersih = calculate_laba_rugi_bersih(df_neraca)
    prive = calculate_prive(df_neraca)
    modal_akhir = calculate_modal_akhir(modal_awal, laba_rugi_bersih, prive)
    
    df_perubahan_modal = create_perubahan_modal_dataframe(modal_awal, laba_rugi_bersih, prive, modal_akhir)
    
    return df_perubahan_modal, modal_awal, laba_rugi_bersih, prive, modal_akhir

def calculate_modal_awal(df_neraca):
    akun_modal = "Modal Pemilik"
    akun_data = df_neraca[df_neraca['akun'] == akun_modal]
    if len(akun_data) > 0:
        saldo_kredit = akun_data.iloc[0]['kredit']
        saldo_debit = akun_data.iloc[0]['debit']
        return max(saldo_kredit - saldo_debit, 0)  
    return 0

def calculate_laba_rugi_bersih(df_neraca):
    total_pendapatan = calculate_total_pendapatan(df_neraca)
    total_hpp = calculate_total_hpp(df_neraca)
    total_beban = calculate_total_beban(df_neraca)
    laba_kotor = total_pendapatan - total_hpp
    return laba_kotor - total_beban

def calculate_prive(df_neraca):
    akun_prive = "Prive"
    akun_data = df_neraca[df_neraca['akun'] == akun_prive]
    if len(akun_data) > 0:
        return akun_data.iloc[0]['debit']  
    return 0

def calculate_modal_akhir(modal_awal, laba_rugi_bersih, prive):
    return modal_awal + laba_rugi_bersih - prive

def create_perubahan_modal_dataframe(modal_awal, laba_rugi_bersih, prive, modal_akhir):
    perubahan_modal_data = []
    
    perubahan_modal_data.append(("**LAPORAN PERUBAHAN MODAL**", "", ""))
    perubahan_modal_data.append(("", "", ""))
    
    perubahan_modal_data.append(("**Modal Awal**", "", f"**Rp {modal_awal:,.0f}**"))
    perubahan_modal_data.append(("", "", ""))
    
    if laba_rugi_bersih >= 0:
        perubahan_modal_data.append(("**Laba Bersih**", f"**Rp {laba_rugi_bersih:,.0f}**", ""))
    else:
        perubahan_modal_data.append(("**Rugi Bersih**", "", f"**(Rp {abs(laba_rugi_bersih):,.0f})**"))
    perubahan_modal_data.append(("", "", ""))
    
    perubahan_modal_data.append(("**Prive**", "", f"**(Rp {prive:,.0f})**"))
    perubahan_modal_data.append(("", "", ""))
    
    perubahan_modal_data.append(("**Perubahan Modal**", "", ""))
    if laba_rugi_bersih >= 0:
        perubahan_modal_data.append(("Penambahan dari Laba", f"Rp {laba_rugi_bersih:,.0f}", ""))
    else:
        perubahan_modal_data.append(("Pengurangan dari Rugi", "", f"Rp {abs(laba_rugi_bersih):,.0f}"))
    
    perubahan_modal_data.append(("Pengurangan untuk Prive", "", f"Rp {prive:,.0f}"))
    perubahan_modal_data.append(("", "", ""))
    
    perubahan_modal = laba_rugi_bersih - prive
    if perubahan_modal >= 0:
        perubahan_modal_data.append(("**Total Perubahan Modal**", f"**Rp {perubahan_modal:,.0f}**", ""))
    else:
        perubahan_modal_data.append(("**Total Perubahan Modal**", "", f"**(Rp {abs(perubahan_modal):,.0f})**"))
    perubahan_modal_data.append(("", "", ""))
    
    perubahan_modal_data.append(("**Modal Akhir**", "", f"**Rp {modal_akhir:,.0f}**"))
    
    return pd.DataFrame(perubahan_modal_data, columns=['Keterangan', 'Penambahan', 'Pengurangan'])

def create_empty_perubahan_modal():
    empty_data = [
        ("**LAPORAN PERUBAHAN MODAL**", "", ""),
        ("", "", ""),
        ("**Modal Awal**", "", "**Rp 0**"),
        ("", "", ""),
        ("**Laba Bersih**", "**Rp 0**", ""),
        ("", "", ""),
        ("**Prive**", "", "**(Rp 0)**"),
        ("", "", ""),
        ("**Perubahan Modal**", "", ""),
        ("Penambahan dari Laba", "Rp 0", ""),
        ("Pengurangan untuk Prive", "", "Rp 0"),
        ("", "", ""),
        ("**Total Perubahan Modal**", "**Rp 0**", ""),
        ("", "", ""),
        ("**Modal Akhir**", "", "**Rp 0**")
    ]
    
    df_empty = pd.DataFrame(empty_data, columns=['Keterangan', 'Penambahan', 'Pengurangan'])
    return df_empty, 0, 0, 0, 0

def download_laporan_perubahan_modal(df_perubahan_modal):
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_perubahan_modal.to_excel(writer, sheet_name='Laporan Perubahan Modal', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Laporan Perubahan Modal']
            
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 20
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Laporan Perubahan Modal (.xlsx)",
            data=output,
            file_name=f"Laporan_Perubahan_Modal_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Error generating download: {e}")
def save_perubahan_modal_to_excel(df_perubahan_modal, modal_awal, laba_rugi_bersih, prive, modal_akhir):
    try:
        df = pd.DataFrame(columns=['id', 'periode', 'keterangan', 'penambahan', 'pengurangan', 'created_at'])
        
        periode = datetime.now().strftime('%Y-%m')
        
        perubahan_modal_data = []
        
        perubahan_modal_data.append({
            'keterangan': 'Modal Awal',
            'penambahan': 0,
            'pengurangan': 0,
            'jumlah_modal': modal_awal
        })
        
        if laba_rugi_bersih >= 0:
            perubahan_modal_data.append({
                'keterangan': 'Laba Bersih',
                'penambahan': laba_rugi_bersih,
                'pengurangan': 0,
                'jumlah_modal': laba_rugi_bersih
            })
        else:
            perubahan_modal_data.append({
                'keterangan': 'Rugi Bersih',
                'penambahan': 0,
                'pengurangan': abs(laba_rugi_bersih),
                'jumlah_modal': -abs(laba_rugi_bersih)
            })
        
        perubahan_modal_data.append({
            'keterangan': 'Prive',
            'penambahan': 0,
            'pengurangan': prive,
            'jumlah_modal': -prive
        })
        
        perubahan_modal_data.append({
            'keterangan': 'Modal Akhir',
            'penambahan': 0,
            'pengurangan': 0,
            'jumlah_modal': modal_akhir
        })
        
        df_db = pd.DataFrame(columns=['id', 'periode', 'keterangan', 'penambahan', 'pengurangan', 'jumlah_modal', 'created_at'])
        
        for i, item in enumerate(perubahan_modal_data, 1):
            new_entry = {
                'id': i,
                'periode': periode,
                'keterangan': item['keterangan'],
                'penambahan': item['penambahan'],
                'pengurangan': item['pengurangan'],
                'jumlah_modal': item['jumlah_modal'],
                'created_at': datetime.now()
            }
            df_db = pd.concat([df_db, pd.DataFrame([new_entry])], ignore_index=True)
        
        df_db.to_excel(PERUBAHAN_MODAL_DB_PATH, index=False)
        return True
        
    except Exception as e:
        st.error(f"Error saving laporan perubahan modal: {e}")
        return False

def get_perubahan_modal_from_db():
    try:
        df = pd.read_excel(PERUBAHAN_MODAL_DB_PATH)
        return df
    except FileNotFoundError:
        init_laporan_keuangan_db()
        return pd.DataFrame()
    
def get_laporan_posisi_keuangan():
    return generate_posisi_keuangan_from_neraca()

def generate_posisi_keuangan_from_neraca():
    df_neraca = get_neraca_saldo_data()
    
    if len(df_neraca) == 0:
        return create_empty_posisi_keuangan()
    
    total_aktiva = calculate_total_aktiva(df_neraca)
    total_kewajiban = calculate_total_kewajiban(df_neraca)
    total_modal = calculate_total_modal(df_neraca)
    total_pasiva = total_kewajiban + total_modal
    
    df_posisi_keuangan = create_posisi_keuangan_dataframe(df_neraca, total_aktiva, total_kewajiban, total_modal, total_pasiva)
    
    return df_posisi_keuangan, total_aktiva, total_kewajiban, total_modal, total_pasiva

def calculate_total_aktiva(df_neraca):
    akun_aktiva = [
        "Kas", "Bank", "Piutang Usaha", "Persediaan Sapi", "Perlengkapan",
        "Peralatan", "Kandang", "Tanah", "Akumulasi Penyusutan Kendaraan"
        "Akumulasi Penyusutan Kandang"
    ]
    
    total_aktiva = 0
    for akun in akun_aktiva:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0:
            total_aktiva += akun_data.iloc[0]['debit']
    
    return total_aktiva

def calculate_total_kewajiban(df_neraca):
    akun_kewajiban = ["Utang Usaha", "Utang Bank"]
    
    total_kewajiban = 0
    for akun in akun_kewajiban:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0:
            total_kewajiban += akun_data.iloc[0]['kredit']
    
    return total_kewajiban

def calculate_total_modal(df_neraca):
    akun_modal = ["Modal Pemilik", "Prive"]
    
    total_modal = 0
    for akun in akun_modal:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0:
            if akun == "Modal Pemilik":
                total_modal += akun_data.iloc[0]['kredit']
            elif akun == "Prive":
                total_modal -= akun_data.iloc[0]['debit']
    
    laba_rugi_bersih = calculate_laba_rugi_bersih(df_neraca)
    total_modal += laba_rugi_bersih
    
    return max(total_modal, 0)  

def create_posisi_keuangan_dataframe(df_neraca, total_aktiva, total_kewajiban, total_modal, total_pasiva):
    posisi_keuangan_data = []
    
    posisi_keuangan_data.append(("**AKTIVA**", "", "**PASIVA**", ""))
    posisi_keuangan_data.append(("", "", "", ""))
    
    posisi_keuangan_data.append(("**ASET**", "", "**KEWAJIBAN**", ""))
    
    akun_aset = ["Kas", "Bank", "Piutang Usaha", "Persediaan Sapi", "Perlengkapan", "Peralatan", "Gedung", "Tanah", "Akumulasi Penyusutan Kendaraan"]
    total_aset = 0
    
    for akun in akun_aset:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0 and akun_data.iloc[0]['debit'] > 0:
            jumlah = akun_data.iloc[0]['debit']
            total_aset += jumlah
            posisi_keuangan_data.append((akun, f"Rp {jumlah:,.0f}", "", ""))
    
    if total_aset > 0:
        posisi_keuangan_data.append(("**Total Aset**", f"**Rp {total_aset:,.0f}**", "", ""))
    else:
        posisi_keuangan_data.append(("Tidak ada aset", "Rp 0", "", ""))
    
    posisi_keuangan_data.append(("", "", "", ""))
    
    kewajiban_akun = ["Utang Usaha", "Utang Bank"]
    total_kewajiban_detail = 0
    
    for akun in kewajiban_akun:
        akun_data = df_neraca[df_neraca['akun'] == akun]
        if len(akun_data) > 0 and akun_data.iloc[0]['kredit'] > 0:
            jumlah = akun_data.iloc[0]['kredit']
            total_kewajiban_detail += jumlah
            posisi_keuangan_data.append(("", "", akun, f"Rp {jumlah:,.0f}"))
    
    if total_kewajiban_detail > 0:
        posisi_keuangan_data.append(("", "", "**Total Kewajiban**", f"**Rp {total_kewajiban_detail:,.0f}**"))
    else:
        posisi_keuangan_data.append(("", "", "Tidak ada kewajiban", "Rp 0"))
    
    posisi_keuangan_data.append(("", "", "", ""))
    
    posisi_keuangan_data.append(("", "", "**EKUITAS**", ""))
    
    modal_akun_data = df_neraca[df_neraca['akun'] == "Modal Pemilik"]
    modal_pemilik = modal_akun_data.iloc[0]['kredit'] if len(modal_akun_data) > 0 else 0
    
    if modal_pemilik > 0:
        posisi_keuangan_data.append(("", "", "Modal Pemilik", f"Rp {modal_pemilik:,.0f}"))
    
    laba_rugi_bersih = calculate_laba_rugi_bersih(df_neraca)
    if laba_rugi_bersih >= 0:
        posisi_keuangan_data.append(("", "", "Laba Bersih", f"Rp {laba_rugi_bersih:,.0f}"))
    else:
        posisi_keuangan_data.append(("", "", "Rugi Bersih", f"(Rp {abs(laba_rugi_bersih):,.0f})"))
    
    prive_akun_data = df_neraca[df_neraca['akun'] == "Prive"]
    prive = prive_akun_data.iloc[0]['debit'] if len(prive_akun_data) > 0 else 0
    
    if prive > 0:
        posisi_keuangan_data.append(("", "", "Prive", f"(Rp {prive:,.0f})"))
    
    posisi_keuangan_data.append(("", "", "**Total Ekuitas**", f"**Rp {total_modal:,.0f}**"))
    
    posisi_keuangan_data.append(("", "", "", ""))
    
    posisi_keuangan_data.append(("**TOTAL AKTIVA**", f"**Rp {total_aktiva:,.0f}**", "**TOTAL PASIVA**", f"**Rp {total_pasiva:,.0f}**"))
    
    return pd.DataFrame(posisi_keuangan_data, columns=['Aktiva', 'Jumlah_Aktiva', 'Pasiva', 'Jumlah_Pasiva'])

def create_empty_posisi_keuangan():
    empty_data = [
        ("**AKTIVA**", "", "**PASIVA**", ""),
        ("", "", "", ""),
        ("**ASET**", "", "**KEWAJIBAN**", ""),
        ("Tidak ada aset", "Rp 0", "Tidak ada kewajiban", "Rp 0"),
        ("", "", "", ""),
        ("", "", "**EKUITAS**", ""),
        ("", "", "Modal Pemilik", "Rp 0"),
        ("", "", "Laba Bersih", "Rp 0"),
        ("", "", "**Total Ekuitas**", "**Rp 0**"),
        ("", "", "", ""),
        ("**TOTAL AKTIVA**", "**Rp 0**", "**TOTAL PASIVA**", "**Rp 0**")
    ]
    
    df_empty = pd.DataFrame(empty_data, columns=['Aktiva', 'Jumlah_Aktiva', 'Pasiva', 'Jumlah_Pasiva'])
    return df_empty, 0, 0, 0, 0

def save_posisi_keuangan_to_excel(total_aktiva, total_kewajiban, total_modal, total_pasiva):
    try:
        df = pd.DataFrame(columns=['id', 'periode', 'kategori', 'akun', 'jumlah', 'created_at'])
        
        periode = datetime.now().strftime('%Y-%m')
        
        posisi_keuangan_data = []
        
        posisi_keuangan_data.append({
            'kategori': 'aktiva',
            'akun': 'Total Aktiva',
            'jumlah': total_aktiva
        })
        
        posisi_keuangan_data.append({
            'kategori': 'kewajiban', 
            'akun': 'Total Kewajiban',
            'jumlah': total_kewajiban
        })
        
        posisi_keuangan_data.append({
            'kategori': 'modal',
            'akun': 'Total Modal',
            'jumlah': total_modal
        })
        
        posisi_keuangan_data.append({
            'kategori': 'pasiva',
            'akun': 'Total Pasiva',
            'jumlah': total_pasiva
        })
        
        for i, item in enumerate(posisi_keuangan_data, 1):
            new_entry = {
                'id': i,
                'periode': periode,
                'kategori': item['kategori'],
                'akun': item['akun'],
                'jumlah': item['jumlah'],
                'created_at': datetime.now()
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        
        df.to_excel(POSISI_KEUANGAN_DB_PATH, index=False)
        return True
        
    except Exception as e:
        st.error(f"Error saving laporan posisi keuangan: {e}")
        return False

def get_posisi_keuangan_from_db():
    try:
        df = pd.read_excel(POSISI_KEUANGAN_DB_PATH)
        return df
    except FileNotFoundError:
        init_laporan_keuangan_db()
        return pd.DataFrame()

def download_laporan_posisi_keuangan(df_posisi_keuangan):
    try:
        _, total_aktiva, total_kewajiban, total_modal, total_pasiva = get_laporan_posisi_keuangan()
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_posisi_keuangan.to_excel(writer, sheet_name='Laporan Posisi Keuangan', index=False)
            
            raw_data = {
                'Komponen': ['Total Aktiva', 'Total Kewajiban', 'Total Modal', 'Total Pasiva'],
                'Nilai': [total_aktiva, total_kewajiban, total_modal, total_pasiva]
            }
            df_raw = pd.DataFrame(raw_data)
            df_raw.to_excel(writer, sheet_name='Data Raw', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Laporan Posisi Keuangan']
            
            worksheet.column_dimensions['A'].width = 30
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 30
            worksheet.column_dimensions['D'].width = 20
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Laporan Posisi Keuangan (.xlsx)",
            data=output,
            file_name=f"Laporan_Posisi_Keuangan_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Error generating download: {e}")

init_db()
init_excel_db()
init_buku_besar_db()
init_neraca_saldo_db() 
init_ajp_db()
init_persediaan_db()
init_laporan_keuangan_db()
load_css()

if st.session_state['logged_in']:
    main_dashboard()
else:
    if st.session_state['page'] == 'register':
        register_page()
    else:
        login_page()