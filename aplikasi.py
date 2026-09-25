import streamlit as st
import pandas as pd
import io
import os

# 1. KONDISI AWAL HALAMAN WEB
st.set_page_config(page_title="Sistem Keamanan Koperasi RSUD", layout="wide")

# =====================================================================
# 🔐 SISTEM FITUR PENGUNCI KATA SANDI (LOGIN)
# =====================================================================
# Silakan ganti kata "koperasi2026" di bawah ini dengan kata sandi rahasia pilihan Anda
KATA_SANDI_BENAR = "koperasi2026"

# Inisialisasi status login di memori aplikasi
if "sudah_login" not in st.session_state:
    st.session_state.sudah_login = False

# Tampilan halaman login jika pengguna belum sukses masuk
if not st.session_state.sudah_login:
    st.markdown("<h2 style='text-align: center;'>🔐 Gerbang Keamanan Koperasi SSM</h2>", unsafe_allow_index=False)
    st.markdown("<h5 style='text-align: center;'>RSUD Kota Tangerang</h5>", unsafe_allow_index=False)
    st.write("---")
    
    # Kotak dialog login di tengah layar
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("form_login"):
            input_password = st.text_input("🔑 Masukkan Kata Sandi Akses Pengurus:", type="password", placeholder="Ketik kata sandi di sini...")
            tombol_masuk = st.form_submit_button("🔓 Buka Akses Sistem")
            
            if tombol_masuk:
                if input_password == KATA_SANDI_BENAR:
                    st.session_state.sudah_login = True
                    st.success("✅ Kata sandi benar! Membuka database...")
                    st.rerun()
                else:
                    st.error("❌ Kata sandi salah! Akses ditolak.")
    st.stop() # Menghentikan seluruh kode di bawah agar tidak dieksekusi sebelum login

# =====================================================================
# 🏛️ KODE UTAMA APLIKASI KOPERASI (HANYA JALAN JIKA SUDAH LOGIN)
# =====================================================================
st.title("🏛️ Sistem Administrasi & Buku Kas Master Koperasi")
st.markdown("##### Unit Kerja: RSUD Kota Tangerang — Mode Akses Pengurus Resmi")
st.write("---")

# Tombol Keluar (Log Out) di pojok kiri bawah sidebar
if st.sidebar.button("🔒 Keluar dari Sistem (Log Out)"):
    st.session_state.sudah_login = False
    st.rerun()

# Nama file Excel Fisik sebagai Brankas Data Permanen
DATABASE_FISIK = "MASTER_DATABASE_KOPERASI.xlsx"

def muat_database_permanen():
    if os.path.exists(DATABASE_FISIK):
        return pd.read_excel(DATABASE_FISIK, converters={'NOMOR REKENING BANK': str, 'NOMOR REKENING': str, 'NO TELEPON': str, 'NOMOR TELEPON': str})
    return None

def standarkan_nama(nama_mentah):
    text = str(nama_mentah).upper().strip()
    return " ".join(text.split())

df_master = muat_database_permanen()

# PANEL NAVIGASI UTAMA (SIDEBAR)
menu_navigasi = st.sidebar.radio(
    "📂 Menu Navigasi:",
    ["📊 Dashboard Utama (Brankas)", "🔄 Sinkronisasi Iuran Bulanan", "🚀 Inisialisasi Data Awal"]
)

# =====================================================================
# MENU 0: INISIALISASI DATA AWAL (HANYA UNTUK PERTAMA KALI)
# =====================================================================
if menu_navigasi == "🚀 Inisialisasi Data Awal":
    st.subheader("🚀 Langkah Awal: Daftarkan File Master Pertama Anda")
    st.write("Gunakan menu ini **hanya sekali** untuk memasukkan file data PNS bersih Anda agar tertanam permanen di sistem.")
    
    file_mentah = st.file_uploader("Unggah File Excel Berisi 96 Anggota PNS Bersih Anda:", type=["xlsx"])
    if file_mentah is not None:
        if st.button("🔒 Tanamkan File Ini Sebagai Database Master Permanen"):
            df_init = pd.read_excel(file_mentah, converters={'NOMOR REKENING BANK': str, 'NOMOR REKENING': str, 'NO TELEPON': str, 'NOMOR TELEPON': str})
            df_init.columns = [str(c).strip().upper() for c in df_init.columns]
            df_init = df_init.fillna(0)
            
            df_init.to_excel(DATABASE_FISIK, index=False)
            st.success("🔒 Sukses! File master Anda kini telah tertanam paten di sistem komputer Anda. Silakan pindah ke menu 'Dashboard Utama (Brankas)'.")
            st.rerun()

# =====================================================================
# MENU 1: DASHBOARD UTAMA (MEMBACA DATA PERMANEN)
# =====================================================================
elif menu_navigasi == "📊 Dashboard Utama (Brankas)":
    if df_master is not None:
        df = df_master.copy()
        df["TOTAL SIMPANAN"] = df["SIMPANAN POKOK"] + df["SIMPANAN WAJIB"] + df["SIMPANAN SUKARELA"]
        
        if st.sidebar.button("🗑️ Ganti / Reset File Master"):
            if os.path.exists(DATABASE_FISIK):
                os.remove(DATABASE_FISIK)
            st.session_state.clear()
            st.success("Database berhasil di-reset. Silakan unggah file master baru.")
            st.rerun()
            
        # METRICS GLOBAL
        col1, col2, col3 = st.columns(3)
        col1.metric("👥 Total Anggota Terdaftar", f"{len(df)} Orang")
        col2.metric("💰 Total Simpanan Terkumpul", f"Rp {df['TOTAL SIMPANAN'].sum():,.0f}")
        col3.metric("📉 Total Pinjaman Beredar", f"Rp {df['PINJAMAN'].sum():,.0f}")
        st.write("---")
        
        # PENCARIAN ANGGOTA
        cari_nama = st.text_input("🔍 Ketik Nama Anggota untuk Cek Saldo rill:")
        if cari_nama:
            df_tampil = df[df["NAMA"].astype(str).str.contains(cari_nama, case=False)]
            st.dataframe(df_tampil, use_container_width=True, hide_index=True)
        else:
            row_total = {df.columns[i]: "" for i in range(len(df.columns))}
            row_total["NAMA"] = "TOTAL KESELURUHAN"
            for k in ["SIMPANAN POKOK", "SIMPANAN WAJIB", "SIMPANAN SUKARELA", "JASA PINJAMAN", "PINJAMAN", "TOTAL SIMPANAN"]:
                if k in df.columns:
                    row_total[k] = df[k].sum()
            df_cetak = pd.concat([df, pd.DataFrame([row_total])], ignore_index=True)
            st.dataframe(df_cetak, use_container_width=True, hide_index=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_cetak.to_excel(writer, index=False, sheet_name='Laporan Terhitung')
            
            st.download_button(
                label="📥 Cetak Laporan / Unduh Excel Terhitung Otomatis",
                data=buffer.getvalue(),
                file_name="Laporan_Koperasi_Terhitung_Otomatis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("💡 Sistem belum memiliki database master yang tertanam. Silakan masuk ke menu '🚀 Inisialisasi Data Awal' terlebih dahulu.")

# =====================================================================
# MENU 2: SINKRONISASI BULANAN
# =====================================================================
elif menu_navigasi == "🔄 Sinkronisasi Iuran Bulanan":
    if df_master is not None:
        st.success("✅ Database Master Aktif Terdeteksi di Komputer.")
        update_file = st.file_uploader("📂 Unggah File Excel Potongan Bulan Ini:", type=["xlsx"])
        
        if update_file is not None:
            df_update = pd.read_excel(update_file)
            df_update.columns = [str(c).strip().upper() for c in df_update.columns]
            
            if st.button("🚀 Jalankan Sinkronisasi & Kunci ke Harddisk"):
                df_master["NAMA_KEY"] = df_master["NAMA"].apply(standarkan_nama)
                df_update["NAMA_KEY"] = df_update["NAMA"].apply(standarkan_nama)
                
                jumlah_success = 0
                for idx, row in df_update.iterrows():
                    nama_orang = row["NAMA_KEY"]
                    if nama_orang in df_master["NAMA_KEY"].values:
                        wajib_baru = pd.to_numeric(row.get("SIMPANAN WAJIB", 0), errors='coerce') or 0
                        sukarela_baru = pd.to_numeric(row.get("SIMPANAN SUKARELA", 0), errors='coerce') or 0
                        jasa_baru = pd.to_numeric(row.get("JASA PINJAMAN", 0), errors='coerce') or 0
                        pinjaman_baru = pd.to_numeric(row.get("PINJAMAN", 0), errors='coerce') or 0
                        
                        df_master.loc[df_master["NAMA_KEY"] == nama_orang, "SIMPANAN WAJIB"] += wajib_baru
                        df_master.loc[df_master["NAMA_KEY"] == nama_orang, "SIMPANAN SUKARELA"] += sukarela_baru
                        df_master.loc[df_master["NAMA_KEY"] == nama_orang, "JASA PINJAMAN"] += jasa_baru
                        df_master.loc[df_master["NAMA_KEY"] == nama_orang, "PINJAMAN"] = pinjaman_baru
                        jumlah_success += 1
                
                df_master = df_master.drop(columns=["NAMA_KEY"])
                df_master.to_excel(DATABASE_FISIK, index=False)
                st.success(f"🎉 Sukses Memperbarui! {jumlah_success} data anggota telah dikunci permanen.")
                st.rerun()
    else:
        st.warning("⚠️ Silakan tanamkan data master awal terlebih dahulu di menu '🚀 Inisialisasi Data Awal'.")