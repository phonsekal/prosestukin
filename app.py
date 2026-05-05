import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Master Pembayaran Tukin", layout="wide", page_icon="📑")

st.title("📑 Generator Master Pembayaran Tukin")
st.markdown("Aplikasi ini menggabungkan data Tukin menjadi satu file **MASTER PEMBAYARAN.xlsx** yang bersih.")

# Sidebar untuk konfigurasi statis
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

def smart_numeric_cleaner(value):
    """Fungsi tangguh untuk memastikan semua input menjadi angka float murni."""
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    
    # Jika sudah berupa angka murni (int/float), langsung ambil
    if isinstance(value, (int, float)):
        return float(value)
    
    try:
        # Jika berupa string, bersihkan karakter non-angka
        text = str(value).replace('Rp', '').replace(' ', '')
        
        # Penanganan format angka Indonesia (titik ribuan, koma desimal)
        if ',' in text and '.' in text:
            text = text.replace('.', '').replace(',', '.')
        elif ',' in text:
            text = text.replace(',', '.')
            
        # Ambil hanya digit dan titik desimal
        clean_str = re.sub(r'[^\d.]', '', text)
        return float(clean_str) if clean_str else 0.0
    except:
        return 0.0

# Upload Multi File
uploaded_files = st.file_uploader("Upload semua file Tukin sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Buat Master Pembayaran"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Cari baris header yang berisi 'NIP'
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        break
                
                # 2. Baca ulang file dari baris header tersebut[cite: 1]
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Tangani Merge Cells pada Header (Penting untuk versi Pandas baru)[cite: 2]
                df.columns = pd.Series(df.columns).ffill().str.strip()

                # 4. Cari kolom NIP dan Nama secara spesifik
                col_nip = next((c for c in df.columns if 'NIP' in str(c).upper()), None)
                col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

                if col_nip is None or col_nama is None:
                    st.error(f"❌ Kolom NIP atau Nama tidak ditemukan di file: {file.name}")
                    continue

                # 5. Inisialisasi DataFrame bersih untuk satu file tersebut
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # 6. Pembersihan Nilai Keuangan (Bruto, Potongan, Bersih)[cite: 2]
                # Menangani potensi adanya beberapa kolom potongan dalam satu file
                for target, keywords in [('BRUTO', 'BRUTO'), ('POTONGAN', 'POTONGAN'), ('BERSIH', 'BERSIH')]:
                    matched_cols = [c for c in df.columns if keywords in str(c).upper()]
                    
                    temp_series = pd.Series(0.0, index=df.index)
                    if len(matched_cols) > 0:
                        for m_col in matched_cols:
                            # Terapkan pembersihan cerdas agar angka .0 tidak hilang
                            temp_series += df[m_col].apply(smart_numeric_cleaner)
                    
                    df_clean[target] = temp_series

                # 7. Validasi: Hanya ambil yang memiliki NIP valid (>= 9 digit)[cite: 1]
                # Menghapus otomatis baris judul, baris kosong, atau baris TOTAL di bawah
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9].copy()
                
                # Pastikan tipe data kolom keuangan adalah float murni sebelum penggabungan[cite: 2]
                for col in ['BRUTO', 'POTONGAN', 'BERSIH']:
                    df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0.0)
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil memproses: {file.name}")
                
            except Exception as e:
                st.error(f"❌ Gagal memproses {file.name}: {str(e)}")

        # 8. Penggabungan Akhir menjadi File Master
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.divider()
            st.subheader("📊 Pratinjau Master Pembayaran")
            st.dataframe(final_df, use_container_width=True)
            
            # Statistik untuk verifikasi cepat (Sudah aman dari TypeError)
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Nilai Bruto", f"Rp {final_df['BRUTO'].sum():,.2f}")
            c3.metric("Total Nilai Bersih", f"Rp {final_df['BERSIH'].sum():,.2f}")

            # Menyiapkan file Excel untuk diunduh
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='MASTER_PEMBAYARAN')
            
            st.divider()
            st.download_button(
                label="📥 Download MASTER PEMBAYARAN.xlsx",
                data=output.getvalue(),
                file_name="MASTER PEMBAYARAN.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Tidak ada data yang berhasil digabungkan. Periksa kembali file yang Anda unggah.")
