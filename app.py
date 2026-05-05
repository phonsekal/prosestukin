import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide", page_icon="💰")

# Tampilan Header
st.title("💰 Penggabung Data Tukin ke Format ADK")
st.markdown("""
Aplikasi ini otomatis mendeteksi baris header, menangani *merge cells*, dan membersihkan format mata uang 
dari file rekap Tukin Anda.
""")

# Sidebar untuk konfigurasi statis
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

st.sidebar.markdown("---")
st.sidebar.info("""
**Instruksi:**
1. Upload file Excel sumber (bisa lebih dari satu).
2. Pastikan file memiliki kolom bertajuk **NIP** dan **Nama**.
3. Sistem akan otomatis membuang baris judul atau total yang tidak valid.
""")

# Fungsi Pembersihan Mata Uang (Pencegah TypeError)
def clean_currency(value):
    if pd.isna(value) or value == "":
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    # Hapus Rp, titik ribuan, koma, dan spasi
    clean_str = re.sub(r'[^\d.]', '', str(value).replace(',', ''))
    try:
        return float(clean_str)
    except:
        return 0

# Upload Multi File
uploaded_files = st.file_uploader("Pilih file-file Excel sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Gabungkan Sekarang"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Baca mentah untuk mencari baris header NIP
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        break
                
                # 2. Baca ulang file dengan baris header yang ditemukan
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Menangani Merge Cells pada Header (Horizontal)
                df.columns = pd.Series(df.columns).fillna(method='ffill').str.strip()

                # 4. Cari kolom NIP dan Nama secara fleksibel
                col_nip = next((c for c in df.columns if 'NIP' in str(c).upper()), None)
                col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

                if not col_nip or not col_nama:
                    st.error(f"❌ Kolom NIP/Nama tidak ditemukan di file: {file.name}")
                    continue

                # 5. Ekstraksi Data Utama
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # Mengambil nilai keuangan dengan filter nama kolom yang mirip
                df_clean['BRUTO'] = df.filter(like='Bruto').iloc[:, 0] if not df.filter(like='Bruto').empty else df.get('Tunjangan', 0)
                df_clean['POTONGAN'] = df.filter(like='Potongan').iloc[:, 0] if not df.filter(like='Potongan').empty else 0
                df_clean['BERSIH'] = df.filter(like='Bersih').iloc[:, 0] if not df.filter(like='Bersih').empty else 0

                # 6. Pembersihan Data (Angka & Validasi NIP)
                for col in ['BRUTO', 'POTONGAN', 'BERSIH']:
                    df_clean[col] = df_clean[col].apply(clean_currency)
                
                # Hanya ambil yang NIP-nya valid (minimal 9 digit angka)
                # Ini akan membuang baris "Total" atau baris kosong secara otomatis
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9]
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil memproses: {file.name} ({len(df_clean)} baris)")
                
            except Exception as e:
                st.error(f"❌ Gagal memproses {file.name}: {str(e)}")

        # 7. Penggabungan & Visualisasi
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.divider()
            st.subheader("📊 Preview Data Gabungan")
            st.dataframe(final_df, use_container_width=True)
            
            # Statistik (Sekarang aman dari TypeError)
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.0f}")
            c3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.0f}")

            # Persiapkan file download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='ADK_GABUNGAN')
            
            st.divider()
            st.download_button(
                label="📥 Download Hasil Gabungan (Excel)",
                data=output.getvalue(),
                file_name=f"ADK_TUKIN_GABUNGAN_{bulan}_{tahun}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Tidak ada data valid yang bisa digabungkan.")
else:
    st.info("Silakan pilih file Excel untuk memulai.")
