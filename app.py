import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide", page_icon="💰")

st.title("💰 Penggabung Data Tukin ke Format ADK")
st.markdown("Versi Final: Perbaikan menyeluruh pada logika deteksi kolom dan perhitungan angka.")

# Sidebar
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

def clean_currency(value):
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).replace('Rp', '').replace(' ', '')
        if ',' in text and '.' in text:
            text = text.replace('.', '').replace(',', '.')
        elif ',' in text:
            text = text.replace(',', '.')
        clean_str = re.sub(r'[^\d.]', '', text)
        return float(clean_str) if clean_str else 0.0
    except:
        return 0.0

uploaded_files = st.file_uploader("Pilih file Excel sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Gabungkan Sekarang"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Cari baris header yang mengandung 'NIP'
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                found_header = False
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        found_header = True
                        break
                
                if not found_header:
                    st.error(f"❌ Tidak menemukan kolom NIP di file: {file.name}")
                    continue

                # 2. Baca ulang file dari baris header[cite: 1]
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Tangani Merge Cells pada Header
                # Menggunakan .ffill() untuk Pandas versi terbaru
                df.columns = pd.Series(df.columns).ffill().str.strip()

                # 4. Cari kolom NIP dan Nama secara eksplisit
                col_nip = None
                col_nama = None
                for c in df.columns:
                    c_str = str(c).upper()
                    if 'NIP' in c_str and col_nip is None:
                        col_nip = c
                    if 'NAMA' in c_str and col_nama is None:
                        col_nama = c

                if col_nip is None or col_nama is None:
                    st.error(f"❌ Kolom NIP atau Nama tidak ditemukan di: {file.name}")
                    continue

                # 5. Ekstraksi Data
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # Penanganan Kolom Keuangan (Bruto, Potongan, Bersih)[cite: 2]
                # Loop untuk setiap target kolom yang kita inginkan
                for target, keywords in [('BRUTO', 'BRUTO'), ('POTONGAN', 'POTONGAN'), ('BERSIH', 'BERSIH')]:
                    # Cari semua kolom yang mengandung kata kunci tersebut
                    matched_cols = [c for c in df.columns if keywords in str(c).upper()]
                    
                    if len(matched_cols) > 0:
                        # Inisialisasi kolom dengan angka nol
                        temp_series = pd.Series(0.0, index=df.index)
                        for m_col in matched_cols:
                            # Bersihkan dan jumlahkan setiap kolom yang cocok
                            temp_series += df[m_col].apply(clean_currency)
                        df_clean[target] = temp_series
                    else:
                        df_clean[target] = 0.0

                # 6. Validasi Data: Ambil hanya yang memiliki NIP valid (>= 9 digit)[cite: 1]
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9].copy()
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil memproses: {file.name}")
                
            except Exception as e:
                st.error(f"❌ Error pada {file.name}: {str(e)}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.divider()
            st.subheader("📊 Preview Data Gabungan")
            st.dataframe(final_df, use_container_width=True)
            
            # Tampilkan statistik akhir
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.2f}")
            c3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.2f}")

            # Tombol Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='ADK_GABUNGAN')
            
            st.divider()
            st.download_button(
                label="📥 Download Hasil Gabungan (Excel)",
                data=output.getvalue(),
                file_name=f"ADK_GABUNGAN_{bulan}_{tahun}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
