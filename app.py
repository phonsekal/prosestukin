import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide")

st.title("🚀 Penggabung Data Tukin ke Format ADK")
st.write("Upload 4 file sumber untuk digabungkan menjadi format ADK PEMBAYARAN.")

# Sidebar untuk konfigurasi statis
st.sidebar.header("Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="02")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

# Upload Multi File
uploaded_files = st.file_uploader("Pilih file Excel sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Proses & Gabungkan Data"):
        all_data = []
        
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                
                # Pembersihan data (Menghapus baris kosong/header tambahan jika ada)
                df = df.dropna(subset=['NIP', 'Nama'], how='all')
                
                # Mapping kolom secara dinamis berdasarkan data Anda[cite: 1, 2]
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df['NIP'].astype(str)
                df_clean['NAMA_PEGAWAI'] = df['Nama']
                
                # Mengambil nilai keuangan dengan opsi nama kolom berbeda
                df_clean['BRUTO'] = df.get('Nilai Bruto', df.get('Tunjangan', 0))
                df_clean['POTONGAN'] = df.get('Nilai Potongan', df.get('Potongan', 0))
                df_clean['BERSIH'] = df.get('Nilai Bersih', df.get('Bersih', 0))
                
                # Menambahkan kolom periode
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"Berhasil membaca: {file.name}")
            except Exception as e:
                st.error(f"Gagal memproses {file.name}: {e}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            # Tampilkan preview data
            st.subheader("Preview Hasil Gabungan")
            st.dataframe(final_df.head(10))
            
            # Tombol Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='ADK_GABUNGAN')
            
            st.download_button(
                label="📥 Download Hasil Gabungan (Excel)",
                data=output.getvalue(),
                file_name=f"ADK_PEMBAYARAN_GABUNGAN_{bulan}_{tahun}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.info("Catatan: Pastikan kolom di file Excel sumber minimal memiliki header 'NIP' dan 'Nama'.")