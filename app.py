import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide", page_icon="💰")

# Tampilan Header
st.title("💰 Penggabung Data Tukin ke Format ADK")
st.markdown("""
Aplikasi ini menggabungkan file rekap Tukin yang berantakan (banyak *merge cells* atau baris kosong) 
menjadi satu file bersih untuk keperluan ADK.
""")

# Sidebar untuk konfigurasi statis
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

st.sidebar.markdown("---")
st.sidebar.info("""
**Tips:**
1. Anda bisa upload banyak file sekaligus.
2. Script akan mencari kolom 'NIP' dan 'Nama' secara otomatis.
3. Baris 'Total' di bawah akan dibuang otomatis.
""")

# Upload Multi File
uploaded_files = st.file_uploader("Pilih file-file Excel sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Gabungkan Sekarang"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Baca mentah untuk mencari baris header
                # Kita cari sel yang mengandung kata 'NIP'
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        break
                
                # 2. Baca ulang file dengan baris header yang tepat
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Menangani Merge Cells pada Header (Horizontal)
                # Mengisi nama kolom yang kosong (Unnamed) dengan nama kolom sebelumnya
                new_cols = []
                last_col = "Unknown"
                for col in df.columns:
                    if "Unnamed" not in str(col):
                        last_col = str(col).strip()
                    new_cols.append(last_col)
                df.columns = new_cols

                # 4. Cari kolom NIP dan Nama secara fleksibel (Case Insensitive)
                col_nip = next((c for c in df.columns if 'NIP' in str(c).upper()), None)
                col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

                if not col_nip or not col_nama:
                    st.error(f"❌ File '{file.name}' tidak memiliki kolom NIP atau Nama.")
                    continue

                # 5. Ekstraksi Data Utama
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(" ", "").str.strip()
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # Mengambil nilai keuangan (Mencari kolom yang mengandung kata kunci)
                # Ambil iloc[:, 0] untuk antisipasi jika ada dua kolom dengan nama mirip
                df_clean['BRUTO'] = df.filter(like='Bruto').iloc[:, 0] if not df.filter(like='Bruto').empty else df.get('Tunjangan', 0)
                df_clean['POTONGAN'] = df.filter(like='Potongan').iloc[:, 0] if not df.filter(like='Potongan').empty else 0
                df_clean['BERSIH'] = df.filter(like='Bersih').iloc[:, 0] if not df.filter(like='Bersih').empty else 0

                # 6. Pembersihan Akhir (Data Cleaning)
                # - Isi NaN dengan 0 untuk angka
                df_clean[['BRUTO', 'POTONGAN', 'BERSIH']] = df_clean[['BRUTO', 'POTONGAN', 'BERSIH']].fillna(0)
                
                # - Hanya ambil baris yang NIP-nya berisi angka (minimal 9 digit)
                # Ini akan otomatis membuang baris judul, baris kosong, atau baris "TOTAL"
                df_clean = df_clean[df_clean['NIP'].str.contains(r'\d{9,}', na=False)]
                
                # - Tambahkan periode
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil memproses: {file.name} ({len(df_clean)} Pegawai)")
                
            except Exception as e:
                st.error(f"❌ Gagal memproses {file.name}: {str(e)}")

        # 7. Penggabungan Akhir
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.divider()
            st.subheader("📊 Preview Data Gabungan")
            st.dataframe(final_df, use_container_width=True)
            
            # Statistik Singkat
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Pegawai", f"{len(final_df)} orang")
            col2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.0f}")
            col3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.0f}")

            # Persiapkan file download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='ADK_GABUNGAN')
            
            st.divider()
            st.download_button(
                label="📥 Download File Hasil Gabungan (Excel)",
                data=output.getvalue(),
                file_name=f"HASIL_GABUNGAN_TUKIN_{bulan}_{tahun}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Tidak ada data yang berhasil digabungkan. Periksa kembali format file Anda.")

else:
    st.info("Silakan pilih satu atau beberapa file Excel di atas untuk memulai.")
