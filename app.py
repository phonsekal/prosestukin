import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide", page_icon="💰")

st.title("💰 Penggabung Data Tukin ke Format ADK")
st.markdown("Versi Perbaikan: Penanganan angka murni dan format mata uang campuran.")

# Sidebar
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

# FUNGSI PERBAIKAN: Lebih cerdas dalam membedakan angka dan teks
def clean_currency(value):
    if pd.isna(value) or value == "":
        return 0.0
    
    # Jika sudah berupa angka (int/float), langsung kembalikan nilainya
    if isinstance(value, (int, float)):
        return float(value)
    
    # Jika berupa string, bersihkan karakter non-angka kecuali titik desimal
    try:
        # Hapus Rp, spasi, dan titik ribuan (titik yang diikuti 3 angka)
        text = str(value).replace('Rp', '').replace(' ', '')
        
        # Logika Indonesia: jika ada titik dan koma (misal 1.000,00), ubah ke format standar (1000.00)
        if ',' in text and '.' in text:
            text = text.replace('.', '').replace(',', '.')
        elif ',' in text: # Jika hanya koma (misal 1000,00)
            text = text.replace(',', '.')
            
        # Hapus semua karakter kecuali angka dan titik
        clean_str = re.sub(r'[^\d.]', '', text)
        
        return float(clean_str) if clean_str else 0.0
    except:
        return 0.0

uploaded_files = st.file_uploader("Pilih file-file Excel sumber", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Proses & Gabungkan Sekarang"):
        all_data = []
        
        for file in uploaded_files:
            try:
                # 1. Cari baris header NIP[cite: 1]
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        break
                
                # 2. Baca ulang file
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Tangani Merge Cells Header
                df.columns = pd.Series(df.columns).ffill().str.strip()

                # 4. Cari kolom NIP dan Nama
                col_nip = next((c for c in df.columns if 'NIP' in str(c).upper()), None)
                col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

                if not col_nip or not col_nama:
                    st.error(f"❌ Kolom NIP/Nama tidak ditemukan di: {file.name}")
                    continue

                # 5. Ekstraksi Data
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # Mengambil kolom keuangan (Bruto, Potongan, Bersih)
                # Menggunakan filter agar lebih fleksibel mencari nama kolom
                for target, keywords in [('BRUTO', 'Bruto'), ('POTONGAN', 'Potongan'), ('BERSIH', 'Bersih')]:
                    found_col = df.filter(like=keywords).columns
                    if not found_col.empty:
                        df_clean[target] = df[found_col[0]]
                    else:
                        df_clean[target] = 0.0

                # 6. Pembersihan Nilai Keuangan
                for col in ['BRUTO', 'POTONGAN', 'BERSIH']:
                    df_clean[col] = df_clean[col].apply(clean_currency)
                
                # Validasi NIP (minimal 9 digit) untuk membuang baris sampah/total[cite: 1]
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9]
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil: {file.name}")
                
            except Exception as e:
                st.error(f"❌ Error {file.name}: {str(e)}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            st.divider()
            st.subheader("📊 Preview Data Gabungan")
            st.dataframe(final_df, use_container_width=True)
            
            # Statistik
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.2f}")
            c3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.2f}")

            # Download
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
