import streamlit as st
import pandas as pd
import io
import re

# Konfigurasi Halaman
st.set_page_config(page_title="Pengolah Tukin ADK", layout="wide", page_icon="💰")

st.title("💰 Penggabung Data Tukin ke Format ADK")
st.markdown("Versi Stabil: Perbaikan deteksi kolom keuangan dan penanganan format angka.")

# Sidebar
st.sidebar.header("⚙️ Konfigurasi Data")
satker = st.sidebar.text_input("Kode Satker", value="693266")
bulan = st.sidebar.text_input("Bulan (MM)", value="04")
tahun = st.sidebar.text_input("Tahun (YYYY)", value="2026")

def clean_currency(value):
    """Mengubah berbagai format teks mata uang menjadi angka float murni."""
    if pd.isna(value) or str(value).strip() == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        # Menghapus Rp dan spasi
        text = str(value).replace('Rp', '').replace(' ', '')
        # Jika format Indonesia (1.234.567,89)
        if ',' in text and '.' in text:
            text = text.replace('.', '').replace(',', '.')
        elif ',' in text: # Jika hanya koma (1234,56)
            text = text.replace(',', '.')
        
        # Ambil hanya angka dan titik desimal
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
                # 1. Deteksi baris header NIP[cite: 1]
                df_raw = pd.read_excel(file, header=None)
                header_row = 0
                for i, row in df_raw.iterrows():
                    if row.astype(str).str.contains('NIP', case=False, na=False).any():
                        header_row = i
                        break
                
                # 2. Baca ulang file dari baris header tersebut[cite: 1, 2]
                df = pd.read_excel(file, skiprows=header_row)
                
                # 3. Tangani Merge Cells Header
                df.columns = pd.Series(df.columns).ffill().str.strip()

                # 4. Cari kolom NIP dan Nama secara eksplisit
                col_nip = next((c for c in df.columns if 'NIP' in str(c).upper()), None)
                col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

                if col_nip is None or col_nama is None:
                    st.error(f"❌ Kolom NIP atau Nama tidak ditemukan di: {file.name}")
                    continue

                # 5. Buat DataFrame baru untuk menampung hasil
                df_clean = pd.DataFrame()
                df_clean['KODE_SATKER'] = [satker] * len(df)
                df_clean['NIP'] = df[col_nip].astype(str).str.replace(r'\D', '', regex=True)
                df_clean['NAMA_PEGAWAI'] = df[col_nama]
                
                # 6. Proses Kolom Keuangan (Bruto, Potongan, Bersih)
                # Logika ini menjumlahkan otomatis jika ada kolom ganda (seperti Potongan Pph + Potongan Absen)
                for target, keywords in [('BRUTO', 'BRUTO'), ('POTONGAN', 'POTONGAN'), ('BERSIH', 'BERSIH')]:
                    matched_cols = [c for c in df.columns if keywords in str(c).upper()]
                    
                    temp_series = pd.Series(0.0, index=df.index)
                    if len(matched_cols) > 0:
                        for m_col in matched_cols:
                            temp_series += df[m_col].apply(clean_currency)
                    
                    df_clean[target] = temp_series

                # 7. Filter: Hanya ambil baris yang NIP-nya minimal 9 digit (Buang baris Total/Judul)[cite: 1]
                df_clean = df_clean[df_clean['NIP'].str.len() >= 9].copy()
                
                df_clean['BULAN'] = bulan
                df_clean['TAHUN'] = tahun
                
                all_data.append(df_clean)
                st.success(f"✅ Berhasil: {file.name}")
                
            except Exception as e:
                st.error(f"❌ Error pada {file.name}: {str(e)}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.divider()
            st.subheader("📊 Preview Data Gabungan")
            st.dataframe(final_df, use_container_width=True)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Pegawai", f"{len(final_df)} orang")
            c2.metric("Total Bruto", f"Rp {final_df['BRUTO'].sum():,.2f}")
            c3.metric("Total Bersih", f"Rp {final_df['BERSIH'].sum():,.2f}")

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
