import streamlit as st
import pdfplumber
import pandas as pd
import io
import datetime
import requests
import xlrd
from xlutils.copy import copy

# Page Config
st.set_page_config(page_title="Sami PDF to XLS", layout="centered")

st.title("🚛 PDF to XLS Converter (Template Based)")

# --- 1. User Inputs ---
direction_choice = st.radio("İşlem Türü:", ('Çıkış', 'Giriş'))
files = st.file_uploader("PDF Dosyalarını Seçin", type="pdf", accept_multiple_files=True)
belge_no_input = st.text_input("Beyanname No (Sadece Çıkış için):")

# Logic for inputs
if direction_choice == 'Çıkış':
    yon = 'Ç'
    belge_tur = 3
    belge_no = belge_no_input
else:
    yon = 'G'
    belge_tur = ''
    belge_no = '' 

# --- 2. Processing ---
if st.button("Dönüştür ve İndir"):
    if not files:
        st.error("Lütfen en az bir PDF dosyası seçin.")
    else:
        try:
            st.info("PDF'ler işleniyor...")
            
            all_pages = []

            # Process uploaded files
            for uploaded_file in files:
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        table = page.extract_table()
                        if table:
                            # Assume first row is header
                            df1 = pd.DataFrame(table[1:], columns=table[0])
                            all_pages.append(df1)

            if not all_pages:
                st.warning("PDF'lerden tablo okunamadı.")
                st.stop()

            # Combine dataframes
            all_data = pd.concat(all_pages)
            
            # --- Data Transformation ---
            # Extract Vehicle Plate
            df_output = pd.DataFrame(all_data['Araç Plaka'])

            # Extract Date/Time (Last Column)
            last_col_name = all_data.columns[-1]
            date_time_raw = all_data[last_col_name].str.split(expand=True)
            
            if date_time_raw.shape[1] >= 2:
                date_time = date_time_raw.iloc[:, 0:2]
                date_time.columns = ['date', 'time']
            else:
                date_time = pd.DataFrame({'date': date_time_raw[0], 'time': ''})

            # Fix Date Format
            date_time['date'] = date_time['date'].astype(str).apply(lambda x: x if '/' in x else x.replace('.', '/'))

            # Construct Final DataFrame Columns
            df_output.insert(0, 'YÖN', yon)
            df_output.insert(1, 'BELGE_TÜRÜ', belge_tur)
            df_output.insert(2, 'BELGE_NO', belge_no)
            df_output.insert(4, 'DORSE1', '')
            df_output.insert(5, 'DORSE2', '')
            df_output.insert(6, 'date', date_time['date'])
            df_output.insert(7, 'time', date_time['time'])
            
            # IMPORTANT: In your original code, you set index to YÖN.
            # xlwings writes the index. To mimic this manually, we don't drop the index,
            # but since 'YÖN' is already a column, we just ensure the order is correct.
            # The dataframe currently looks like: [YÖN, BELGE_TÜRÜ, BELGE_NO, Araç Plaka, DORSE1, DORSE2, date, time]
            # This matches the column structure implicitly. We extract values directly.

            # --- 3. Template Handling (.xls) ---
            st.info("Şablon indiriliyor ve dolduruluyor...")
            
            # A. Download the Template
            template_url = 'http://www.mavi.web.tr/ygms/Arac_Giris_Cikis_Aktarim_Sablon.xls'
            response = requests.get(template_url)
            response.raise_for_status()
            
            # B. Open with xlrd (formatting_info=True keeps the styles)
            # Note: xlrd only supports .xls, which is exactly what we want.
            rb = xlrd.open_workbook(file_contents=response.content, formatting_info=True)
            
            # C. Create a writable copy using xlutils
            wb = copy(rb)
            sheet = wb.get_sheet(0) # Get the first sheet

            # D. Write data into the sheet starting at Row 2 (Index 1)
            # We iterate over the rows and columns of the dataframe
            # Convert dataframe to list of lists for easy iteration
            data_values = df_output.values.tolist()
            
            start_row = 1 # Row 2 in Excel (0-based index)
            
            for r_idx, row_data in enumerate(data_values):
                for c_idx, cell_value in enumerate(row_data):
                    # Write data. Row = start + current_index
                    # We convert to string to ensure safety, or keep as is if int/float
                    sheet.write(start_row + r_idx, c_idx, cell_value)

            # --- 4. Save to Memory and Download ---
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            file_name = f"Sami-{datetime.datetime.today().strftime('%Y-%m-%d-%H-%M-%S')}.xls"
            
            st.success("Dosya Hazır!")
            st.download_button(
                label="📥 Oluşturulan .xls Dosyasını İndir",
                data=output_buffer,
                file_name=file_name,
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"Bir hata oluştu: {e}")