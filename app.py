import io
import os
import shutil
import tempfile
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import datetime
import time
import zipfile
from io import BytesIO
from openpyxl import load_workbook

def format_number(num): 
    return f"{num:,}".replace(',', '.')

def convert_date_format(date_str):
    # Parse the date string
    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')

    # Format the date in the desired format
    formatted_date = f"Ngày {date_obj.day:02} tháng {date_obj.month:02} năm {date_obj.year}"
    return formatted_date

def extract_zipfile(zip_file, extract_to):
    extracted_files = []
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith(".xml"):
                extracted_files.append(file)
                zip_ref.extract(file, extract_to)
    return extracted_files


def zip_folder(folder_path, output_zip):
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file),
                           os.path.relpath(os.path.join(root, file),
                           os.path.join(folder_path, '..')))

def extract_number(string):
    # Find the position of the last underscore and the '.zip' extension
    last_underscore_pos = string.rfind('_')
    dot_zip_pos = string.find('.zip')
    
    # Extract the number between the last underscore and '.zip'
    number = string[last_underscore_pos + 1:dot_zip_pos]
    return number

# Function to extract data from an XML file
def extract_data_from_xml(file):
    data=[]
    ggia = 0  # Initialize ggia with a default value
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = int(root.find('.//TTChung/SHDon').text) 
    tendvi = root.find('.//NMua/Ten').text 
    date = root.find('.//NLap').text 
    tbc = root.find('.//TgTTTBChu').text 
    
    for item in root.findall('.//HHDVu'):
        stt = item.find('STT').text if item.find('STT') is not None else ''
        thhdv = item.find('THHDVu').text if item.find('THHDVu') is not None else ''
        dvtinh = item.find('DVTinh').text if item.find('DVTinh') is not None else ''
        sluong = item.find('SLuong').text if item.find('SLuong') is not None else ''
        dgia = item.find('DGia').text if item.find('DGia') is not None else ''
        thtien = item.find('ThTien').text if item.find('ThTien') is not None else ''

        try:
            dgia = int(dgia.replace(',', '').replace('.', '')) if dgia else 0
            thtien = int(thtien.replace(',', '').replace('.', '')) if thtien else 0
            sluong = int(sluong if sluong else 0)
            stt = int(stt if stt else 0)
            
        except ValueError:
            dgia = 0
            thtien = 0
        
        if 'Đã giảm' not in thhdv:  # Check if Sluong is not empty or None
            data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon])
        else:
            ggia = thhdv[8:14].replace('.', '')
            ggia = int(ggia)
            data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon, ggia])

        
    return  shdon, tendvi, date, tbc, ggia, data

def display_invoice(shdon, tendvi, date, tbc, ggia, data, all_data):
    columns = ['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Số hóa đơn']
    if ggia != 0:
        columns.append('Giảm giá')

    df = pd.DataFrame(data, columns=columns)
    
    all_data.append((shdon, tendvi, date, tbc, ggia, df))

    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Năm-Tháng-Ngày: {date}")
    st.text(f"Tên khách: {tendvi}")
    df.loc[:, 'Tên hàng hóa, dịch vụ'] = df['Tên hàng hóa, dịch vụ'].str.capitalize()
    df.loc[:, 'Đơn vị tính'] = df['Đơn vị tính'].str.capitalize()
    

    df = df[['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền']]
    df = df[df['Đơn vị tính'].notna()]
    df['Đơn giá'] = df['Đơn giá'].astype(int)
    df['Thành tiền'] = df['Thành tiền'].astype(int)
    # print(df['Tên hàng hóa, dịch vụ'])
    st.dataframe(df.style.hide(axis="index"), width=800)
    st.text(f"Giảm giá: {ggia} đồng")
    st.text(f"Tổng thành tiền: {tbc}")

    return df

def update_excel(wb, shdon, tendvi, date, tbc, ggia, df):
    template_sheet = wb['Template']
    
    # Create a new sheet by copying the template sheet
    new_sheet_name = f"Hóa đơn_{shdon}"  # Ensure the sheet name is within 31 characters
    wb.copy_worksheet(template_sheet).title = new_sheet_name
    ws = wb[new_sheet_name]

    date_formatted = convert_date_format(date)
    

    ws['A4'] = date_formatted
    ws['C38'] = shdon
    ws['C37'] = tbc
    ws['A6'] = f"Tên đơn vị mua hàng: {tendvi}"

    if ggia:
        ws['D36'] = ggia 

    start_row = 11
    start_column = 1  # Column A
    for i, row_data in enumerate(df.itertuples(index=False), start=start_row):
        if row_data[2]:
            ws.cell(row=i, column=start_column + 0, value=row_data[0])  # STT
            ws.cell(row=i, column=start_column + 1, value=row_data[1])  # Tên hàng hóa, dịch vụ
            ws.cell(row=i, column=start_column + 2, value=row_data[2])  # Đơn vị tính
            ws.cell(row=i, column=start_column + 3, value=row_data[3])  # Số lượng
            ws.cell(row=i, column=start_column + 4, value=row_data[4])  # Đơn giá
            ws.cell(row=i, column=start_column + 5, value=row_data[5])  # Thành tiền


def create():
    st.session_state['create_success'] = True

def download():
    st.session_state['download_success'] = True

def main():
    all_data = []
    with st.sidebar:
        xml_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')
        if len(xml_files) == 0:
            st.warning("Vui lòng tải tệp XML của bạn")
        else:
            st.success(f'Bạn đã tải thành công {len(xml_files)} tệp')
        st.divider()
            

    tab1, tab2 = st.tabs(['Phiếu xuất kho', 'Zip'])
    with tab1:
        if xml_files:
            for uploaded_file in xml_files:
                shdon, tendvi,date, tbc, ggia, data = extract_data_from_xml(uploaded_file)
                with st.container(border=True): 
                    display_invoice(shdon, tendvi, date, tbc, ggia, data, all_data)

            if 'create_success' not in st.session_state:
                st.session_state['create_success'] = False
            if 'download_success' not in st.session_state:
                st.session_state['download_success'] = False
            with st.sidebar:
                if st.session_state['download_success']:
                    downloading_message = 'Đang tải phiếu ...'
                    progress_bar = st.progress(0, text=downloading_message)
                    for percent_complete in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(percent_complete + 1, text=downloading_message)
                    time.sleep(1)
                    st.success("Đã tải phiếu xuất kho thành công")
                    progress_bar.empty()
                    st.session_state['download_success'] = False

                        
                else:
                    if not st.session_state['create_success']:
                        if st.button('Tạo phiếu xuất kho', type='primary', disabled=False, key='xk_btn', on_click=create):
                            pass
                    else:
                        with st.spinner("Đang tạo phiếu ..."):
                            time.sleep(1.5)
                            st.success("Tạo phiếu xuất kho thành công")

                            excel_file_path = 'excel.xlsx'
                            wb = load_workbook(excel_file_path)

                            for shdon, tendvi, date, tbc, ggia, df in all_data:
                                update_excel(wb, shdon, tendvi,date, tbc, ggia, df)

                            if len(wb.sheetnames) > 1:
                                wb.active = 1  # Set any other sheet as the active one
                                # Remove the template sheet
                                wb.remove(wb['Template'])
                        
                            # Save the workbook to a bytes buffer
                            buffer = BytesIO()
                            wb.save(buffer)
                            buffer.seek(0)

                
                            st.download_button(
                                on_click=download,
                                type="primary",
                                label="Tải phiếu xuất kho",
                                data=buffer,
                                file_name="PHIEU XUAT KHO QUY (0x-2024).xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_excel_unique"
                        )
    with tab2:
        st.title("zip")

        # File uploader
        uploaded_files = st.file_uploader("Choose a ZIP file", type="zip", accept_multiple_files=True)

        if not uploaded_files:
            st.warning("Vui lòng tải tệp Zip của bạn")
        else:
            st.success(f"Bạn đã tải thành công {len(uploaded_files)} zip")
            all_xml_files = []
            temp_folder = tempfile.mkdtemp()

            for zip_files in uploaded_files:
                shd = extract_number(zip_files.name)
                original_in_zip_file = extract_zipfile(zip_files, temp_folder)
                for file in original_in_zip_file:
                    xml_file = shd + file[file.index('.xml'):]
                    all_xml_files.append((xml_file, file))
                    os.rename(os.path.join(temp_folder, file), os.path.join(temp_folder, xml_file))

            shutil.make_archive(temp_folder, 'tar', temp_folder)
            with open(temp_folder + '.tar', 'rb') as f:
                if st.download_button("Tải thư mục XML", f, file_name="extracted_xml_files.tar"):
                    downloading_message = 'Đang tải thư mục ...'
                    progress_bar = st.progress(0, text=downloading_message)
                    for percent_complete in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(percent_complete + 1, text=downloading_message)
                    time.sleep(1)
                    st.success("Đã tải thư mục XML thành công")

    
if __name__ == "__main__":
    main()