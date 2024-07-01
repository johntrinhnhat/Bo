import os
import re
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

def extract_number(string):
    # Find the position of the last underscore and the '.zip' extension
    last_underscore_pos = string.rfind('_')
    dot_zip_pos = string.find('.zip')
    
    # Extract the number between the last underscore and '.zip'
    number = string[last_underscore_pos + 1:dot_zip_pos]
    return number

# Function to extract data from an XML file
def pxk_data_from_xml(file):
    data=[]
    ggia = 0
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = int(root.find('.//TTChung/SHDon').text) 
    tendvi = root.find('.//NMua/Ten').text 
    date = convert_date_format(root.find('.//NLap').text) 
    tbc = root.find('.//TgTTTBChu').text 
    ts = int(root.find('.//TgTTTBSo').text)
    for item in root.findall('.//HHDVu'):
        stt = int(item.find('STT').text) if item.find('STT') is not None else ""
        thhdv = item.find('THHDVu').text if item.find('THHDVu') is not None else ""
        dvtinh = item.find('DVTinh').text if item.find('DVTinh') is not None else ""
        sluong = float(item.find('SLuong').text.replace(',', '')) if item.find('SLuong') is not None and item.find('SLuong').text else 0
        dgia = int(item.find('DGia').text.replace(',', '').replace('.', '')) if item.find('DGia') is not None and item.find('DGia').text else 0
        thtien = int(item.find('ThTien').text.replace(',', '').replace('.', '')) if item.find('ThTien') is not None and item.find('ThTien').text else 0
        if "Đã giảm" in thhdv:
            ggia_number = re.search(r'(\d{1,3}(?:\.\d{3})*)\s+đồng', thhdv)
            ggia = int(ggia_number.group(1).replace('.', ''))
            data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon, ggia])
        else:
            data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon])

    return  shdon, tendvi, date, tbc, ts, ggia, data

def display_pxk(shdon, tendvi, date, tbc, ts, ggia, data, all_data):
    columns = ['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Số hóa đơn']
    if ggia:
        columns.append('Giảm giá')

    df = pd.DataFrame(data, columns=columns)
    df = df[df['Thành tiền'] != 0]
    df.loc[:, 'Tên hàng hóa, dịch vụ'] = df['Tên hàng hóa, dịch vụ'].str.capitalize()
    df.loc[:, 'Đơn vị tính'] = df['Đơn vị tính'].str.capitalize()

    df = df[[
        'STT', 
        'Tên hàng hóa, dịch vụ', 
        'Đơn vị tính', 
        'Số lượng', 
        'Đơn giá', 
        'Thành tiền'
    ]]

    all_data.append((shdon, tendvi, date, tbc, ts, ggia, df))

    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Ngày-Tháng-Năm: {date}")
    st.text(f"Tên khách: {tendvi}")


    st.dataframe(
        df.style.format({
        'Số lượng': lambda x: f'{x:.1f}'.replace('.', ','),
        'Đơn giá': lambda x: f"{x:,.0f}".replace(',', '.'),
        'Thành tiền': lambda x: f"{x:,.0f}".replace(',', '.')}),
        width=800
    )
    
    st.text(f"Giảm giá: {ggia} đồng")
    st.text(f"Tổng thành tiền: {tbc}")

    return df

def display_ptt(shdon, tendvi, date, tbc, ts):
    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Ngày-Tháng-Năm: {date}")
    st.text(f"Tên khách: {tendvi}")
    st.text(f"Tổng tiền: {ts} đồng")
    st.text(f"Tổng iền bằng chữ: {tbc}")

def ptt_excel(wb, shdon, tendvi, date, tbc, ts):
    template_sheet = wb['Template']
    # Create a new sheet by copying the template sheet
    new_sheet_name = f"{shdon}"  # Ensure the sheet name is within 31 characters
    wb.copy_worksheet(template_sheet).title = new_sheet_name
    ws = wb[new_sheet_name]
    print(type(shdon))
    ws['B4'] = date
    ws['F14'] = date
    ws['D6'] = tendvi
    ws['D12'] = shdon
    ws['C10'] = tbc
    ws['C24'] = tbc
    ws['B9'] = ts
    ws['C23'] = ts


def pxk_excel(wb, shdon, tendvi, date, tbc, ggia, df):
    template_sheet = wb['Template']
    
    # Create a new sheet by copying the template sheet
    new_sheet_name = f"{shdon}"  # Ensure the sheet name is within 31 characters
    wb.copy_worksheet(template_sheet).title = new_sheet_name
    ws = wb[new_sheet_name]

    ws['A4'] = date
    ws['C41'] = shdon
    ws['C40'] = tbc
    ws['A6'] = f"Tên đơn vị mua hàng: {tendvi}"

    if ggia:
        ws['D39'] = ggia 

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

def download_ptt():
    st.session_state['download_success_ptt'] = True
    

def main():
    all_data = []
    with st.sidebar:
        xml_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')
        if len(xml_files) == 0:
            st.warning("Vui lòng tải tệp XML của bạn")
        else:
            st.success(f'Bạn đã tải thành công {len(xml_files)} tệp')
        st.divider()
            

    tab1, tab2, tab3 = st.tabs(['Phiếu xuất kho', 'Phiếu thu tiền', 'Zip'])

    tab1.title("Phiếu xuất kho")
    tab2.title("Phiếu thu tiền")
    ## TAB 1
    with tab1:
        if xml_files:
            for uploaded_file in xml_files:
                shdon, tendvi,date, tbc, ts, ggia, data= pxk_data_from_xml(uploaded_file)
                with st.container(border=True): 
                    display_pxk(shdon, tendvi, date, tbc, ts, ggia, data, all_data)
                ## TAB 2
                with tab2:
                    with st.container(border=True):
                        display_ptt(shdon, tendvi, date, tbc, ts)

            if 'create_success' not in st.session_state:
                st.session_state['create_success'] = False
            if 'download_success' not in st.session_state:
                st.session_state['download_success'] = False
            if 'download_success_ptt' not in st.session_state:
                st.session_state['download_success_ptt'] = False

            with st.sidebar:
                if st.session_state['download_success']:
                    downloading_message = 'Đang tải phiếu xuất kho ...'
                    progress_bar = st.progress(0, text=downloading_message)
                    for percent_complete in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(percent_complete + 1, text=downloading_message)
                    time.sleep(1)
                    st.success("Đã tải phiếu xuất kho thành công")
                    progress_bar.empty()
                    st.session_state['download_success'] = False
                    
                if st.session_state['download_success_ptt']:
                    downloading_message = 'Đang tải phiếu thu tiền ...'
                    progress_bar = st.progress(0, text=downloading_message)
                    for percent_complete in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(percent_complete + 1, text=downloading_message)
                    time.sleep(1)
                    st.success("Đã tải phiếu thu tiền thành công")
                    progress_bar.empty()
                    st.session_state['download_success_ptt'] = False

                if not st.session_state['create_success']:
                    if st.button('Tạo phiếu xuất kho và thu tiền', type='primary', key='xk_btn', on_click=create):
                        pass
                else:
                    with st.spinner("Đang tạo phiếu ..."):
                        time.sleep(1.5)
                        st.success("Tạo phiếu xuất kho và thu tiền thành công")

                        pxk_file_path = 'pxk.xlsx'
                        ptt_file_path = 'ptt.xlsx'

                        pxk_wb = load_workbook(pxk_file_path)
                        ptt_wb = load_workbook(ptt_file_path)


                        for shdon, tendvi, date, tbc, ts, ggia, df in all_data:
                            pxk_excel(pxk_wb, shdon, tendvi,date, tbc, ggia, df)
                            ptt_excel(ptt_wb, shdon, tendvi,date, tbc, ts)


                        if len(pxk_wb.sheetnames) > 1:
                            pxk_wb.active = 1  # Set any other sheet as the active one
                            # Remove the template sheet
                            pxk_wb.remove(pxk_wb['Template'])

                        if len(ptt_wb.sheetnames) > 1:
                            ptt_wb.active = 1  # Set any other sheet as the active one
                            # Remove the template sheet
                            ptt_wb.remove(ptt_wb['Template'])
                    
                        # Save the workbook to a bytes buffer
                        pxk_buffer = BytesIO()
                        pxk_wb.save(pxk_buffer)
                        pxk_buffer.seek(0)

                        # Save the workbook to a bytes buffer
                        ptt_buffer = BytesIO()
                        ptt_wb.save(ptt_buffer)
                        ptt_buffer.seek(0)

            
                        st.download_button(
                            on_click=download,
                            type="primary",
                            label="Tải phiếu xuất kho",
                            data=pxk_buffer,
                            file_name="PHIEU XUAT KHO QUY ( TI - EM ).xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="pxk"
                        )

                        st.download_button(
                            on_click=download_ptt,
                            type="primary",
                            label="Tải phiếu thu tiền",
                            data=ptt_buffer,
                            file_name="PHIEU THU TIEN QUY ( TI-EM ).xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="ptt"
                        )
                            
    
    #     if xml_files:
    #         print(type(xml_files))
    #         for uploaded_file in xml_files:
    #             print(uploaded_file)
    #             timing, shd, tdv, ts, tc = ptt_data_from_xml(uploaded_file)
                # print(timing, shd, tdv, ts, tc)
            #     with st.container(border=True): 
            #         display_ptt(shdon, tendvi, date, tbc, ggia, data, all_data)


    # with tab3:
    #     st.title("Zip => XML")

    #     # File uploader
    #     uploaded_files = st.file_uploader("Nhập Zip files", type="zip", accept_multiple_files=True)

    #     if not uploaded_files:
    #         st.warning("Vui lòng tải tệp Zip của bạn")
    #     else:
    #         st.success(f"Bạn đã tải thành công {len(uploaded_files)} zip")
    #         all_xml_files = []
    #         temp_folder = tempfile.mkdtemp()

    #         for zip_files in uploaded_files:
    #             shd = extract_number(zip_files.name)
    #             original_in_zip_file = extract_zipfile(zip_files, temp_folder)
    #             for file in original_in_zip_file:
    #                 xml_file = shd + file[file.index('.xml'):]
    #                 all_xml_files.append((xml_file, file))
    #                 os.rename(os.path.join(temp_folder, file), os.path.join(temp_folder, xml_file))

    #         shutil.make_archive(temp_folder, 'tar', temp_folder)
    #         with open(temp_folder + '.tar', 'rb') as f:
    #             if st.download_button("Tải thư mục XML", f, file_name="extracted_xml_files.tar"):
    #                 downloading_message = 'Đang tải thư mục ...'
    #                 progress_bar = st.progress(0, text=downloading_message)
    #                 for percent_complete in range(100):
    #                     time.sleep(0.01)
    #                     progress_bar.progress(percent_complete + 1, text=downloading_message)
    #                 time.sleep(1)
    #                 st.success("Đã tải thư mục XML thành công")
    
    
if __name__ == "__main__":
    main()