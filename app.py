import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import os
import datetime
import openpyxl
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

# Function to extract data from an XML file
def extract_data_from_xml(file):
    data=[]
    ggia = None  # Initialize ggia with a default value
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = root.find('.//TTChung/SHDon').text if root.find('.//TTChung/SHDon') is not None else ''
    tendvi = root.find('.//NMua/Ten').text if root.find('.//NMua/Ten') is not None else ''
    date = root.find('.//NLap').text if root.find('.//NLap') is not None else ''
    tbc = root.find('.//TgTTTBChu').text if root.find('.//TgTTTBChu') is not None else ''
    
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
    if ggia is not None:
        columns.append('Giảm giá')

    df = pd.DataFrame(data, columns=columns)

    
    all_data.append((shdon, tendvi, date, tbc, ggia, df))

    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Năm-Tháng-Ngày: {date}")
    st.text(f"Tên khách: {tendvi}")

    df = df[['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền']]
    df = df[df['Đơn vị tính'].notna()]
    # df['Đơn giá'] = df['Đơn giá'].astype(int).apply(format_number)
    # df['Thành tiền'] = df['Thành tiền'].astype(int).apply(format_number)
    df['Đơn giá'] = df['Đơn giá'].astype(int)
    df['Thành tiền'] = df['Thành tiền'].astype(int)

    st.dataframe(df.style.hide(axis="index"), width=800)
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

    

def main():
    all_data = []
    st.title('XML to Excel Converter')
    xml_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')
    st.header(f'Tệp đã tải lên: {len(xml_files)}')

    with st.container(border=True):

        if xml_files:
            for uploaded_file in xml_files:
                shdon, tendvi,date, tbc, ggia, data = extract_data_from_xml(uploaded_file)
                display_invoice(shdon, tendvi, date, tbc, ggia, data, all_data)
            if st.button('Generate Excel File', type="primary"):
                excel_file_path = 'excel.xlsx'
                wb = load_workbook(excel_file_path)

                for shdon, tendvi, date, tbc, ggia, df in all_data:
                    update_excel(wb, shdon, tendvi,date, tbc, ggia, df)
            # st.success("Excel file updated successfully!")
                if len(wb.sheetnames) > 1:
                    wb.active = 1  # Set any other sheet as the active one
                    # Remove the template sheet
                    wb.remove(wb['Template'])
            
                # Save the workbook to a bytes buffer
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                # Provide download button
                st.download_button(
                    type="secondary",
                    label="Download Excel File",
                    data=buffer,
                    file_name="PHIEU XUAT KHO QUY (0x-2024).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_unique"
            )
                st.success("Generate file thành công")
                
                
                

if __name__ == "__main__":
    main()