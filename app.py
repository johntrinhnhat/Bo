import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import os

# Function to extract data from an XML file
def extract_data_from_xml(file):
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = root.find('.//TTChung/SHDon').text if root.find('.//TTChung/SHDon') is not None else ''
    data=[]
    for item in root.findall('.//HHDVu'):
        stt = item.find('STT').text if item.find('STT') is not None else ''
        thhdv = item.find('THHDVu').text if item.find('THHDVu') is not None else ''
        dvtinh = item.find('DVTinh').text if item.find('DVTinh') is not None else ''
        sluong = item.find('SLuong').text if item.find('SLuong') is not None else ''
        dgia = item.find('DGia').text.replace(',', '.') if item.find('DGia') is not None else ''
        thtien = item.find('ThTien').text.replace(',', '.') if item.find('ThTien') is not None else ''
        data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon])
    return shdon, data
def main():

    st.title('XML to Excel Converter')
    uploaded_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')

    if uploaded_files:
        st.header(f'Tổng số file: {len(uploaded_files)}')
        for i, uploaded_file in enumerate(uploaded_files, start=1):
            shdon, file_data = extract_data_from_xml(uploaded_file)        
            df = pd.DataFrame(file_data, columns=['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Số hóa đơn'])
            st.subheader(f"Số hóa đơn: {shdon}")
            df = df[['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền']]
            st.dataframe(df.style.hide(axis="index"))

    # print(df)
    







if __name__ == "__main__":
    main()