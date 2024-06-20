import io
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import os

def format_number(num):
    return f"{num:,}".replace(',', '.')

# Function to extract data from an XML file
def extract_data_from_xml(file):
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = root.find('.//TTChung/SHDon').text if root.find('.//TTChung/SHDon') is not None else ''
    tendvi = root.find('.//Ten').text if root.find('.//Ten') is not None else ''
    date = root.find('.//NLap').text if root.find('.//NLap') is not None else ''
    tbc = root.find('.//TgTTTBChu').text if root.find('.//TgTTTBChu') is not None else ''
    data=[]
    print(tendvi, date, tbc)
    for item in root.findall('.//HHDVu'):
        stt = item.find('STT').text if item.find('STT') is not None else ''
        thhdv = item.find('THHDVu').text if item.find('THHDVu') is not None else ''
        dvtinh = item.find('DVTinh').text if item.find('DVTinh') is not None else ''
        sluong = item.find('SLuong').text if item.find('SLuong') is not None else ''
        dgia = item.find('DGia').text if item.find('DGia') is not None else ''
        thtien = item.find('ThTien').text if item.find('ThTien') is not None else ''
        data.append([stt, thhdv, dvtinh, sluong, dgia, thtien, shdon])
    return shdon, tendvi, date, tbc, data
def main():
    all_data = []
    st.title('XML to Excel Converter')
    uploaded_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')
    st.header(f'Tệp đã tải lên: {len(uploaded_files)}')
    with st.container(border=True):
        if uploaded_files:
            for i, uploaded_file in enumerate(uploaded_files, start=1):
                shdon, tendvi,date, tbc, file_data = extract_data_from_xml(uploaded_file)        
                df = pd.DataFrame(file_data, columns=['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Số hóa đơn'])
                all_data.append(df)
                
                st.subheader(f"Số hóa đơn: {shdon}")
                st.text(f"Năm-Tháng-Ngày: {date}")
                st.text(f"Tên đơn vị: {tendvi}")
                df = df[['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền']]
                df = df[df['Số lượng'].notna()]
                df['Đơn giá'] = df['Đơn giá'].astype(int).apply(format_number)
                df['Thành tiền'] = df['Thành tiền'].astype(int).apply(format_number)
                st.dataframe(df.style.hide(axis="index"), width=800)
                st.text(f"Tổng thành tiền: {tbc}")

                
        
            # # Combine all dataframes
            # combined_df = pd.concat(all_data, ignore_index=True)
            
            # # Display combined data
            # # Initialize session state for toggling visibility
            # if combined_df is not None:
            #     with st.popover("Show data combined"):
            #         st.dataframe(combined_df[['STT', 'Tên hàng hóa, dịch vụ', 'Đơn vị tính', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Số hóa đơn']].style.hide(axis="index"))
                
            #     # Option to download combined data as Excel
            #     st.header("Download Combined Data")
            #     buffer = io.BytesIO()
            #     with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            #         combined_df.to_excel(writer, index=False, sheet_name='Combined Data')
            #     st.download_button(
            #         type='primary',
            #         label="Download Excel",
            #         data=buffer,
            #         file_name="combined_data.xlsx",
            #         mime="application/vnd.ms-excel"
            #     ) 



if __name__ == "__main__":
    main()