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
import tarfile
from io import BytesIO
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def stream_data(data):
    for word in data.split(" "):
        yield word + " "
        time.sleep(0.03)

### TAB 1 FUNCTIONS
def convert_date_format(date_str):
    # Parse the date string
    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d')

    # Format the date in the desired format
    formatted_date = f"Ngày {date_obj.day:02} tháng {date_obj.month:02} năm {date_obj.year}"
    return formatted_date

@st.cache_data
def pxk_data_from_xml(file):
    data=[]
    ggia = 0
    tree = ET.parse(file)
    root = tree.getroot()
    shdon = int(root.find('.//TTChung/SHDon').text) 
    nmua = root.find('.//NMua/Ten').text 
    nmua_dc = root.find('.//NMua/DChi').text
    nban = (root.find('.//NBan/Ten').text).replace('HỘ KINH DOANH', '').strip().title()
    nban_mst = root.find('.//NBan/MST').text
    nban_dc = (root.find('.//NBan/DChi').text).replace(', Bà Rịa - Vũng Tàu', '').strip()
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

    return  shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, data

@st.cache_data
def display_pxk(shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, data):
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

    # all_data.append((shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, df))

    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Ngày-Tháng-Năm: {date}")
    st.text(f"Tên khách: {nmua}")


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

def pxk_excel(wb, shdon, nmua, nban, nban_dc, nban_mst, date, tbc, ggia, df):
    template_sheet = wb['Template']
    
    # Create a new sheet by copying the template sheet
    new_sheet_name = f"{shdon}"  # Ensure the sheet name is within 31 characters
    wb.copy_worksheet(template_sheet).title = new_sheet_name
    ws = wb[new_sheet_name]

    if 'Địa chỉ' in template_sheet['A2'].value:
        parts = template_sheet['A2'].value.split('\n', 1) 
        new_value = f"{parts[0].strip()} {nban_dc}\n{parts[1].strip()} {nban_mst}"
    ws['A2'] = new_value
    ws['A1'] = f"Hộ kinh doanh: {nban}"
    ws['A4'] = date
    ws['A8'] = f"Địa điểm xuất kho: {nban_dc}"
    ws['C41'] = shdon
    ws['C45'] = nban
    ws['F45'] = nban
    ws['C40'] = tbc
    ws['A6'] = f"Tên đơn vị mua hàng: {nmua}"

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

### TAB 2 FUNCTIONS
@st.cache_data
def display_ptt(shdon, nmua, date, tbc, ts):
    st.subheader(f"Số hóa đơn: {shdon}")
    st.text(f"Ngày-Tháng-Năm: {date}")
    st.text(f"Tên khách: {nmua}")
    st.text(f"Tổng tiền: {ts} đồng")
    st.text(f"Tổng iền bằng chữ: {tbc}")

def ptt_excel(wb, shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts):
    template_sheet = wb['Template']
    new_sheet_name = f"{shdon}"  
    wb.copy_worksheet(template_sheet).title = new_sheet_name
    ws = wb[new_sheet_name]

    if 'Địa chỉ' in template_sheet['A2'].value:
        parts = template_sheet['A2'].value.split('\n', 1) 
        new_value = f"{parts[0].strip()} {nban_dc}\n{parts[1].strip()} {nban_mst}"
    ws['A7'] = f"Địa chỉ: {nmua_dc}"
    ws['A2'] = new_value
    ws['A1'] = f"Hộ kinh doanh: {nban}"
    ws['A17'] = nban
    ws['C17'] = nban
    ws['B4'] = date
    ws['F14'] = date
    ws['C10'] = tbc
    ws['C24'] = tbc
    ws['B9'] = ts
    ws['C23'] = ts
    ws['D6'] = nmua
    ws['D12'] = shdon

### TAB 3 FUNCTIONS
def extract_number_vnpt(string):
    # Find the position of the last underscore and the '.zip' extension
    last_underscore_pos = string.rfind('_')
    dot_zip_pos = string.find('.zip')
    
    # Extract the number between the last underscore and '.zip'
    number = string[last_underscore_pos + 1:dot_zip_pos]
    return number

def extract_zipfile(zip_file, extract_to):
    extracted_files = []
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith(".xml"):
                extracted_files.append(file)
                zip_ref.extract(file, extract_to)
    return extracted_files

def download_zip(driver, action, wait, download_path):
    xml_files = []
    icons = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, "//a[@title='Xem chi tiết hóa đơn']"))
    )
    icons = icons[:len(icons)//2]
    for icon in icons:
        try:
            action.move_to_element(icon).perform()
            driver.execute_script("arguments[0].click();", icon)
            time.sleep(3)
            
            invoice_form = driver.find_element(By.XPATH, "//div[@class='modal-content']")

            download_button = invoice_form.find_element(By.XPATH, "//div[@id='taiXml']")
            driver.execute_script("arguments[0].click();", download_button)
            time.sleep(3)

            downloaded_file = wait_for_download(download_path)
            if downloaded_file:
                # Extract the number from the downloaded file (assuming the naming convention is the same)
                shd = extract_number_vnpt(os.path.basename(downloaded_file))
                # Extract the zip file contents
                extracted_files = extract_zipfile(downloaded_file, download_path)

                # Rename and store the extracted XML files in the list
                for file in extracted_files:
                    xml_file = shd + file[file.index('.xml'):]
                    xml_files.append((xml_file, file))
                    os.rename(os.path.join(download_path, file), os.path.join(download_path, xml_file))
            

            close_button = invoice_form.find_element(By.XPATH, "//button[@class='close']")
            driver.execute_script("arguments[0].click();", close_button)
            time.sleep(2)

        except StaleElementReferenceException:
            icon = driver.find_element(By.XPATH, "//a[@title='Xem chi tiết hóa đơn']")
            action.move_to_element(icon).perform()

    return xml_files
     
def wait_for_download(download_path, timeout=30):
    '''Wait for a file to be downloaded to the download path'''
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = os.listdir(download_path)
        if files:
            # Get the most recently downloaded file
            latest_file = max([os.path.join(download_path, f) for f in files], key=os.path.getctime)
            
            # Ensure the file is fully downloaded (not a .crdownload file)
            if not latest_file.endswith('.crdownload') and os.path.getsize(latest_file) > 0:
                return latest_file
        
        time.sleep(1)

    return None

def selenium_web_driver(download_path):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_path,  # Ensure this path matches your temp dir
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    # Point the browser to the correct location
    chrome_options.binary_location = "/usr/bin/chromium"

    # Use chromedriver installed by the system package manager
    driver = webdriver.Chrome(service=Service("/usr/bin/chromedriver"), options=chrome_options)
    action=ActionChains(driver,10)
    wait = WebDriverWait(driver, 10)
    
    return driver, action, wait

### TAB 4 FUNCTIONS
def extract_number_viettel(string):
    match = re.search(r'TAV(\d+)', string)
    if match:
        return match.group(1)
    else:
        return None

def download_XML(driver, action, wait, temp_folder):
    xml_files = []
    icons = wait.until(
        EC.presence_of_all_elements_located((By.XPATH, "//button[i[contains(@class, 'fa-info icon-info')]]"))
    )
    for icon in icons:
        try:
            st.write(icon)
            action.move_to_element(icon).perform()
            driver.execute_script("arguments[0].click();", icon)
            time.sleep(3)
            
            invoice_form = driver.find_element(By.XPATH, "//div[@class='modal-content']")
            download_button = invoice_form.find_element(By.XPATH, "//button[span[text()='Tải xml']]")
            driver.execute_script("arguments[0].click();", download_button)
            time.sleep(3)

            downloaded_file = wait_for_download(temp_folder)
            if downloaded_file:
                shd = extract_number_viettel(os.path.basename(downloaded_file))
                xml_file = shd + downloaded_file[downloaded_file.index('.xml'):]
                xml_files.append(xml_file)

            close_button = invoice_form.find_element(By.XPATH, "//button[@class='close']")
            driver.execute_script("arguments[0].click();", close_button)
            time.sleep(2)

        except StaleElementReferenceException:
            icons = wait.until(
                EC.presence_of_all_elements_located((By.XPATH, "//button[i[contains(@class, 'fa-info icon-info')]]"))
            )
            action.move_to_element(icon).perform()
    return xml_files
### Streamlit State FUNCTIONS
def create():
    st.session_state['create_success'] = True

def download():
    st.session_state['download_success'] = True

def download_ptt():
    st.session_state['download_success_ptt'] = True
    
### MAIN FUNCTION
def main():
    with st.sidebar:
        xml_files = st.file_uploader("Nhập XML files", accept_multiple_files=True, type='xml')
        if len(xml_files) == 0:
            st.warning("Vui lòng tải tệp XML của bạn")
        else:
            st.success(f'Bạn đã tải thành công {len(xml_files)} tệp')
        st.divider()
            
    tab1, tab2, tab3, tab4= st.tabs(['Phiếu xuất kho', 'Phiếu thu tiền', 'VNPT', 'Viettel'])

    tab1.title("Phiếu xuất kho")
    tab2.title("Phiếu thu tiền")
    tab3.title("Tải zip VNPT")
    tab4.title("Viettel")

    with tab1:
        if xml_files:
            all_data = []
            for uploaded_file in xml_files:
                
                shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, data= pxk_data_from_xml(uploaded_file)

                with st.container(border=True): 
                    df = display_pxk(shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, data)
                    all_data.append((shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts ,ggia, df))
                    
                ## TAB 2
                with tab2:
                    with st.container(border=True):
                        display_ptt(shdon, nmua, date, tbc, ts)

            print(len(all_data))

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
                    if st.button('Tạo phiếu xuất kho và thu tiền', type='primary', key='btn', on_click=create):
                        pass
                else:
                    with st.spinner("Đang tạo phiếu ..."):
                        time.sleep(1.5)
                        st.success("Tạo phiếu xuất kho và thu tiền thành công")

                        pxk_file_path = 'pxk.xlsx'
                        ptt_file_path = 'ptt.xlsx'

                        pxk_wb = load_workbook(pxk_file_path)
                        ptt_wb = load_workbook(ptt_file_path)

                        for shdon, nmua, _ , nban, nban_dc, nban_mst, date, tbc, ts, ggia, df in all_data:
                            pxk_excel(pxk_wb, shdon, nmua, nban, nban_dc, nban_mst, date, tbc, ggia, df)

                        for shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts, ggia, df in all_data:    
                            ptt_excel(ptt_wb, shdon, nmua, nmua_dc, nban, nban_dc, nban_mst, date, tbc, ts)

                        if len(pxk_wb.sheetnames) > 1 and len(ptt_wb.sheetnames) > 1:
                            pxk_wb.active = 1  
                            ptt_wb.active = 1  
                            # Remove the template sheet
                            pxk_wb.remove(pxk_wb['Template'])
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
    
    with tab3:
        user = st.radio(
            "Hộ kinh doanh:",
            ["Trần Minh Đạt", "Nguyễn Thị Thanh Thúy"]
        )

        date_start, date_end = st.columns(2)
        with date_start:
            start_date = st.date_input(
                "Ngày bắt đầu:", 
                format="DD/MM/YYYY", 
                key='vnpt_start').strftime("%d/%m/%Y")
        with date_end:
            end_date = st.date_input(
                "Ngày kết thúc:", 
                format="DD/MM/YYYY",
                key='vnpt_end').strftime("%d/%m/%Y")

        if st.button("Tải Zip tự động"):
            download_path = tempfile.mkdtemp()
            driver, action, wait = selenium_web_driver(download_path)            

            with st.status("Đang tải Zip tự động ...", expanded=True) as status:
                try:
                    driver.get('https://hkd.vnpt.vn/account/login')
                    driver.implicitly_wait(2)

                    wait.until(
                        EC.presence_of_element_located((By.CLASS_NAME, 'form-horizontal'))
                    )

                    username = driver.find_element(By.NAME, 'UserName')
                    password = driver.find_element(By.NAME, 'Password')

                    if user == "Trần Minh Đạt":
                        username.send_keys(os.getenv('username'))
                        password.send_keys(os.getenv('password'))
                        password.send_keys(Keys.RETURN)
                    else:
                        username.send_keys(os.getenv('username_2'))
                        password.send_keys(os.getenv('password_2'))
                        password.send_keys(Keys.RETURN)
                        
                    driver.implicitly_wait(5)

                    st.write_stream(stream_data((f"Đang đăng nhập tài khoản {user}...")))
                    time.sleep(2)

                    qlhd = wait.until(
                        EC.presence_of_element_located((By.XPATH,"//a[@href='/Thue/QuanLyHoaDon']"))
                    )
                    driver.execute_script("arguments[0].click();", qlhd)
                    st.write_stream(stream_data(("Đang vào mục Quản Lý Hóa Đơn ...")))
                    time.sleep(2)
                    
                    date_btn = driver.find_elements(By.CLASS_NAME, "dx-texteditor-input")
                    date_btn[0].clear()
                    date_btn[0].send_keys(start_date)
                    st.write_stream(stream_data((f"Đang nhập ngày bắt đầu: {start_date}")))
                    time.sleep(2)

                    date_btn[1].clear()
                    date_btn[1].send_keys(end_date)
                    st.write_stream(stream_data((f"Đang nhập ngày kết thúc: {end_date}")))
                    time.sleep(2)


                    search_btn = wait.until(
                        EC.presence_of_all_elements_located((By.CLASS_NAME, "dx-button-content"))
                    )
                    search_btn[3].click()
                    st.write_stream(stream_data(("Đang tìm hóa đơn ...")))
                    time.sleep(3)

                    page_size = driver.find_element(By.XPATH, "//div[@aria-label='Display 50 items on page']")
                    page_size.click()
                    st.write_stream(stream_data(("Chọn hiển thị 50 hóa đơn ...")))
                    time.sleep(2)
                    
                    final_xml_files = []
                    all_pages = driver.find_elements(By.XPATH, "//div[@class='dx-page-indexes']")
                    if all_pages:
                        st.write_stream(stream_data((f"Tổng số trang: {len(all_pages)}")))
                        for i, page in enumerate(all_pages):
                            st.write_stream(stream_data((f"Đang tải hóa đơn ở trang số {i + 1} ...")))
                            xml_files = download_zip(driver, action, wait, download_path)
                            final_xml_files.append(xml_files)
                            page.click()
                            time.sleep(3)
                    else:
                        st.write_stream(stream_data(("Không có trang nào được tìm thấy")))
                    
                    final_xml_files = [item for sublist in final_xml_files for item in sublist]
                    st.write_stream(stream_data((f"Tổng số hóa đơn: :red[{len(final_xml_files)}]")))
                    time.sleep(3)
                    
                except Exception as e:
                    st.error(f"Lỗi: {e}")
                finally:
                    if driver:
                        driver.quit()  # Close the driver if it was initialized
                    status.update(label="Tải thành công !!!", expanded=True)
                    
            # Define the tar archive path
            tar_path = os.path.join(download_path, 'extracted_xml_files.tar')
            
            # Create the tar archive with only XML files in the download path
            with tarfile.open(tar_path, 'w') as tar:
                # Filter XML files only
                for file in filter(lambda f: f.endswith('.xml'), os.listdir(download_path)):  
                    tar.add(os.path.join(download_path, file), arcname=file)

            # Provide download of the tar file
            with open(tar_path, 'rb') as f:
                if st.download_button("Tải thư mục XML", f, file_name="XML_files.tar", type="primary"):
                    downloading_message = 'Đang tải thư mục ...'
                    progress_bar = st.progress(0, text=downloading_message)
                    for percent_complete in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(percent_complete + 1, text=downloading_message)
                    time.sleep(1)
                    st.success("Đã tải thư mục XML thành công")

    with tab4:
        user = st.radio(
            "Hộ kinh doanh:",
            ["An Vinh"]
        )

        date_start, date_end = st.columns(2)
        with date_start:
            start_date = st.date_input(
                "Ngày bắt đầu:", 
                format="DD/MM/YYYY", 
                key='viettel_start').strftime("%d/%m/%Y")
        with date_end:
            end_date = st.date_input(
                "Ngày kết thúc:",  
                format="DD/MM/YYYY",
                key='viettel_end').strftime("%d/%m/%Y")

        if st.button("Tải XML tự động"):
            temp_folder = tempfile.mkdtemp()
            driver, action, wait = selenium_web_driver(temp_folder)            

            with st.status("Đang tải Zip tự động ...", expanded=True) as status:
                try:
                    driver.get('https://vinvoice.viettel.vn/account/login')
                    driver.implicitly_wait(2)

                    wait.until(
                        EC.presence_of_element_located((By.XPATH, '//form[@role="form"]'))
                    )

                    username = driver.find_element(By.ID, 'username')
                    password = driver.find_element(By.NAME, 'password')

                    username.send_keys(os.getenv('username_viettel'))
                    password.send_keys(os.getenv('password_viettel'))
                    password.send_keys(Keys.RETURN)

                    qlhd = wait.until(
                        EC.presence_of_element_located((By.XPATH,"//a[@href='/invoice-management/invoice']")))
                    driver.execute_script("arguments[0].click();", qlhd)
                    st.write_stream(stream_data(("Đang vào mục Quản Lý Hóa Đơn ...")))
                    time.sleep(2)

                    date_btn = driver.find_elements(By.XPATH, "//input[@formcontrolname='datePicker']")
                    date_btn[0].clear()
                    date_btn[0].send_keys(start_date)
                    st.write_stream(stream_data((f"Đang nhập ngày bắt đầu: {start_date}")))
                    time.sleep(2)

                    date_btn[1].clear()
                    date_btn[1].send_keys(end_date)
                    st.write_stream(stream_data((f"Đang nhập ngày kết thúc: {end_date}")))
                    time.sleep(2)

                    search_btn = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//button[span[text()='Tìm kiếm']]"))
                    )
                    search_btn.click()
                    st.write_stream(stream_data(("Đang tìm hóa đơn ...")))
                    time.sleep(3)


                    select_size = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//select[@name='pageSize']"))
                    )

                    # Use JavaScript to set the value
                    driver.execute_script("arguments[0].value = '50';", select_size)

                    # Optionally, dispatch a change event to simulate the user's interaction
                    # driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", select_size)

                    st.write_stream(stream_data(("Chọn hiển thị 50 hóa đơn ...")))
                    time.sleep(2)

                    final_xml_files = []
                    all_pages = driver.find_elements(By.XPATH, "//a[@class='page-link ng-star-inserted']")
                    if all_pages:
                        st.write_stream(stream_data((f"Tổng số trang: {len(all_pages)}")))
                        for i, page in enumerate(all_pages):
                            st.write_stream(stream_data((f"Đang tải hóa đơn ở trang số {i + 1} ...")))
                            xml_files = download_XML(driver, action, wait, temp_folder)
                            st.write_stream(xml_files)
                            final_xml_files.append(xml_files)
                            
                    final_xml_files = [item for sublist in final_xml_files for item in sublist]
                    st.write_stream(stream_data((f"Tổng số hóa đơn: :red[{len(final_xml_files)}]")))
                    time.sleep(3)
                except Exception as e:  
                    st.error(f"Lỗi: {e}")

if __name__ == "__main__":
    main()