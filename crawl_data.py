import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import re

# URL của trang chứa thông tin sim
base_url = 'https://sim.vn/sim-so-dep-duoi-1368'

# List các dãy số không mong muốn
unwanted_numbers = ['89', '46', '64', '97', '79', '38', '83']

# Hàm kiểm tra nếu số điện thoại chứa dãy không mong muốn
def is_unwanted_number(phone_number):
    for unwanted in unwanted_numbers:
        if unwanted in phone_number:
            return True
    if phone_number[0] == '0' and '0' in phone_number[1:]:
        return True
    if re.search(r'(\d)\1{2,3}', phone_number):  # 3 hoặc 4 số giống nhau
        return True
    return False

# Hàm lấy tổng số trang
def get_total_pages(soup):
    pagination = soup.find('div', class_='pagination')
    if not pagination:
        return 1
    pages = pagination.find_all('a')
    
    last_page = 1
    for page in pages:
        if page.text.isdigit():
            last_page = max(last_page, int(page.text))
    
    return last_page

# Hàm lấy thông tin sim từ từng trang
def get_sim_data_from_page(page_num):
    page_url = f'{base_url}?page={page_num}'
    response = requests.get(page_url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Lấy thông tin sim từ trang
    sim_items = soup.find_all('a', class_='sim')
    
    page_data = []
    for item in sim_items:
        phone_number = item.get('href').split('/')[-1]  # Lấy số điện thoại từ href
        
        # Lọc bỏ số không mong muốn
        if is_unwanted_number(phone_number):
            continue
        
        price = item.find('div', class_='sim__price').text.strip()
        network_logo_src = item.find('img')['src']  
        network_name = network_logo_src.split('/')[-1].split('.')[0].capitalize()

        page_data.append({
            'Số': phone_number,
            'Giá': price,
            'Nhà mạng': network_name
        })
    
    return page_data

# Bước 1: Lấy nội dung trang đầu tiên để xác định số trang
response = requests.get(base_url)
soup = BeautifulSoup(response.content, 'html.parser')

# Bước 2: Lấy tổng số trang từ phần pagination
total_pages = get_total_pages(soup)

# Danh sách lưu dữ liệu sim
all_sim_data = []

# Bước 3: Lặp qua từng trang và lấy dữ liệu sim
for page in range(1, total_pages + 1):
    print(f"Đang lấy dữ liệu từ trang {page}/{total_pages}")
    page_data = get_sim_data_from_page(page)
    
    # Gộp dữ liệu vào danh sách chung
    all_sim_data.extend(page_data)

# Tạo DataFrame từ dữ liệu thu thập được
df = pd.DataFrame(all_sim_data)

# Sắp xếp dữ liệu theo 3 số đầu của số điện thoại
df['Số'] = df['Số'].astype(str)  # Đảm bảo rằng số điện thoại là kiểu chuỗi
df_sorted = df.sort_values(by='Số', key=lambda col: col.str[:3])  # Sắp xếp theo 3 số đầu

# Lấy thời gian hiện tại để thêm vào tên file
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
file_name = f'sim_filtered_{current_time}.xlsx'

# Xuất dữ liệu ra file Excel
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df_sorted.to_excel(writer, index=False)

    # Lấy đối tượng workbook và worksheet
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Tạo định dạng cỡ chữ 25
    format_text = workbook.add_format({'font_size': 35})

    # Thiết lập chiều rộng của tất cả các cột là 120 và áp dụng cỡ chữ 25 cho dữ liệu
    for col_num, col in enumerate(df_sorted.columns):
        worksheet.set_column(col_num, col_num, 50, format_text)

print(f"Dữ liệu đã được xuất ra file {file_name}")