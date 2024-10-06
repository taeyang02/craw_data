import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import re

# URL của trang chứa thông tin sim
base_url = 'https://sim.vn/sim-so-dep-duoi-268'

# List các dãy số không mong muốn
unwanted_numbers = ['89', '46', '64', '97', '79', '38', '83']

# Hàm kiểm tra nếu số điện thoại chứa dãy không mong muốn
def is_unwanted_number(phone_number):
    # Kiểm tra các dãy không mong muốn
    for unwanted in unwanted_numbers:
        if unwanted in phone_number:
            return True
    
    # Kiểm tra nếu có nhiều hơn 1 số 0 ngoài số đầu tiên
    if phone_number[0] == '0' and '0' in phone_number[1:]:
        return True
    
    # Kiểm tra nếu có ba hoặc bốn số giống nhau đứng liền nhau
    if re.search(r'(\d)\1{2,3}', phone_number):  # 3 hoặc 4 số giống nhau
        return True

    return False

# Hàm lấy tất cả các số trang
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
        page_data.append({'Số': phone_number, 'Giá': price, 'Nhà mạng': network_name})
    return page_data

# Bước 1: Lấy nội dung trang đầu tiên để xác định số trang
response = requests.get(base_url)
soup = BeautifulSoup(response.content, 'html.parser')

# Bước 2: Lấy tổng số trang từ phần pagination
total_pages = get_total_pages(soup)

# Danh sách lưu dữ liệu sim
all_sim_data = {}

# Bước 3: Lặp qua từng trang và lấy dữ liệu sim
for page in range(1, total_pages + 1):
    print(f"Đang lấy dữ liệu từ trang {page}/{total_pages}")
    page_data = get_sim_data_from_page(page)
    
    # Lưu dữ liệu vào dictionary với khóa là số trang
    all_sim_data[f'Page {page}'] = page_data

# Tạo DataFrame từ dữ liệu thu thập được
max_length = max(len(data) for data in all_sim_data.values())
data_for_excel = {f'Page {i + 1}': [None] * max_length for i in range(total_pages)}

for i, (page, sim_data) in enumerate(all_sim_data.items()):
    for j, sim in enumerate(sim_data):
        data_for_excel[page][j] = f"Số: {sim['Số']}, Giá: {sim['Giá']}, Nhà mạng: {sim['Nhà mạng']}"

# Lấy thời gian hiện tại để thêm vào tên file
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
file_name = f'sim_viettel_filtered_{current_time}.xlsx'

# Chuyển đổi dữ liệu thành DataFrame và xuất ra file Excel
df = pd.DataFrame(data_for_excel)

# Sử dụng xlsxwriter để đặt chiều rộng của ô và cỡ chữ
with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False)

    # Lấy đối tượng workbook và worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Tạo định dạng cỡ chữ 25
    format_text = workbook.add_format({'font_size': 25})

    # Thiết lập chiều rộng của tất cả các cột là 90 và áp dụng cỡ chữ 25 cho dữ liệu
    for col_num, col in enumerate(df.columns):
        worksheet.set_column(col_num, col_num, 120, format_text)

print(f"Dữ liệu đã được xuất ra file {file_name}")
