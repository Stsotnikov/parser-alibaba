import bs4
import xlsxwriter
import requests
import time  # Импортируем модуль для добавления задержки

# URL страницы для парсинга (общая часть)
main_url = 'https://bwintool.en.alibaba.com/'

# Заголовки для запроса (иногда требуется для обхода блокировок)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
}

# Функция для получения и парсинга страницы
def get_soup(url):
    res = requests.get(url, headers=headers)
    return bs4.BeautifulSoup(res.text, 'html.parser')

# Функция для парсинга товаров с одной страницы
def parse_products(page_url, data):
    products_page = get_soup(page_url)
    products = products_page.findAll('div', class_='icbu-product-card')

    # Парсим каждый товар на странице
    for product in products:
        # Извлекаем URL товара
        url_tag = product.find('a', class_='product-image')
        product_url = url_tag['href'].strip() if url_tag else None

        if product_url:
            full_product_url = 'https:' + product_url + '?language=ru_RU'  # Добавляем параметр языка
            print(f'Добавляем ссылку: {full_product_url}')

            # Добавляем только URL в категорию
            data.append([full_product_url])

        # Добавляем таймаут после парсинга каждого товара
        time.sleep(1)  # Пауза в 1 секунду между парсингом товаров

# Массивы для данных каждой категории
categories_data = {
    "End Mill": [['URL']],
    "Inch Size End Mill": [['URL']],
    "Carbide Insert": [['URL']],
    "Turning Bar Tool": [['URL']],
    "Face Milling Tool": [['URL']],
    "Hole Drilling Tool": [['URL']],
    "Thread Tapping Tool": [['URL']],
    "Tool Holder": [['URL']],
    "Chuck": [['URL']],
    "Brand Carbide Insert": [['URL']]
}

# Цикл для категории "End Mill" 1-16
for page_num in range(1, 16):
    page_url = f'{main_url}productgrouplist-801458118-{page_num}/End_mill.html'
    print(f'Парсим категорию End Mill: {page_url}')
    parse_products(page_url, categories_data["End Mill"])
    time.sleep(1)  # Добавляем паузу в 3 секунды между запросами категорий

# Цикл для категории "Inch Size End Mill"
for page_num in range(1, 2):
    page_url = f'{main_url}productgrouplist-943142891/Inch_Size_End_Mill.html'
    print(f'Парсим категорию Inch Size End Mill: {page_url}')
    parse_products(page_url, categories_data["Inch Size End Mill"])
    time.sleep(1)

# Цикл для категории "Carbide Insert" 
for page_num in range(1, 19):
    page_url = f'{main_url}productgrouplist-925017726-{page_num}/Carbide_Insert.html'
    print(f'Парсим категорию Carbide Insert: {page_url}')
    parse_products(page_url, categories_data["Carbide Insert"])
    time.sleep(5)

# Цикл для категории "Turning Bar Tool"
for page_num in range(1, 7):
    page_url = f'{main_url}productgrouplist-801799662-{page_num}/Turning_Bar_Tool.html'
    print(f'Парсим категорию Turning Bar Tool: {page_url}')
    parse_products(page_url, categories_data["Turning Bar Tool"])
    time.sleep(1)

# Цикл для категории "Face Milling Tool"
for page_num in range(1, 4):
    page_url = f'{main_url}productgrouplist-815719367-{page_num}/Face_Milling_Tool.html'
    print(f'Парсим категорию Face Milling Tool: {page_url}')
    parse_products(page_url, categories_data["Face Milling Tool"])
    time.sleep(1)

# Цикл для категории "Hole Drilling Tool"
for page_num in range(1, 6):
    page_url = f'{main_url}productgrouplist-811420827-{page_num}/Hole_Drilling_Tool.html'
    print(f'Парсим категорию Hole Drilling Tool: {page_url}')
    parse_products(page_url, categories_data["Hole Drilling Tool"])
    time.sleep(5)

# Цикл для категории "Thread Tapping Tool"
for page_num in range(1, 5):
    page_url = f'{main_url}productgrouplist-811267315-{page_num}/Thread_Tapping_Tool.html'
    print(f'Парсим категорию Thread Tapping Tool: {page_url}')
    parse_products(page_url, categories_data["Thread Tapping Tool"])
    time.sleep(1)

# Цикл для категории "Tool Holder"
for page_num in range(1, 2):
    page_url = f'{main_url}productgrouplist-813195154/Tool_Holder.html'
    print(f'Парсим категорию Tool Holder: {page_url}')
    parse_products(page_url, categories_data["Tool Holder"])
    time.sleep(1)

# Цикл для категории "Chuck"
for page_num in range(1, 2):
    page_url = f'{main_url}productgrouplist-931814882/Chuck.html'
    print(f'Парсим категорию Chuck: {page_url}')
    parse_products(page_url, categories_data["Chuck"])
    time.sleep(1)

# Цикл для категории "Brand Carbide Insert"
for page_num in range(1, 9):
    page_url = f'{main_url}productgrouplist-801473331-{page_num}/Brand_Carbide_Insert.html'
    print(f'Парсим категорию Brand Carbide Insert: {page_url}')
    parse_products(page_url, categories_data["Brand Carbide Insert"])
    time.sleep(1)

# Сохраняем данные в Excel
with xlsxwriter.Workbook('FaceMillingTool2.xlsx') as workbook:
   # Для каждой категории создаем отдельный лист
    for category, data in categories_data.items():
        worksheet = workbook.add_worksheet(category)

        # Записываем данные категории в соответствующий лист
        for row_num, row_data in enumerate(data):
            worksheet.write_row(row_num, 0, row_data)

print("Парсинг завершен, данные сохранены в 'categories_data_urls.xlsx'.")