import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter

# Основной URL
main_url = 'https://bwintool.en.alibaba.com/'

# Инициализация драйвера
driver = webdriver.Firefox()

# Массив для данных
data = [['Название', 'URL', 'Картинка1', 'Картинка2', 'Картинка3', 'Картинка4', 'Картинка5', 'Картинка6', 'Диаметр',
         'Длина', 'Тип', 'Основные характеристики', 'Другие характеристики', 'Упаковка и доставка', 'Описание товара']]

# Список конкретных страниц для парсинга
pages_to_scrape = [
    'https://www.alibaba.com/product-detail/ZCC-CT-any-cut-cutting-tool_60840643160.html?language=ru_RU',
    'https://www.alibaba.com/product-detail/ZCC-CT-CNMG120404-EF-YBG205-Tungsten_60870120047.html?language=ru_RU'
]

# Функция для парсинга товара на странице товара
def parse_product_page(product_url):
    driver.get(product_url)
    time.sleep(2)  # Даем странице время на загрузку

    # Извлекаем название товара
    try:
        title = driver.find_element(By.TAG_NAME, 'h1').text
    except Exception:
        title = 'No Title'

    # Картинки
    images = [extract_image(i) for i in range(1, 7)]

    # Диаметры
    diameters = [extract_diameter_by_xpath(i) for i in range(1, 7)]

    # Длина
    lengths = [extract_length_by_xpath(i) for i in range(1, 7)]

    # Тип
    try:
        type_tag = driver.find_element(By.CSS_SELECTOR,
                                       'div.module_sku div.sku-info div:nth-child(6) a span')
        type = type_tag.text
    except Exception:
        type = 'No Title'

    # Атрибуты
    key_attributes = extract_attributes(2)
    other_attributes = extract_attributes(4)
    packaging_and_delivery = extract_attributes(6)

    # Описание товара
    try:
        detail_decorate_root = driver.find_element(By.ID, 'detail_decorate_root').get_attribute('outerHTML')
    except Exception:
        detail_decorate_root = '-'

    return title, *images, *diameters, *lengths, type, key_attributes, other_attributes, packaging_and_delivery, detail_decorate_root

# Функция для извлечения изображения
def extract_image(image):
    try:
        image_tag = driver.find_element(By.XPATH, f'//*[@id="container"]/div[1]/div[1]/div[4]/div/div/div[1]/div/div/div/div[{image}]/div')
        style_attr = image_tag.get_attribute('style')
        image_url = style_attr.split('url(')[1].split(')')[0].replace('//', 'https://').replace('"', '')
    except Exception:
        image_url = ''
    return image_url

# Вспомогательные функции для диаметра и длины
def extract_diameter_by_xpath(diameter):
    try:
        diameter_tag = driver.find_element(By.XPATH, f'//*[@id="container"]/div[1]/div[2]/div/div/div[6]/div/div[2]/div[1]/a[{diameter}]/span')
        return diameter_tag.text
    except Exception:
        return ''

def extract_length_by_xpath(length):
    try:
        length_tag = driver.find_element(By.XPATH, f'//*[@id="container"]/div[1]/div[2]/div/div/div[6]/div/div[2]/div[2]/a[{length}]/span')
        return length_tag.text
    except Exception:
        return ''

# Вспомогательная функция для извлечения атрибутов
def extract_attributes(nth_child):
    try:
        attributes_tag = driver.find_element(By.CSS_SELECTOR, f'#container > div.layout-body > div.layout-left > div.module_attribute > div > div > div:nth-child({nth_child})')
        return attributes_tag.get_attribute('outerHTML')
    except Exception:
        return '-'

# Цикл по страницам товаров из списка
for page_url in pages_to_scrape:
    print(f'Парсим товар: {page_url}')
    title, image1_url, image2_url, image3_url, image4_url, image5_url, image6_url, diameter1, diameter2, diameter3, diameter4, diameter5, diameter6, length1, length2, length3, length4, length5, length6, type, key_attributes, other_attributes, packaging_and_delivery, detail_decorate_root = parse_product_page(page_url)

    # Добавляем данные в таблицу
    data.append(
        [title, page_url, image1_url, image2_url, image3_url, image4_url, image5_url, image6_url,
         f'{diameter1}  {diameter2}  {diameter3} {diameter4}  {diameter5}  {diameter6}',
         f'{length1}  {length2}  {length3} {length4}  {length5}  {length6}', type, key_attributes, other_attributes,
         packaging_and_delivery, detail_decorate_root]
    )

# Сохраняем данные в Excel
with xlsxwriter.Workbook('result.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    # Записываем данные в Excel
    for row_num, info in enumerate(data):
        worksheet.write_row(row_num, 0, info)

print("Парсинг завершен, данные сохранены в 'result.xlsx'.")

# Закрываем драйвер
driver.quit()