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


# Функция для парсинга товара на странице товара
def parse_product_page(product_url):
    driver.get(product_url)
    time.sleep(1)  # Даем странице время на загрузку

    # Извлекаем название товара
    try:
        title = driver.find_element(By.TAG_NAME, 'h1').text
    except Exception:
        title = 'No Title'

    # Картинки
    image1_url = extract_image(1)
    image2_url = extract_image(2)
    image3_url = extract_image(3)
    image4_url = extract_image(4)
    image5_url = extract_image(5)
    image6_url = extract_image(6)

    # Диаметры
    diameter1 = extract_diameter_by_xpath(1)
    diameter2 = extract_diameter_by_xpath(2)
    diameter3 = extract_diameter_by_xpath(3)
    diameter4 = extract_diameter_by_xpath(4)
    diameter5 = extract_diameter_by_xpath(5)
    diameter6 = extract_diameter_by_xpath(6)

    # Длина
    length1 = extract_length_by_xpath(1)
    length2 = extract_length_by_xpath(2)
    length3 = extract_length_by_xpath(3)
    length4 = extract_length_by_xpath(4)
    length5 = extract_length_by_xpath(5)
    length6 = extract_length_by_xpath(6)

    # Тип
    try:
        type_tag = driver.find_element(By.CSS_SELECTOR,
                                       '#container > div.layout-body > div.layout-right > div > div > div.module_sku > div > div.sku-info > div:nth-child(6) > a > span')
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

    return title, image1_url, image2_url, image3_url, image4_url, image5_url, image6_url, diameter1, diameter2, diameter3, diameter4, diameter5, diameter6, length1, length2, length3, length4, length5, length6, type, key_attributes, other_attributes, packaging_and_delivery, detail_decorate_root


# Функция для извлечения изображения
def extract_image(image):
    try:
        image_tag = driver.find_element(By.XPATH,f'//*[@id="container"]/div[1]/div[1]/div[4]/div/div/div[1]/div/div/div/div[{image}]/d')
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
        length_tag = driver.find_element(By.XPATH,f'//*[@id="container"]/div[1]/div[2]/div/div/div[6]/div/div[2]/div[2]/a[{length}]/span')
        return length_tag.text
    except Exception:
        return ''


# Вспомогательная функция для извлечения атрибутов
def extract_attributes(nth_child):
    try:
        attributes_tag = driver.find_element(By.CSS_SELECTOR,
                                             f'#container > div.layout-body > div.layout-left > div.module_attribute > div > div > div:nth-child({nth_child})')
        return attributes_tag.get_attribute('outerHTML')
    except Exception:
        return '-'


# Функция для парсинга товаров с одной страницы
def parse_products(page_url):
    driver.get(page_url)
    WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME, 'icbu-product-card')))  # Ждем загрузки товаров

    products = driver.find_elements(By.CLASS_NAME, 'icbu-product-card')  # Находим все товары

    # Парсим каждый товар на странице
    for product in products:
        try:
            product_url = product.find_element(By.CSS_SELECTOR, 'a.product-image').get_attribute('href')
            if product_url:
                full_product_url = product_url + '?language=ru_RU'
                print(f'Парсим товар: {full_product_url}')
                title, image1_url, image2_url, image3_url, image4_url, image5_url, image6_url, diameter1, diameter2, diameter3, diameter4, diameter5, diameter6, length1, length2, length3, length4, length5, length6, type, key_attributes, other_attributes, packaging_and_delivery, detail_decorate_root = parse_product_page(
                    full_product_url)

                # Добавляем данные в таблицу
                data.append(
                    [title, full_product_url, image1_url, image2_url, image3_url, image4_url, image5_url, image6_url,
                     f'{diameter1}  {diameter2}  {diameter3} {diameter4}  {diameter5}  {diameter6}', f'{length1}  {length2}  {length3} {length4}  {length5}  {length6}', type, key_attributes, other_attributes,
                     packaging_and_delivery, detail_decorate_root])
        except Exception as e:
            print(f'Ошибка парсинга товара: {e}')
            continue


# Цикл по страницам товаров
for page_num in range(1, 16):
    page_url = f'{main_url}productgrouplist-801458118-{page_num}/End_mill.html'
    print(f'Парсим раздел: {page_url}')
    parse_products(page_url)

# Сохраняем данные в Excel
with xlsxwriter.Workbook('result.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    # Записываем данные в Excel
    for row_num, info in enumerate(data):
        worksheet.write_row(row_num, 0, info)

print("Парсинг завершен, данные сохранены в 'result.xlsx'.")

# Закрываем драйвер
driver.quit()