import time

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import requests
import sys


class CategoryExtractor:
    def __init__(self, driver_path):
        self.driver_path = driver_path
        self.driver = None

    def initialize_driver(self):
        service = Service(self.driver_path)
        self.driver = webdriver.Chrome(service=service)

    def close_driver(self):
        if self.driver:
            self.driver.quit()

    def extract_categories(self, url):
        self.driver.get(url)
        time.sleep(3)
        html_content = self.driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        category_elements = soup.find_all('a', class_='vn-link vn-nav__link')

        categories = []
        for element in category_elements:

            if element.text.strip() == 'Zobacz wszystko':
                continue

            category = {}
            category['id'] = element['href'].split('/')[-2].split('-')[-1]
            category['text'] = element.text.strip()
            category['url'] = element['href']
            category['parentID'] = element.find_parent('ul')['id'].split('-')[-1]
            categories.append(category)

        return categories


# Пример использования класса CategoryExtractor
if __name__ == '__main__':

    # Разные счетчики!
    current_row = 2
    current_category_row = 10
    added_categories = []
    product_counter = 0

    # Создаем новый файл Excel и добавляем рабочий лист

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Export Products Sheet"

    # Указываем заголовки для столбцовgi
    headers = ['Код товара',  # 1
               'Url',  # 2_a
               'Название_позиции',  # 3
               'Название_позиции_укр',  # 4
               'Поисковые_запросы',  # 5
               'Поисковые_запросы_укр',  # 6
               'Описание',  # 7
               'Описание_укр',  # 8
               'Описание_pl',  # 8.2
               'Тип_товара',  # 9
               'Цена',  # 10
               'Валюта',  # 11
               'Единица_измерения',  # 12
               'Минимальный_объем_заказа',  # 13
               'Оптовая_цена',  # 14
               'Минимальный_заказ_опт',  # 15
               'Ссылка_изображения',  # 16
               'Наличие',  # 17
               'Количество',  # 18
               'Номер_группы',  # 19
               'Название_группы',  # 20
               'Адрес_подраздела',  # 21
               'Возможность_поставки',  # 22
               'Срок_поставки',  # 23
               'Способ_упаковки',  # 24
               'Способ_упаковки_укр',  # 25
               'Уникальный_идентификатор',  # 26
               'Идентификатор_товара',  # 27
               'Идентификатор_подраздела',  # 28
               'Идентификатор_группы',  # 29
               'Производитель',  # 30
               'Страна_производитель',  # 31
               'Скидка',  # 32
               'ID_группы_разновидностей',  # 33
               'Личные_заметки',  # 34
               'Продукт_на_сайте',  # 35
               'Cрок действия скидки от',  # 36
               'Cрок действия скидки до',  # 37
               'Цена от',  # 38
               'Ярлык',  # 39
               'HTML_заголовок',  # 40
               'HTML_заголовок_укр',  # 41
               'HTML_описание',  # 42
               'HTML_описание_укр',  # 43
               'HTML_ключевые_слова',  # 44
               'HTML_ключевые_слова_укр',  # 45
               'Вес,кг',  # 46
               'Ширина,см',  # 47
               'Высота,см',  # 48
               'Длина,см',  # 49
               'Где_находится_товар',  # 50
               'Код_маркировки_(GTIN)',  # 51
               'Номер_устройства_(MPN)',  # 52
               'Название_Характеристики',  # 53
               'Измерение_Характеристики',  # 54
               'Значение_Характеристики',  # 55
               'Название_Характеристики',  # 56
               'Измерение_Характеристики',  # 57
               'Значение_Характеристики',  # 58
               'Название_Характеристики',  # 59
               'Измерение_Характеристики',  # 60
               'Значение_Характеристики',  # 61
               ]

    ws.append(headers)
    ws2 = wb.create_sheet(title="Export Groups Sheet")

    # Указываем заголовки для столбцов

    headers = ['Номер_группы',  # 1
               'Название_группы',  # 2
               'Название_группы_укр',  # 3
               'Идентификатор_группы',  # 4
               'Номер_родителя',  # 5
               'Идентификатор_родителя',  # 6
               'HTML_заголовок_группы',  # 7
               'HTML_заголовок_группы_укр',  # 8
               'HTML_описание_группы',  # 9
               'HTML_описание_группы_укр',  # 10
               'HTML_ключевые_слова_группы',  # 11
               'HTML_ключевые_слова_группы_укр',  # 12
               ]

    ws2.append(headers)

    # Categories
    categories_url = 'https://www.ikea.com/pl/pl/cat/produkty-products/'
    driver_path = 'chromedriver.exe'

    extractor = CategoryExtractor(driver_path)
    extractor.initialize_driver()

    categories_list = extractor.extract_categories(categories_url)
    print(categories_list)

    extractor.close_driver()

    for category in categories_list:

        category_id = category['id']
        category_text = category['text']
        category_url = category['url']
        parent_category_id = category['parentID']

        product_url = ''
        changefreq = ''

        # Загружаем страницу категорий
        print("Parsing category: ", category_text)

        service = Service(driver_path)
        driver = webdriver.Chrome(service=service)

        driver.get(category_url + '?page=500')

        time.sleep(1)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Явное ожидание загрузки элементов с классом 'plp-product-card'
        wait = WebDriverWait(driver, 10)
        products = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'plp-fragment-wrapper')))

        product_data = []

        for product in products:
            number = product.find_element(By.CLASS_NAME, 'plp-mastercard  ').get_attribute('data-product-number')
            price = product.find_element(By.CLASS_NAME, 'plp-mastercard  ').get_attribute('data-price')
            link = product.find_element(By.TAG_NAME, 'a').get_attribute('href')

            product_info = {
                'number': number,
                'price': price,
                'link': link
            }
            product_data.append(product_info)

        for product in product_data:

            print(product_counter)

            time.sleep(1)

            product_counter += 1

            link = product['link']
            print("link", link)

            number = product['number']
            print("number", number)

            price = product['price']
            print("price", price)

            # Переходим на страницу товара
            driver.get(link)

            if product_counter == 1:
                time.sleep(1)

            try:
                # ID товара у нас уже сть
                print("Product ID is ", number)

                # Price у нас тоже есть
                print("Price is ", price)

                # Get Product Name
                NameElement = driver.find_element(By.CLASS_NAME, 'pip-header-section__title--big')
                productName = NameElement.text
                print(productName)

                # Get description
                description = driver.execute_script('''
                   const description = document.querySelector('.pip-product-details__container').innerHTML;
                   return description;
                    ''')
                print(description)

                # Get all amages

                # Get description
                images = driver.execute_script('''
                    let images = document.querySelectorAll('.pip-product__left-top .pip-image');
                    let urlString = '';
                    
                    images.forEach(img => {
                    // Получаем атрибут href каждого изображения
                    let url = img.getAttribute('src');
                    url = url.replace('?f=s', '');
                    urlString += url + ',';
                    });
                    
                    return urlString;
                ''')

                print(images)

                # check package counts
                packageCount = driver.execute_script('''
                    // Находим все элементы с классом "pip-product-dimensions__measurement-value"
                    const packageElements = document.querySelectorAll('.pip-product-dimensions__measurement-value');
                    let totalPackages = 0;

                    // Перебираем найденные элементы и суммируем количество упаковок
                    packageElements.forEach(packageElement => {
                        totalPackages += parseInt(packageElement.textContent);
                    });

                    // Возвращаем общее количество упаковок
                    return totalPackages;
                ''')

                if packageCount == 1:
                    width = driver.execute_script('''
                        const parent = document.querySelectorAll('.pip-product-dimensions__package-container');
                        const elements = parent[0].querySelectorAll('.pip-product-dimensions__measurement-wrapper');
                        originalString = elements[0].innerText;
                        const match = originalString.match(/\d+[.,]?\d*/);
                        if (match) {
                            const extractedNumber = match[0].replace('.', ',');
                            return extractedNumber;
                        } 
                    ''')

                    height = driver.execute_script('''
                        const parent = document.querySelectorAll('.pip-product-dimensions__package-container');
                        const elements = parent[0].querySelectorAll('.pip-product-dimensions__measurement-wrapper');
                        originalString = elements[1].innerText;
                        const match = originalString.match(/\d+[.,]?\d*/);
                        if (match) {
                            const extractedNumber = match[0].replace('.', ',');
                            return extractedNumber;
                        } 
                     ''')

                    length = driver.execute_script('''
                        const parent = document.querySelectorAll('.pip-product-dimensions__package-container');
                        const elements = parent[0].querySelectorAll('.pip-product-dimensions__measurement-wrapper');
                        originalString = elements[2].innerText;
                        const match = originalString.match(/\d+[.,]?\d*/);
                        if (match) {
                            const extractedNumber = match[0].replace('.', ',');
                            return extractedNumber;
                        } 
                    ''')

                    weight = driver.execute_script('''
                        const parent = document.querySelectorAll('.pip-product-dimensions__package-container');
                        const elements = parent[0].querySelectorAll('.pip-product-dimensions__measurement-wrapper');
                        originalString = elements[3].innerText;
                        const match = originalString.match(/\d+[.,]?\d*/);
                        if (match) {
                            const extractedNumber = match[0].replace('.', ',');
                            return extractedNumber;
                        } 
                    ''')

                    print(f"Вес: {weight}")
                    print(f"Ширина: {width}")
                    print(f"Высота: {height}")
                    print(f"Длина: {length}")


            except NoSuchElementException:
                print("Some field not found", link)
                continue  # Пропускаем этот товар и переходим к следующему

            ws.append([
                number,  # 1
                link,  # 2_a
                f'=GOOGLETRANSLATE(E{current_row},"UK","RU")',  # 3
                productName,  # 4
                '',  # 5
                '',  # 6
                f'=GOOGLETRANSLATE(I{current_row},"UK","RU")',  # 7 description
                '',  # 8
                description,  # 8.2
                '',  # 9
                price,  # 10
                'EUR',  # 11
                'шт.',  # 12
                '',  # 13
                '',  # 14
                '',  # 15
                images,  # 16
                '+',  # 17'
                '',  # 18
                '',  # 19
                '',  # 20
                '',  # 21
                '',  # 22
                '',  # 23
                '',  # 24
                '',  # 25
                number,  # 26
                '',  # 26
                '',  # 26
                category_id,  # 29 ИДЕНТИФИКАТОР ГРУППЫ
                'IKEA',  # 30
                '',  # 31
                '',  # 32
                '',  # 33
                '',  # 34
                '',  # 35
                '',  # 36
                '',  # 37
                '',  # 38
                '',  # 39
                '',  # 40
                '',  # 41
                '',  # 42
                '',  # 43
                '',  # 44
                '',  # 45
                '',  # 46
                '',  # 47
                '',  # 48
                '',  # 49
                '',  # 50
                '',  # 51
                '',  # 52
                '',  # 53
                '',  # 54
                '',  # 55
            ])

            # Проверяем, есть ли текущая категория в списке уже добавленных категорий
            if category_text not in added_categories:
                # Если категория ещё не добавлена, добавляем её в лист Excel и в список добавленных категорий
                ws2.append([
                    '',  # 1
                    f'=GOOGLETRANSLATE(C{current_row},"UK","RU")',  # 2
                    category_text,  # 3
                    category_id,  # 4
                    '',  # 5
                    parent_category_id,  # 6
                ])
                added_categories.append(category_text)  # Добавляем категорию в список добавленных категорий
                current_category_row += 1

            current_row += 1

            print("Data written for", product_url)

            # Сохраняем файл Excel после каждой записи
            wb.save('output.xlsx')

# Save the Excel file after processing all products
wb.save('output.xlsx')

# Закрываем веб-драйвер и сохраняем файл Excel
driver.quit()
wb.close()

print("Data saved to output.xlsx")
