import os
import re
import time
import random
import string
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup as bs
import plyer
from datetime import datetime
import argparse
import sys

class SocialMediaParser:
    def __init__(self, links_file, date_range):
        self.project_root = os.path.dirname(os.path.abspath(__file__))  # Корневая папка проекта
        self.profile_dir = os.path.join(self.project_root, "profile")  # Папка для профиля
        self.driver_dir = os.path.join(self.project_root, "driver")    # Папка для драйвера
        self.results_dir = os.path.join(self.project_root, "results")  # Папка для результатов
        self.links_file = links_file
        self.date_range = date_range  # Фиксированный диапазон дат
        self.driver = None
        self.data = []  # Список кортежей (link, numbers, labels)
        self.links = []

    def setup_driver(self):
        """Настройка драйвера Chrome и профиля в портативном режиме"""
        # Создание папок, если их нет
        os.makedirs(self.profile_dir, exist_ok=True)
        os.makedirs(self.driver_dir, exist_ok=True)
        os.makedirs(self.results_dir, exist_ok=True)

        # Настройка опций Chrome
        options = Options()
        options.add_argument(f"user-data-dir={self.profile_dir}")  # Используем локальный профиль
        options.add_argument("start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # Установка пути для сохранения ChromeDriver
        driver_path = os.path.join(self.driver_dir, "chromedriver.exe")
        
        # Скачивание и сохранение драйвера
        if not os.path.exists(driver_path):
            manager = ChromeDriverManager()
            downloaded_path = manager.install()  # Скачиваем драйвер в стандартное место
            if downloaded_path != driver_path:
                os.replace(downloaded_path, driver_path)
        
        # Использование локального драйвера
        service = Service(executable_path=driver_path)
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.get("https://popsters.ru/app/dashboard")
        
        # Ожидание авторизации
        print("Браузер открыт. Пожалуйста, авторизуйтесь на сайте popsters.ru.")
        print("После авторизации вернитесь в терминал и нажмите Enter для начала парсинга.")
        input("Нажмите Enter, чтобы продолжить...")

    def load_input_data(self):
        """Загрузка ссылок из текстового файла"""
        try:
            with open(self.links_file, 'r', encoding='utf-8') as file:
                self.links = [line.strip() for line in file.readlines() if line.strip()]
            if not self.links:
                raise ValueError("Файл ссылок пустой или содержит ошибки")
        except Exception as e:
            print(f"Ошибка загрузки входного файла: {e}")
            sys.exit(1)

    def parser(self, link):
        """Парсинг одной ссылки с фиксированным диапазоном дат"""
        try:
            self.driver.refresh()
            time.sleep(2)

            # Ввод ссылки
            input_tab = self.driver.find_element(By.TAG_NAME, 'textarea')
            input_tab.send_keys(link)
            time.sleep(2)

            # Нажатие на поиск
            self.driver.find_element(By.TAG_NAME, 'button').click()
            WebDriverWait(self.driver, 20).until(ec.presence_of_element_located(('id', 'datepicker')))
            
            # Установка фиксированного диапазона дат
            date_input = self.driver.find_element('id', 'datepicker')
            date_input.clear()
            date_input.send_keys(self.date_range)
            self.driver.find_element('xpath', '//button[@class="app-button r-button"]').click()
            WebDriverWait(self.driver, 600).until(ec.presence_of_element_located(('xpath', '//label[@for="v2"]')))

            # Нажатие на кнопку "Общие"
            self.driver.find_element('xpath', '//label[@for="v2"]').click()
            time.sleep(2)

            # Извлечение данных
            soup = bs(self.driver.page_source, 'html.parser')
            stat_links = soup.find_all('ul', class_='common-data')
            
            if not stat_links:
                self.data.append((link, ['0'], ['Нет данных']))
            else:
                text = stat_links[0].text.replace(" ", "")  # Берем первый ul с классом common-data
                numbers = re.findall(r'\d+', text)
                labels = re.findall(r'[А-Яа-я]+', text)
                if not numbers or not labels:
                    self.data.append((link, ['0'], ['Нет данных']))
                else:
                    # Синхронизируем числа и метки (если меток больше, чем чисел, оставляем лишние метки как None)
                    min_length = min(len(numbers), len(labels))
                    numbers = numbers[:min_length]
                    labels = labels[:min_length]
                    self.data.append((link, numbers, labels))

        except Exception as e:
            print(f"Ошибка парсинга {link}: {e}")
            self.data.append((link, ['0'], ['Нет данных']))
            self.driver.refresh()

    def process_data(self):
        """Обработка всех ссылок с фиксированным диапазоном дат"""
        for i, link in enumerate(self.links, 1):
            self.parser(link)
            remaining = len(self.links) - i
            print(f"Обработана страница {link}. Осталось: {remaining} ссылок.")

    def clean_data(self):
        """Очистка и форматирование извлечённых данных с правильной разбивкой по столбцам"""
        data = []
        all_labels = set()  # Собираем все уникальные метки

        # Собираем данные и метки
        for link, numbers, labels in self.data:
            row_data = {'link': link, 'date': self.date_range}
            if labels[0] != 'Нет данных':
                for num, label in zip(numbers, labels):
                    row_data[label] = num
                all_labels.update(labels)
            else:
                row_data['Нет данных'] = numbers[0] if numbers else '0'
            data.append(row_data)

        # Создаём DataFrame с динамическими столбцами
        df = pd.DataFrame(data)

        # Убедимся, что все возможные метки присутствуют как столбцы
        for label in all_labels:
            if label not in df.columns:
                df[label] = 0

        # Заполняем пропуски нулями
        df = df.fillna(0)

        # Перемещаем столбцы 'link' и 'date' в начало
        cols = ['link', 'date'] + [col for col in df.columns if col not in ['link', 'date']]
        df = df[cols]

        return df

    def save_results(self, df):
        """Сохранение результатов в Excel с датой и временем в формате ДД.ММ.ГГГГ ЧЧ.ММ"""
        # Определение платформы
        platform = 'unknown'
        if any('vk.com' in link for link in df['link']):
            platform = 'vk'
        elif any('instagram' in link for link in df['link']):
            platform = 'instagram'
        elif any('t.me' in link for link in df['link']):
            platform = 'telegram'

        # Формирование имени файла с рандомными числами и датой/временем в формате ДД.ММ.ГГГГ ЧЧ.ММ
        random_num = str(random.randint(1, 1000))
        timestamp = datetime.now().strftime("%d.%m.%Y %H.%M")
        filename = f"{platform}_{timestamp}.xlsx"
        output_path = os.path.join(self.results_dir, filename)

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='pars', index=False)

        print(f"Результаты сохранены в {output_path}")

    def run(self):
        """Основной метод выполнения"""
        self.setup_driver()  # Сначала открываем браузер и ждём авторизацию
        self.load_input_data()
        self.process_data()
        df = self.clean_data()
        self.save_results(df)

        # Уведомление
        plyer.notification.notify(
            message='Парсинг завершён',
            title='Парсер готов'
        )

    def __del__(self):
        """Очистка ресурсов"""
        if self.driver:
            self.driver.quit()

def main():
    parser = argparse.ArgumentParser(description='Парсер данных социальных сетей в портативном режиме с фиксированным диапазоном дат')
    parser.add_argument('--links', default='links.txt', help='Путь к файлу со ссылками')
    parser.add_argument('--date', default='01.02.2025-28.02.2025', help='Фиксированный диапазон дат (формат ДД.ММ.ГГГГ-ДД.ММ.ГГГГ)')

    args = parser.parse_args()

    try:
        parser_instance = SocialMediaParser(args.links, args.date)
        parser_instance.run()
    except Exception as e:
        print(f"Произошла ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()