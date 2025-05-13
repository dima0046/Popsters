import os
import re
import time
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
import sys
import glob
import shutil

class SocialMediaParser:
    def __init__(self):
        self.project_root = os.path.dirname(os.path.abspath(__file__))  # Корневая папка проекта
        self.source_dir = os.path.join(self.project_root, "source")    # Папка для исходного файла
        self.profile_dir = os.path.join(self.project_root, "profile")  # Папка для профиля
        self.driver_dir = os.path.join(self.project_root, "driver")    # Папка для драйвера
        self.results_dir = os.path.join(self.project_root, "results")  # Папка для результатов
        self.dates_file = os.path.join(self.project_root, "dates.txt") # Файл с диапазонами дат
        self.date_ranges = self.load_date_ranges()  # Загружаем диапазоны дат
        self.driver = None
        self.data = []  # Список кортежей (link, numbers, labels, date_range)
        self.links = []
        self.supported_platforms = ['VK', 'Facebook', 'Instagram', 'Telegram', 'Youtube', 'OK']  # Поддерживаемые платформы

    def load_date_ranges(self):
        """Загрузка диапазонов дат из файла dates.txt в формате ДД.ММ.ГГГГ - ДД.ММ.ГГГГ"""
        try:
            if not os.path.exists(self.dates_file):
                raise FileNotFoundError("Файл dates.txt не найден в корневой директории проекта")
            
            with open(self.dates_file, 'r', encoding='utf-8') as file:
                date_lines = file.read().strip().splitlines()
            
            date_ranges = []
            for date_range in date_lines:
                date_range = date_range.strip()
                # Проверка базового формата с пробелами вокруг тире
                if not re.match(r'^\d{2}\.\d{2}\.\d{4}\s+-\s+\d{2}\.\d{2}\.\d{4}$', date_range):
                    raise ValueError(f"Неверный формат диапазона дат: {date_range}. Ожидается ДД.ММ.ГГГГ - ДД.ММ.ГГГГ с пробелами вокруг тире")
                date_ranges.append(date_range)
            
            if not date_ranges:
                raise ValueError("В dates.txt не указаны корректные диапазоны дат.")
            
            return date_ranges
        except Exception as e:
            print(f"Ошибка загрузки диапазонов дат из dates.txt: {e}")
            sys.exit(1)

    def setup_driver(self):
        """Настройка драйвера Chrome и профиля в портативном режиме с улучшенной маскировкой"""
        # Создание папок, если их нет
        os.makedirs(self.profile_dir, exist_ok=True)
        os.makedirs(self.driver_dir, exist_ok=True)
        os.makedirs(self.results_dir, exist_ok=True)
        os.makedirs(self.source_dir, exist_ok=True)

        # Настройка опций Chrome для маскировки
        options = Options()
        options.add_argument(f"user-data-dir={self.profile_dir}")  # Используем локальный профиль
        options.add_argument("start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.7103.93 Safari/537.36")

        # Автоматическое обновление ChromeDriver
        try:
            # Установка пути для сохранения ChromeDriver
            driver_path = os.path.join(self.driver_dir, "chromedriver.exe")
            
            # Скачиваем последнюю версию ChromeDriver, совместимую с установленным Chrome
            manager = ChromeDriverManager()
            downloaded_path = manager.install()
            
            # Перемещаем скачанный ChromeDriver в нужную папку
            if downloaded_path != driver_path and os.path.exists(downloaded_path):
                if os.path.exists(driver_path):
                    os.remove(driver_path)  # Удаляем старый драйвер, если он есть
                shutil.move(downloaded_path, driver_path)
            
            # Использование локального драйвера
            service = Service(executable_path=driver_path)
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.get("https://popsters.ru/app/dashboard")
            
            # Ожидание авторизации
            print("Браузер открыт. Пожалуйста, авторизуйтесь на сайте popsters.ru.")
            print("После авторизации вернитесь в терминал и нажмите Enter для начала парсинга.")
            input("Нажмите Enter, чтобы продолжить...")
        except Exception as e:
            print(f"Ошибка настройки драйвера: {e}")
            sys.exit(1)

    def load_input_data(self):
        """Загрузка ссылок из Excel-файла в папке source"""
        try:
            # Ищем любой .xlsx файл в папке source
            excel_files = glob.glob(os.path.join(self.source_dir, "*.xlsx"))
            if not excel_files:
                raise FileNotFoundError("В папке source не найден файл .xlsx")
            if len(excel_files) > 1:
                print(f"Найдено несколько .xlsx файлов в папке source: {excel_files}")
                print(f"Будет использован первый файл: {excel_files[0]}")

            # Читаем первый найденный файл
            self.original_df = pd.read_excel(excel_files[0])
            if 'Площадка' not in self.original_df.columns or 'Ссылка' not in self.original_df.columns:
                raise ValueError("В файле отсутствуют столбцы 'Площадка' или 'Ссылка'")

            # Добавляем недостающие столбцы в исходный DataFrame
            for metric in ['Подписчики', 'Лайки', 'Репосты', 'Комментарии', 'Просмотры', 'Публикации']:
                if metric not in self.original_df.columns:
                    self.original_df[metric] = 0

            # Сохраняем все ссылки, но парсим только поддерживаемые
            self.all_links = self.original_df[['Площадка', 'Ссылка']].to_dict('records')
            self.links = []
            for row in self.all_links:
                platform = row['Площадка']
                link = row['Ссылка']
                if platform in self.supported_platforms:
                    self.links.append(link)

            if not self.links:
                print("Не найдено ссылок для поддерживаемых платформ (VK, Facebook, Instagram, Telegram, Youtube, OK). Все ссылки будут перенесены в выходной файл без изменений.")

        except Exception as e:
            print(f"Ошибка загрузки входного файла: {e}")
            sys.exit(1)

    def parser(self, link, date_range):
        """Парсинг одной ссылки с фиксированным диапазоном дат"""
        try:
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: Обновление страницы | Результат: Начало обработки")
            self.driver.refresh()
            time.sleep(2)

            # Ввод ссылки
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: textarea (поле ввода ссылки) | Результат: Поиск элемента")
            input_tab = self.driver.find_element(By.TAG_NAME, 'textarea')
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: textarea (поле ввода ссылки) | Результат: Элемент найден, ввод ссылки")
            input_tab.send_keys(link)
            time.sleep(2)

            # Нажатие на поиск
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: button (кнопка поиска) | Результат: Поиск элемента")
            self.driver.find_element(By.TAG_NAME, 'button').click()
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: button (кнопка поиска) | Результат: Кнопка нажата, ожидание datepicker")
            WebDriverWait(self.driver, 20).until(ec.presence_of_element_located(('id', 'datepicker')))
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: id=datepicker (поле даты) | Результат: Элемент найден")
            
            # Установка фиксированного диапазона дат в формате с пробелами
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: id=datepicker (поле даты) | Результат: Очистка поля")
            date_input = self.driver.find_element('id', 'datepicker')
            date_input.clear()
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: id=datepicker (поле даты) | Результат: Ввод диапазона дат: {date_range}")
            date_input.send_keys(date_range)
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: xpath=//button[@class='app-button r-button'] (кнопка применения даты) | Результат: Поиск элемента")
            self.driver.find_element('xpath', '//button[@class="app-button r-button"]').click()
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: xpath=//button[@class='app-button r-button'] (кнопка применения даты) | Результат: Кнопка нажата, ожидание загрузки данных")
            WebDriverWait(self.driver, 600).until(ec.presence_of_element_located(('xpath', '//label[@for="v2"]')))
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: xpath=//label[@for='v2'] (кнопка 'Общие') | Результат: Элемент найден")

            # Нажатие на кнопку "Общие"
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: xpath=//label[@for='v2'] (кнопка 'Общие') | Результат: Нажатие на кнопку")
            self.driver.find_element('xpath', '//label[@for="v2"]').click()
            time.sleep(2)

            # Извлечение данных
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: HTML-страница (извлечение данных) | Результат: Парсинг страницы")
            soup = bs(self.driver.page_source, 'html.parser')
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: ul.common-data (статистика) | Результат: Поиск элементов")
            stat_links = soup.find_all('ul', class_='common-data')
            
            if not stat_links:
                print(f"Ссылка: {link} | Дата: {date_range} | Элемент: ul.common-data (статистика) | Результат: Элементы не найдены, возвращаем 'Нет данных'")
                self.data.append((link, ['0'], ['Нет данных'], date_range))
            else:
                print(f"Ссылка: {link} | Дата: {date_range} | Элемент: ul.common-data (статистика) | Результат: Элементы найдены, извлечение текста")
                text = stat_links[0].text.replace(" ", "")  # Берем первый ul с классом common-data
                numbers = re.findall(r'\d+', text)
                labels = re.findall(r'[А-Яа-я]+', text)
                if not numbers or not labels:
                    print(f"Ссылка: {link} | Дата: {date_range} | Элемент: ul.common-data (статистика) | Результат: Данные не извлечены (нет чисел или меток), возвращаем 'Нет данных'")
                    self.data.append((link, ['0'], ['Нет данных'], date_range))
                else:
                    # Синхронизируем числа и метки
                    min_length = min(len(numbers), len(labels))
                    numbers = numbers[:min_length]
                    labels = labels[:min_length]
                    print(f"Ссылка: {link} | Дата: {date_range} | Элемент: ul.common-data (статистика) | Результат: Данные извлечены - Числа: {numbers}, Метки: {labels}")
                    self.data.append((link, numbers, labels, date_range))

        except Exception as e:
            print(f"Ссылка: {link} | Дата: {date_range} | Элемент: Неизвестно | Результат: Ошибка: {str(e)}")
            self.data.append((link, ['0'], ['Нет данных'], date_range))
            self.driver.refresh()

    def process_data(self):
        """Обработка всех ссылок с фиксированными диапазонами дат"""
        total_iterations = len(self.links) * len(self.date_ranges)
        iteration = 0
        for link in self.links:
            for date_range in self.date_ranges:
                iteration += 1
                self.parser(link, date_range)
                remaining = total_iterations - iteration
                print(f"Обработана страница {link} для диапазона {date_range}. Осталось: {remaining} итераций.")

    def clean_data(self):
        """Очистка и форматирование извлечённых данных"""
        # Сопоставление меток Popsters с колонками итогового файла
        label_mapping = {
            'Подписчиков': 'Подписчики',
            'Всеголайков': 'Лайки',
            'Всегорепостов': 'Репосты',
            'Всегокомментариев': 'Комментарии',
            'Всегопросмотров': 'Просмотры',
            'Постов': 'Публикации'
        }

        data = []
        for link, numbers, labels, date_range in self.data:
            row_data = {'Ссылка': link, 'Диапазон дат': date_range}
            print(f"Обработка данных для ссылки: {link} | Диапазон: {date_range}")
            print(f"Исходные метки: {labels}, Числа: {numbers}")

            if labels[0] == 'Нет данных':
                print(f"Ссылка: {link} | Диапазон: {date_range} | Результат: Нет данных, все метрики = 0")
                for metric in ['Подписчики', 'Лайки', 'Репосты', 'Комментарии', 'Просмотры', 'Публикации']:
                    row_data[metric] = 0
            else:
                # Маппим метки на ожидаемые столбцы
                for num, label in zip(numbers, labels):
                    if label in label_mapping:
                        mapped_label = label_mapping[label]
                        row_data[mapped_label] = int(num)
                        print(f"Ссылка: {link} | Диапазон: {date_range} | Метка: {label} → {mapped_label} | Значение: {num}")
                    else:
                        print(f"Ссылка: {link} | Диапазон: {date_range} | Метка: {label} | Результат: Метка игнорируется")

                # Заполняем недостающие метрики нулями
                for metric in ['Подписчики', 'Лайки', 'Репосты', 'Комментарии', 'Просмотры', 'Публикации']:
                    if metric not in row_data:
                        row_data[metric] = 0
                        print(f"Ссылка: {link} | Диапазон: {date_range} | Метка: {metric} | Результат: Не найдена, установлено значение 0")

            data.append(row_data)

        # Создаём DataFrame с динамическими столбцами для спарсенных данных
        parsed_df = pd.DataFrame(data)
        print("Создан промежуточный DataFrame:")
        print(parsed_df)

        # Приводим структуру к исходному файлу, фильтруя только поддерживаемые платформы
        result_df = self.original_df.copy()
        result_df = result_df[result_df['Площадка'].isin(self.supported_platforms)]
        print("Фильтрация исходного DataFrame, оставлены только поддерживаемые платформы:")
        print(result_df)

        # Добавляем столбец с диапазоном дат
        parsed_df['key'] = parsed_df['Ссылка'] + '_' + parsed_df['Диапазон дат']
        result_df['key'] = result_df['Ссылка'] + '_' + result_df.apply(lambda x: self.date_ranges[0] if self.date_ranges else '', axis=1)

        # Обновляем значения только для спарсенных ссылок
        for metric in ['Подписчики', 'Лайки', 'Репосты', 'Комментарии', 'Просмотры', 'Публикации']:
            if metric in parsed_df.columns:
                metric_dict = parsed_df.set_index('key')[metric].to_dict()
                result_df[metric] = result_df['key'].map(metric_dict).fillna(result_df[metric]).astype(int)
                print(f"Обновление столбца {metric} в итоговом DataFrame:")
                print(result_df[['Ссылка', metric]])
            else:
                print(f"Столбец {metric} отсутствует в parsed_df, все значения остаются 0")

        # Добавляем столбец с диапазоном дат в итоговый DataFrame
        result_df['Диапазон дат'] = result_df['key'].map(parsed_df.set_index('key')['Диапазон дат'])
        result_df = result_df.drop(columns=['key'])

        return result_df

    def save_results(self, df):
        """Сохранение результатов в Excel"""
        output_path = os.path.join(self.results_dir, "results.xlsx")

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

        # Ожидание ввода перед закрытием браузера
        input("Нажмите Enter, чтобы продолжить...")

    def __del__(self):
        """Очистка ресурсов"""
        # Проверка на существование атрибута driver перед вызовом quit()
        if hasattr(self, 'driver') and self.driver:
            self.driver.quit()

def main():
    try:
        parser_instance = SocialMediaParser()
        parser_instance.run()
    except Exception as e:
        print(f"Произошла ошибка: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()