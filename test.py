import json
import csv
import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# Функция для отправки запроса в API и получения данных о компаниях
def get_companies_data(page, searchText, loc, pageNumber):
    # Подготовить данные payload для запроса
    payload = {
        'number': pageNumber,
        'loc': loc,
        'searchText': searchText,
        'isForReplies': 'false'
    }
    print(payload)
    # Выполнить POST-запрос к API
    response = page.request.post(
        "https://www.orgpage.ru/Search/GetSearchCompany/",
        data=json.dumps(payload),
        headers={"Content-Type": "application/json;charset=UTF-8"}
    )

    # Выводим статус и текст ответа для диагностики
    # print(f"Response status: {response.status}")
    full_text = response.text()
    #print(f"Response text: {full_text[:1500]}...")  # Выводим только первые 500 символов для диагностики
    with open(f"HTML_data/page{pageNumber}.html", "w", encoding="utf-8") as f:
        f.write(full_text)
    # Используем BeautifulSoup для парсинга HTML
    soup = BeautifulSoup(full_text, 'html.parser')
    companies = []

    # Ищем все элементы компаний
    company_items = soup.find_all('div', class_='similar-item__title')
    for item in company_items:
        link_element = item.find('a')
        if link_element:
            link = link_element.get('href')
            if link:
                full_link = "https://www.orgpage.ru" + link  # Добавляем базовый URL, если это относительная ссылка
                companies.append(full_link)
                print(full_link)
    return companies

# Функция для сохранения состояния парсинга в файл
def save_state(file_name, state):
    with open(file_name, 'w') as f:
        json.dump(state, f)

# Функция для загрузки состояния парсинга из файла
def load_state(file_name):
    if os.path.exists(file_name):
        with open(file_name, 'r') as f:
            return json.load(f)
    return None

# Функция для парсинга данных компании
def parse_company_info(page, url, writer):
    print(f"Navigating to: {url}")
    page.goto(url)
    page.wait_for_timeout(2000)
    print(f'Parsing: {url}')

    try:
        name = page.query_selector("h1[itemprop='name']").inner_text().strip()
    except:
        name = "N/A"

    try:
        phones = [phone.inner_text().strip() for phone in page.query_selector_all("span[itemprop='telephone']")]
        phones = ', '.join(phones)
    except:
        phones = "N/A"

    try:
        site = page.query_selector("div.company-information__site-text p a.nofol-link").get_attribute("href")
    except:
        site = "N/A"

    try:
        email = page.query_selector("div.company-information__site-text p.email a[itemprop='email']").inner_text().strip()
    except:
        email = "N/A"

    # Write company information to CSV
    writer.writerow([name, phones, site, email])

if __name__ == "__main__":
    search_params = {
        'page1': {
            'searchText': 'сети домашнего текстиля',
            'loc': 'Россия'
        },
        'page2': {
            'searchText': 'спа салоны',
            'loc': 'Россия'
        },
        'page3': {
            'searchText': 'сети домашнего текстиля',
            'loc': 'Россия'
        },
        'page4': {
            'searchText': 'спа салоны',
            'loc': 'Россия'
        },
        'page5': {
            'searchText': 'салон штор',
            'loc': 'Россия'
        },
        'page6': {
            'searchText': 'магазины текстиля',
            'loc': 'Россия'
        },
        'page7': {
            'searchText': 'салон домашнего текстиля',
            'loc': 'Россия'
        },
        'page8': {
            'searchText': 'магазин постельного белья',
            'loc': 'Россия'
        },
        'page9': {
            'searchText': 'ковры интерьер',
            'loc': 'Россия'
        },
        'page10': {
            'searchText': 'производители готового столового текстиля',
            'loc': 'Россия'
        },
        'page11': {
            'searchText': 'производители полотенец',
            'loc': 'Россия'
        },
        'page12': {
            'searchText': 'дистрибьютеры интерьерных тканей',
            'loc': 'Россия'
        },
        # Добавьте остальные страницы поиска по аналогии
    }

    state_file = "scraping_state.json"

    # Загрузить состояние парсинга
    state = load_state(state_file)
    if state is None:
        # Если состояния нет, начать сначала
        state = {
            'currentPageIndex': 0,
            'currentPos': 1,
            'pageFileNames': list(search_params.keys())  # Получить названия страниц из параметров поиска
        }

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        for i in range(state['currentPageIndex'], len(state['pageFileNames'])):
            current_page_name = state['pageFileNames'][i]
            search_text = search_params[current_page_name]['searchText']
            loc = search_params[current_page_name]['loc']

            data_file = f"{current_page_name}_companies_info.csv"
            with open(data_file, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["name", "phones", "site", "email"])  # Write header

                total_page_counter = state['currentPos']
                while True:
                    company_links = get_companies_data(page, search_text, loc, total_page_counter)
                    if not company_links:
                        break

                    for link in company_links:
                        if link:  # Проверяем, что ссылка не None
                            parse_company_info(page, link, writer)

                    total_page_counter += 1

                    # Закрыть и открыть браузер каждые 100 страниц
                    if total_page_counter % 100 == 0:
                        state['currentPageIndex'] = i
                        state['currentPos'] = total_page_counter
                        save_state(state_file, state)
                        browser.close()

                        # Подождать несколько секунд перед запуском нового браузера
                        time.sleep(3)
                        browser = p.chromium.launch()
                        page = browser.new_page()

            # Сбросить счетчик позиций для следующей страницы
            state['currentPos'] = 1
            save_state(state_file, state)

        browser.close()
