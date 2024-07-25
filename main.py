import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.webdriver.chrome.options import Options
import os
import csv
import time

# Function to parse company information
def parse_company_info(url, writer):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(url)

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[itemprop='name']")))
    try:
        name = driver.find_element(By.CSS_SELECTOR, "h1[itemprop='name']").text.strip()
    except:
        name = "N/A"

    try:
        phones = [phone.text.strip() for phone in driver.find_elements(By.CSS_SELECTOR, "span[itemprop='telephone']")]
        phones = ', '.join(phones)
    except:
        phones = "N/A"

    try:
        site = driver.find_element(By.CSS_SELECTOR, "div.company-information__site-text p a.nofol-link").get_attribute("href")
    except:
        site = "N/A"

    try:
        email = driver.find_element(By.CSS_SELECTOR, "div.company-information__site-text p.email a[itemprop='email']").text.strip()
    except:
        email = "N/A"

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    # Write company information to CSV
    writer.writerow([name, phones, site, email])

# Function to gather all company links on the page
def gather_company_links(writer):
    while True:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.object-item.similar-item"))
            )
            companies = driver.find_elements(By.CSS_SELECTOR, "div.object-item.similar-item")
            for company in companies:
                try:
                    link = company.find_element(By.CSS_SELECTOR, "a.img.decode-link").get_attribute("href")
                    writer.writerow([link])
                except Exception as e:
                    print(f"Failed to gather link: {e}")
                    continue
            break
        except Exception as e:
            print(f"Exception while gathering company links: {e}")
            time.sleep(5)

# Function to navigate to the next page
def go_to_next_page():
    next_button = None

    try:
        # Найти все элементы для навигации по страницам
        pagination_items = driver.find_elements(By.CSS_SELECTOR, "ul.footer-navigation.paging li")

        # Найти кнопку "Следующая" среди элементов
        for item in pagination_items:
            try:
                link = item.find_element(By.CSS_SELECTOR, "a")
                if 'Следующая' in link.text and 'disabled' not in item.get_attribute('class'):
                    next_button = link
                    break
            except:
                continue
    except:
        print('Не удалось найти пагинацию Варианта 1')

    if not next_button:
        try:
            # Прокрутка страницы до конца, чтобы прогрузилась вся пагинация
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.nav.paging.rubricator-paging"))
            )

            # Найти все элементы для навигации по страницам
            try:
                yellow_btn = driver.find_element(By.CSS_SELECTOR, 'div.btn-holder button.rubricator-next-button')
                driver.execute_script("arguments[0].scrollIntoView(true);", yellow_btn)
            except:
                print('ЖЕЛТАЯ КНОПКА НЕ НАЙДЕНА')

            time.sleep(random.randint(0, 2))
            next_button = driver.find_element(By.CSS_SELECTOR, "div.nav.paging.rubricator-paging a.gradient.next")
        except:
            print('Не удалось найти пагинацию Варианта 2')

    if next_button:
        try:
            # Нажатие на кнопку "Следующая"
            company_data = driver.find_element(By.CSS_SELECTOR, "div.object-item.similar-item")
            next_button.click()
            bad_try_counter = 0
            while True:
                if bad_try_counter > 50:
                    return False
                if company_data != driver.find_element(By.CSS_SELECTOR, "div.object-item.similar-item"):
                    return True
                else:
                    bad_try_counter += 1
                    time.sleep(1)
        except Exception as e:
            print(f"Не удалось перейти на следующую страницу: {e}")
            return False
    else:
        print("Следующая страница не найдена или кнопка отключена.")
        return False


if __name__ == "__main__":
    # Initialize the browser with options
    proxy = 'http://10.250.0.1:3128'
    options = Options()
    options.add_argument(f'--proxy-server={proxy}')
    extension_path = os.path.join(os.path.dirname(__file__), 'uBlockOrigin.crx')
    # options.add_argument('--headless=new')

    if os.path.exists(extension_path):
        options.add_extension(extension_path)
    else:
        raise FileNotFoundError(f"The extension file was not found at path: {extension_path}")

    # Base URL for parsing
    #base_url = "https://www.orgpage.ru/search.html?q=%D1%81%D0%B5%D1%82%D0%B8+%D0%B4%D0%BE%D0%BC%D0%B0%D1%88%D0%BD%D0%B5%D0%B3%D0%BE+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F"
    # 'https://www.orgpage.ru/moskva/%D0%BC%D0%B0%D0%B3%D0%B0%D0%B7%D0%B8%D0%BD%D1%8B_%D0%BE%D0%B1%D0%BE%D0%B5%D0%B2/',
    all_urls = [
        'https://www.orgpage.ru/search.html?q=%D1%81%D0%B5%D1%82%D0%B8+%D0%B4%D0%BE%D0%BC%D0%B0%D1%88%D0%BD%D0%B5%D0%B3%D0%BE+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D1%81%D0%BF%D0%B0+%D1%81%D0%B0%D0%BB%D0%BE%D0%BD%D1%8B&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/rossiya/stroitelnye-gipermarkety/',
        'https://www.orgpage.ru/search.html?q=%D1%81%D0%B0%D0%BB%D0%BE%D0%BD+%D1%88%D1%82%D0%BE%D1%80&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D0%BC%D0%B0%D0%B3%D0%B0%D0%B7%D0%B8%D0%BD%D1%8B+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D1%81%D0%B0%D0%BB%D0%BE%D0%BD+%D0%B4%D0%BE%D0%BC%D0%B0%D1%88%D0%BD%D0%B5%D0%B3%D0%BE+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D1%81%D0%B5%D1%82%D0%B8+%D0%B4%D0%BE%D0%BC%D0%B0%D1%88%D0%BD%D0%B5%D0%B3%D0%BE+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D0%BC%D0%B0%D0%B3%D0%B0%D0%B7%D0%B8%D0%BD+%D0%BF%D0%BE%D1%81%D1%82%D0%B5%D0%BB%D1%8C%D0%BD%D0%BE%D0%B3%D0%BE+%D0%B1%D0%B5%D0%BB%D1%8C%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F'
        'https://www.orgpage.ru/search.html?q=%D0%BA%D0%BE%D0%B2%D1%80%D1%8B+%D0%B8%D0%BD%D1%82%D0%B5%D1%80%D1%8C%D0%B5%D1%80&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D0%BF%D1%80%D0%BE%D0%B8%D0%B7%D0%B2%D0%BE%D0%B4%D0%B8%D1%82%D0%B5%D0%BB%D0%B8+%D0%B3%D0%BE%D1%82%D0%BE%D0%B2%D0%BE%D0%B3%D0%BE+%D1%81%D1%82%D0%BE%D0%BB%D0%BE%D0%B2%D0%BE%D0%B3%D0%BE+%D1%82%D0%B5%D0%BA%D1%81%D1%82%D0%B8%D0%BB%D1%8F&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D0%BF%D1%80%D0%BE%D0%B8%D0%B7%D0%B2%D0%BE%D0%B4%D0%B8%D1%82%D0%B5%D0%BB%D0%B8+%D0%BF%D0%BE%D0%BB%D0%BE%D1%82%D0%B5%D0%BD%D0%B5%D1%86&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F',
        'https://www.orgpage.ru/search.html?q=%D0%B4%D0%B8%D1%81%D1%82%D1%80%D0%B8%D0%B1%D1%8C%D1%8E%D1%82%D0%B5%D1%80%D1%8B+%D0%B8%D0%BD%D1%82%D0%B5%D1%80%D1%8C%D0%B5%D1%80%D0%BD%D1%8B%D1%85+%D1%82%D0%BA%D0%B0%D0%BD%D0%B5%D0%B9+&loc=%D0%A0%D0%BE%D1%81%D1%81%D0%B8%D1%8F'
        ]
    # Open the page

    first_itter = True
    for cur_url in all_urls:
        if first_itter == True:
            first_itter = False
        else:
            time.sleep(random.randint(180, 3000))
        # Create the ChromeDriver instance with custom options
        driver = webdriver.Chrome(options=options)
        driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled":True})
        driver.get(cur_url)

        print(f'Page opened: {cur_url}')

        file_name = driver.find_element(By.CSS_SELECTOR, 'h1.strong').text
        # Create CSV file for links
        links_file = "company_links.csv"
        with open(links_file, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["link"])  # Write header

            # Main loop to parse all pages
            total_page_counter = 0
            while True:
                total_page_counter += 1
                if total_page_counter % 15 == 0:
                    time.sleep(5)
                gather_company_links(writer)
                time.sleep(random.randint(0, 2))
                if go_to_next_page() == False:
                    break

        print("Finished gathering all company links.")

        # Process company links and save data
        data_file = "companies_info.csv"
        with open(data_file, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["name", "phones", "site", "email"])  # Write header

            with open(links_file, mode='r', encoding='utf-8') as links_csv:
                reader = csv.reader(links_csv)
                next(reader)  # Skip header
                for row in reader:
                    link = row[0]
                    parse_company_info(link, writer)

        # Close the browser
        driver.quit()

        print("Start converting CSV to XLSX...")
        # Convert CSV to XLSX
        df = pd.read_csv(data_file)
        df.to_excel(f"{file_name}.xlsx", index=False)

        print(f"Data saved to {file_name}.xlsx")
