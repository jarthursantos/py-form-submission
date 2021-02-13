# System
import os, time, random

# Selenium
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException

# XLSX
import pandas as pd

# reCAPTCHA
from twocaptcha import TwoCaptcha

# ENV
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv('API_KEY')

option = Options()
option.headless = True

browser = Firefox(options=option)
browser.implicitly_wait(10)

def handle(url, name, mail, message, phone):
    browser.get(url)

    try:
        nameEl = browser.find_element_by_css_selector("form#contact_form input#ContactFormName")
    except NoSuchElementException as e:
        nameEl = browser.find_element_by_css_selector("form#contact_form input#contactFormName")
    finally:
        nameEl.send_keys(name)

    try:
        mailEl = browser.find_element_by_css_selector("form#contact_form input#ContactFormEmail")
    except NoSuchElementException as e:
        mailEl = browser.find_element_by_css_selector("form#contact_form input#contactFormEmail")
    finally:
        mailEl.send_keys(mail)

    try:
        phoneEl = browser.find_element_by_css_selector("form#contact_form input#ContactFormPhone")
        phoneEl.send_keys(phone)
    except NoSuchElementException as e:
        pass

    try:
        messageEl = browser.find_element_by_css_selector("form#contact_form textarea#ContactFormMessage")
    except NoSuchElementException as e:
        messageEl = browser.find_element_by_css_selector("form#contact_form textarea#contactFormMessage")
    finally:
        messageEl.send_keys(message)

    submitEl = browser.find_element_by_css_selector("form#contact_form input[type='submit']")
    submitEl.click()

    reCAPTCHAEl = browser.find_element_by_css_selector("div#g-recaptcha > div > div > iframe")

    reURL = reCAPTCHAEl.get_attribute('src')
    query = reURL.split('?')[1].split('&')
    site_key = ''

    for arg in query:
        if arg.startswith('k='):
            site_key = arg[2:]

    if site_key == '':
        print("site_key not found")
        return

    solver = TwoCaptcha(api_key)

    try:
        result = solver.recaptcha(sitekey=site_key, url=browser.current_url)
    except Exception as e:
        print('error: ' + str(e))
    else:
        browser.execute_script("document.querySelector('{}').textContent = '{}'".format("form textarea#g-recaptcha-response", result['code']))

        submitEl = browser.find_element_by_css_selector("form input[type='submit']")
        submitEl.click()

        time.sleep(1)

        with open('./results.txt', 'a') as result:
            result.write(browser.current_url + '\n')
            result.close()
        

urls = []

with open('forms.txt', 'r') as forms:
    for form in forms:
        urls.append(form)

xlsx = pd.ExcelFile("./data.xlsx", engine="openpyxl")
sheet = xlsx.parse('Sheet1')
sheet_data = sheet.to_dict()

for url in urls:
    for row in range(0, len(sheet)):
        name=sheet_data['name'][row]
        mail=sheet_data['email'][row]
        message=sheet_data['message'][row]
        phone=sheet_data['phone'][row]

        handle(url, name, mail, message, phone)

        time.sleep(random.randrange(10, 15))


browser.quit()
