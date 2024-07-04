from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from email import encoders
import pandas as pd
import smtplib
import aiohttp
import asyncio
import os
import re

# Anmeldedaten und URLs
username = 'info@caipi.de'
password = 'tca.vxn7ztr8ypr6PAX'
login_url = 'https://www.nebelung.de/account/login'
shop_url = 'https://www.nebelung.de/shop/'
base_url = 'https://www.nebelung.de'

# Email-Konfiguration
email_user = 'webscrapertestnebelungen@gmail.com'
email_password = 'orxa gayk uurh cdkk'
email_send_to = 'steglichmaximilian@gmail.com'
email_subject = 'Nebelung Artikel Update'
smtp_server = 'smtp.gmail.com'
smtp_port = 465  # Using SSL port

# Funktion zum Abrufen der Anmeldeseite und Extrahieren versteckter Eingabefelder
async def get_hidden_inputs(session):
    async with session.get(login_url) as response:
        soup = BeautifulSoup(await response.text(), 'html.parser')
        hidden_inputs = soup.find_all('input', type='hidden')
        login_payload = {hidden_input.get('name'): hidden_input.get('value') for hidden_input in hidden_inputs if hidden_input.get('name') and hidden_input.get('value')}
        login_payload.update({'email': username, 'password': password})
        return login_payload

# Funktion zum Anmelden und Abrufen der ersten Kategorie-URL
async def login_and_get_first_category_url(session, login_payload):
    async with session.post(login_url, data=login_payload) as response:
        if response.status != 200 or "account/logout" not in await response.text():
            print("Fehler bei der Anmeldung.")
            exit()
    
    async with session.get(shop_url) as response:
        soup = BeautifulSoup(await response.text(), 'html.parser')
        first_category_link = soup.find('a', class_='btn btn-primary btn-md btn-block')
        return base_url + first_category_link.get('href') if first_category_link else None

# Funktion zum Abrufen der ersten Produkt-URL aus der ersten Kategorie
async def get_first_product_url(session, category_url):
    async with session.get(category_url) as response:
        soup = BeautifulSoup(await response.text(), 'html.parser')
        first_product_link = soup.find('a', class_='product-name')
        return first_product_link.get('href') if first_product_link else None

# Funktion zum Scrapen des ersten Artikels
async def scrape_first_product_page(session, product_url):
    async with session.get(product_url) as response:
        if response.status != 200:
            return {}

        soup = BeautifulSoup(await response.text(), 'html.parser')
        product_details = soup.find('div', class_='col-lg-5 product-detail-buy')
        product_description = soup.find('div', class_='col-md-7 product-detail-description-text')

        # Get all fields:
        product_manufacturer_a = product_details.find('a', class_='product-detail-manufacturer-link')
        product_name_h1 = product_details.find('h1', class_='product-detail-name')
        product_price_p = product_details.find('p', class_='product-detail-price')
        product_dimension_div = product_details.find('div', class_='icon-information d-flex flex-row')
        product_dimension_span = product_dimension_div.find('span', class_=False) if product_dimension_div else None
        product_number_span = product_details.find('span', class_='product-detail-ordernumber')
        product_number_short_span = product_details.find('div', class_='product-detail-ordernumber-short-container')
        product_content_span = product_details.find('span', class_='price-unit-content')
        product_details_span = product_details.find('span', class_='product-detail-recommended-amount-content')
        product_availabilty_div = product_details.find('div', class_='product-availability-immediatly')
        product_availabilty_span = product_availabilty_div.find('span', class_=False) if product_availabilty_div else None
        product_information_p = product_description.find('p', class_=False)

        product_info = {
            'Product_Manufacturer': product_manufacturer_a.text.strip() if product_manufacturer_a else None,
            'Product_Name': product_name_h1.text.strip() if product_name_h1 else None,
            'Product_Price': product_price_p.text.strip() if product_price_p else None,
            'Product_Dimension': re.sub(r'\s*([x()])\s*', r' \1 ', product_dimension_span.text.strip()) if product_dimension_span else None,
            'Product_Number': product_number_span.text.strip() if product_number_span else None,
            'Product_Number_Short': product_number_short_span.text.replace('Produktnummer (Kurzform):', '').strip() if product_number_short_span else None,
            'Product_Content': product_content_span.text.strip() if product_content_span else None,
            'Product_Recommended_Order_Quantity': product_details_span.text.strip() if product_details_span else None,
            'Product_Availabilty': product_availabilty_span.text.strip() if product_availabilty_span else None,
            'Product_Information': product_information_p.text.strip() if product_information_p else None,
        }
        
        return product_info

async def main():
    async with aiohttp.ClientSession() as session:
        # Anmeldeinformationen abrufen und anmelden
        login_payload = await get_hidden_inputs(session)
        first_category_url = await login_and_get_first_category_url(session, login_payload)
        
        if first_category_url:
            first_product_url = await get_first_product_url(session, first_category_url)
            
            if first_product_url:
                product_info = await scrape_first_product_page(session, first_product_url)
                
                if product_info:
                    print(product_info)
                else:
                    print("Fehler beim Scrapen des Produkts.")
            else:
                print("Keine Produkte gefunden.")
        else:
            print("Keine Kategorien gefunden.")

# Asynchrones Hauptprogramm starten
asyncio.run(main())