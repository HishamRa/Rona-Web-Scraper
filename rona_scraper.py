# importing libraries
import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from fractions import Fraction
import time

# setting up driver
options = Options()
options.add_argument('--headless')
options.add_argument('--window-size=1920,1080')
options.add_argument('--start maximized')
driver = webdriver.Chrome(options=options)
url = "https://www.rona.ca/"
driver.get(url)
# wait five seconds
time.sleep(5)

# importing dataframe and converting it to a list
rona_df_orig = pd.read_excel('rona_items.xlsx', converters={'Rona Item Number': str})
rona_df = rona_df_orig.iloc[0:, 0].tolist()
rona_items = rona_df

# creating lists for the things to be scraped
prices_int = []
prices_dec = []
links = []
models = []
titles = []
promo_statuses = []
urls = []

# defining length of items in the list
length = len(rona_df_orig.iloc[0:, 0])
print('Success! Rona item scraping will begin for', length, 'items...')
n = 0

# loop to iterate over all items in rona_items
for items in rona_items:

    # identifying search box
    search_box = driver.find_element(By.ID, "keywords")
    search_box.clear()
    search_box.send_keys(items)
    search_box.submit()
    # close popup if detected
    try:
        popup = driver.find_element(By.CLASS_NAME, 'signIn')
        if popup.is_displayed():
            driver.find_element(By.CLASS_NAME, 'close.closeEnews.js-close-modal').click()
        else:
            continue
    except NoSuchElementException:
        pass

    # parsing html content into BS
    url = driver.current_url
    response = requests.get(url)

    soup = BeautifulSoup(response.content, 'html.parser')

    # check if item is on site
    try:
        avail = driver.find_element(By.CLASS_NAME, 'productDetails.js-addToCart-loader')
        if avail.is_displayed():
            # scrape the data
            title = soup.find('h1', attrs={'itemprop': 'name'})
            titles.append(title.get_text())
            # scraping the model
            model = soup.find('meta', attrs={'itemprop': 'mpn'})
            model_act = model['content']
            price_int = soup.find('span', attrs={'class': 'price-box__price__amount__integer'})
            # titles.append(title.get_text())
            models.append(model_act)
            prices_int.append(price_int.get_text())

            # identifying if the price has a decimal portion and scraping it if it does
            try:
                decimal = driver.find_element(By.CLASS_NAME, "price-box__price__amount__decimal")
                if decimal.is_displayed():
                    price_dec = soup.find('span', attrs={'class': 'price-box__price__amount'}).find('sup', attrs={
                        'class': 'price-box__price__amount__decimal'})
                    if price_dec is not None:
                        prices_dec.append(price_dec.get_text())
                    else:
                        prices_dec.append('0')
                else:
                    prices_dec.append('0')
            except NoSuchElementException:
                prices_dec.append('0')

            # identifying if the item is on promotion
            try:
                promo = driver.find_element(By.CLASS_NAME, 'price-box__price__amount')
                if promo.is_displayed():
                    promo_status = soup.find('div', attrs={'class': 'product_price_container'}).find('div', attrs={
                        'class': 'price-box__regularPrice'})
                    if promo_status is not None:
                        promo_statuses.append('Item is on promotion!')
                    else:
                        promo_statuses.append('Item not on promotion!')
                else:
                    promo_statuses.append('Item not on promotion!')
            except NoSuchElementException:
                promo_statuses.append('Item not on promotion!')

            # getting item's url
            urls.append(url)
        else:
            models.append(' ')
            titles.append(' ')
            prices_int.append(' ')
            prices_dec.append(' ')
            promo_statuses.append(' ')
            urls.append(' ')
    except NoSuchElementException:
        models.append(' ')
        titles.append(' ')
        prices_int.append(' ')
        prices_dec.append(' ')
        promo_statuses.append(' ')
        urls.append(' ')

    n += 1
    print('Rona item', n, '/', length, 'has been scraped successfully.')

dict_rona = {'Rona Item Number': rona_items, 'Catalog Number': models, 'Product Name': titles, 'Integer': prices_int,
             'Decimal': prices_dec, 'Promo Status': promo_statuses, 'Item URL': urls}

driver.quit()

# creating a dataframe from dictionary
df = pd.DataFrame(dict_rona)

# adding both integer and decimal values

df['Integer'] = df['Integer'].replace(' ', 0, regex=False)
df['Decimal'] = df['Decimal'].replace(' ', 0, regex=False)
df['Integer'] = df['Integer'].replace(',', '', regex=True)
df['Decimal'] = df['Decimal'].astype(int)
df['Integer'] = df['Integer'].astype(int)
df['Decimal'] = df['Decimal'] / 100.0
df['Price ($)'] = df['Integer'] + df['Decimal']
df['Price ($)'] = df['Price ($)'].apply(lambda x: "{:.2f}".format(x))
df['Price ($)'].astype(str)
df['Price ($)'] = df['Price ($)'].replace('0.00', ' ', regex=False)

# cleaning dataframe
df.drop(['Integer', 'Decimal'], axis=1, inplace=True)
df = df.iloc[:, [0, 1, 5, 2, 3, 4]]

# converting the dataframe into xlsx
df.to_excel('Rona_Scraped.xlsx', index=False)

print('Scraping is complete. Refer to the excel file "Rona_Scraped" to find your data.')
