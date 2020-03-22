import os
import time
import numpy as np
import pandas as pd
from datetime import date
from datetime import datetime as dt
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Firefox(service_log_path=os.path.devnull)
driver.maximize_window()

driver.get('http://finra-markets.morningstar.com/BondCenter/Results.jsp')

# click agree
WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, ".button_blue.agree"))).click()

# click edit search
WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'a.qs-ui-btn.blue'))).click()

# click advanced search
WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'a.ms-display-switcher.hide'))).click()

# select bond ratings
WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.CSS_SELECTOR, 'select.range[name=moodysRating]')))
Select((driver.find_elements_by_css_selector(
    'select.range[name=moodysRating]'))[0]).select_by_visible_text('A3')
Select((driver.find_elements_by_css_selector(
    'select.range[name=moodysRating]'))[1]).select_by_visible_text('Aaa')
Select((driver.find_elements_by_css_selector(
    'select.range[name=standardAndPoorsRating]'))[0]).select_by_visible_text('A-')
Select((driver.find_elements_by_css_selector(
    'select.range[name=standardAndPoorsRating]'))[1]).select_by_visible_text('AAA')

# click show results
WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'input.button_blue[type=submit]'))).click()

# wait for results
WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.CSS_SELECTOR, '.rtq-grid-row.rtq-grid-rzrow .rtq-grid-cell-ctn')))
headers = [title.text for title in driver.find_elements_by_css_selector(
    '.rtq-grid-row.rtq-grid-rzrow .rtq-grid-cell-ctn')[1:]]


fig, (ax1, ax2) = plt.subplots(1, 2, clear=True)

# create dataframe from scrape
bonds = []
for page in range(1, 11):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, (f"a.qs-pageutil-btn.on[value='{str(page)}']"))))  # wait for page marker to be on expected page
    time.sleep(2)
    tablerows = driver.find_elements_by_css_selector(
        'div.rtq-grid-bd > div.rtq-grid-row')
    for tablerow in tablerows:
        tablerowdata = tablerow.find_elements_by_css_selector(
            'div.rtq-grid-cell')
        bond = [item.text for item in tablerowdata[1:]]
        print(bond)
        bonds.append(bond)

        # drop bonds with missing yields
        df = pd.DataFrame(bonds, columns=headers)
        df['Yield'].replace('', np.nan, inplace=True)
        df = df.dropna(subset=['Yield'])

        df['Yield'] = df['Yield'].astype(float)
        now = dt.strptime(date.today().strftime('%m/%d/%Y'), '%m/%d/%Y')
        df['Maturity'] = pd.to_datetime(df['Maturity']).dt.strftime('%m/%d/%Y')
        daystillmaturity = []
        for maturity in df['Maturity']:
            daystillmaturity.append(
                (dt.strptime(maturity, '%m/%d/%Y') - now).days)
        df = df.reset_index(drop=True)
        df['Maturity'] = pd.Series(daystillmaturity)

        # get rid of ridiculous yields
        df = df[np.abs(df.Yield - df.Yield.mean()) <= (3 * df.Yield.std())]

        Mgroups = df.groupby("Moody's®")
        ax1.clear()
        ax1.margins(0.05)
        ax1.set_xlabel('Days Until Maturity')
        ax1.set_ylabel('Yield')
        ax1.set_title("Moody's® Ratings")
        for name, group in Mgroups:
            ax1.plot(group['Maturity'], group['Yield'],
                     marker='o', linestyle='', ms=12, label=name)
        ax1.legend(numpoints=1, loc='upper left')

        SPgroups = df.groupby("S&P")
        ax2.clear()
        ax2.margins(0.05)
        ax2.set_xlabel('Days Until Maturity')
        ax2.set_ylabel('Yield')
        ax2.set_title("S&P Ratings")

        for name, group in SPgroups:
            ax2.plot(group['Maturity'], group['Yield'],
                     marker='o', linestyle='', ms=12, label=name)
        ax2.legend(numpoints=1, loc='upper left')
        plt.pause(.001)

    print('\npage completed...\n')
    driver.find_element_by_css_selector('a.qs-pageutil-next').click()
    print(df)

df.to_excel('data.xlsx')
os.startfile('data.xlsx')
plt.show()