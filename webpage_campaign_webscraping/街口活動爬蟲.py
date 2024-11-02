from datetime import datetime
from selenium import webdriver
from bs4 import BeautifulSoup as sp
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
import pandas as pd
import gspread


driver_dir=r'C:\Program Files (x86)\chromedriver.exe'
url='https://mkt.jkopay.com/campaign/newevent2023'

# save to google sheet
credential_dir=r'C:\Users\Vivian\Desktop\credentials.json'
sheet_key='googlesheetid'

def web_driver():
    options=webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    options.add_argument("--headless")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(driver_dir, options=options)
    return driver

def googlesheet():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(credential_dir, scopes=scopes)
    gc = gspread.authorize(credentials)
    gs = gc.open_by_key(sheet_key)
    worksheet = gs.worksheet('工作表1')
    return gs, worksheet

# web scraping campaign information
def webscrap():
    driver=web_driver()
    driver.get(url)
    soup = sp(driver.page_source, 'html.parser')
    contents = soup.find_all('div',{'class': 'sc-jFJHMl dRpxLr'})
    
    data = []
    for content in contents:
        cols = [col.text for col in content]
        data.append(cols)
        
    df = pd.DataFrame(data, columns=['daterange', 'campaign', 'description', 'hastag'])
    df['updatedate'] = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    driver.close()
    return df
    
# update new campaign to the sheet
def update_to_googlesheet():
    gs, worksheet = googlesheet()
    df_old = pd.DataFrame(worksheet.get_all_records())
    df_new = webscrap()
    
    if df_old.empty == True :
        df = df_new
    else:
        df = pd.concat([df_old, df_new], sort=False)
    
    df.drop_duplicates(['daterange', 'campaign', 'description', 'hastag'], keep='first', inplace=True)
    # print(df)
    worksheet.clear()
    set_with_dataframe(worksheet=worksheet, dataframe=df, include_index=False, include_column_header=True, resize=True)

def main():
    print('Start')
    update_to_googlesheet()
    print('Update Successfully!')

if __name__ == '__main__':
    main()