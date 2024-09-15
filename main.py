import os
import pandas as pd
import time
import json
import requests
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC

url2="https://www.google.com/finance/?hl=en&authuser=5"
GF_EMAIL=os.getenv('GF_EMAIL')
GF_PW=os.getenv('GF_PW')
script_dir = os.path.dirname(os.path.abspath(__file__))
cookies_path=os.path.join(script_dir, 'gfCookies.pkl')
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_driver_path=os.path.join(script_dir, 'chromedriver.exe')

'''
Remember to install the the correct chrome drivers for this to run, you can get the installation guide from youtube ensure that you download the driver and only keep chromedriver.exe file in this directory
'''

def import_json_data(path):
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(data_dir, exist_ok=True)
    file_path = os.path.join(data_dir, path)
    print(f"Reading data from:\t {file_path}")
    with open(file_path, "r") as json_file:
        data = json.load(json_file)
    return data

def export_json_data(path, data):
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(data_dir, exist_ok=True)
    file_path = os.path.join(data_dir, path)
    print(f"Writing data at:\t {file_path}")
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=4)

def wait_and_send_keys(driver, locator, keys):
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(locator))
    element.clear()
    element.send_keys(keys)
    time.sleep(.5)

def wait_and_click(driver, locator):
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator))
    element.click()
    time.sleep(.5)

def gf_login(driver,cookies_path):
    driver.get(url2)
    signin_btn='/html/body/div[2]/header/div[2]/div[3]/div[1]/a/span'
    email_xpath='//*[@id="identifierId"]'
    email_next_xpath='//*[@id="identifierNext"]/div/button/span'
    pw_xpath='//*[@id="password"]/div[1]/div/div[1]/input'
    password_next_xpath='//*[@id="passwordNext"]/div/button/span'
    wait_and_click(driver,(By.XPATH,signin_btn))
    wait_and_send_keys(driver,(By.XPATH,email_xpath),GF_EMAIL)
    wait_and_click(driver,(By.XPATH,email_next_xpath))
    time.sleep(2)
    wait_and_send_keys(driver,(By.XPATH,pw_xpath),GF_PW)
    wait_and_click(driver,(By.XPATH,password_next_xpath))
    time.sleep(4)
    print("Execution completed :\t gf_login function...")
    
def gf_open(driver,cookies_path):
    try:
        gf_login(driver,cookies_path)
        print("Execution completed :\t gf_open function...")
        return 1
    except:
        print("Could not login")
        return 0

def create_gf_basket(driver, name):
    create_btn = '//*[@id="yDmH0d"]/c-wiz[2]/div/div[4]/div/div/div[2]/c-wiz[1]/div/div/div[5]/div/div/button/span'
    create_btn2 = '//*[@id="yDmH0d"]/c-wiz[2]/div/div[4]/div/div/div[2]/c-wiz[1]/div/div/div[2]/div/button/span'
    input_name = '/html/body/div[11]/div[2]/div/div[1]/div/div[2]/label/input'
    save_btn = '/html/body/div[11]/div[2]/div/div[2]/div/div[2]/button'
    driver.get(url2)
    
    def is_element_visible(xpath):
        try:
            element = driver.find_element(By.XPATH, xpath)
            return element.is_displayed()
        except NoSuchElementException:
            return False
    
    if is_element_visible(create_btn2):
        btn_xpath = create_btn2
    else:
        btn_xpath = create_btn
    
    wait_and_click(driver, (By.XPATH, btn_xpath))
    time.sleep(1)
    wait_and_click(driver, (By.XPATH, input_name))
    wait_and_send_keys(driver, (By.XPATH, input_name), name)
    wait_and_click(driver, (By.XPATH, save_btn))
    time.sleep(5)

def add_stocks_to_basket(driver,tickers, qtys, dates,prices=[None]):
    
    b1='#yDmH0d > c-wiz.zQTmif.SSPGKf.yd8gve > div > c-wiz > div.e1AOyf > div > div:nth-child(2) > div > div > c-wiz > div > div > div.T7rHJe > div > div.VfPpkd-dgl2Hf-ppHlrf-sM5MNb > button'

    b2='#yDmH0d div > c-wiz > div.e1AOyf > div > div:nth-child(2) > div > div > c-wiz > div > div > div > div > div.VfPpkd-dgl2Hf-ppHlrf-sM5MNb > button > span'

    search_stock='#yDmH0d > div.VfPpkd-Sx9Kwc.cC1eCc.UDxLd.PzCPDd.OHihnb.VfPpkd-Sx9Kwc-OWXEXe-FNFY6c > div.VfPpkd-wzTsW > div > div.VfPpkd-cnG4Wd > div > div:nth-child(2) > div > div > div > div.d1dlne > input.Ax4B8.ZAGvjd'

    choice_1 = '#yDmH0d > div.VfPpkd-Sx9Kwc.cC1eCc.UDxLd.PzCPDd.OHihnb.VfPpkd-Sx9Kwc-OWXEXe-FNFY6c > div.VfPpkd-wzTsW > div > div.VfPpkd-cnG4Wd > div > div:nth-child(2) > div > div > div > div:nth-child(2) > div > div > div > div:nth-child(3) > div:nth-child(1)'

    common_selector='#yDmH0d > div.VfPpkd-Sx9Kwc.cC1eCc.UDxLd.PzCPDd.OHihnb.VfPpkd-Sx9Kwc-OWXEXe-FNFY6c > div.VfPpkd-wzTsW > div > div.VfPpkd-cnG4Wd > div > div:nth-child(3) > div:nth-child(1) >'
    
    qty_btn=f'{common_selector} div > div:nth-child(1) > div > div > div > label > input'

    date_btn=f'{common_selector} div > div:nth-child(2) > div > div > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input'

    price_btn=f'{common_selector} div > div:nth-child(3) > div > label > input'

    addMore_btn=f'#yDmH0d > div.VfPpkd-Sx9Kwc.cC1eCc.UDxLd.PzCPDd.OHihnb.VfPpkd-Sx9Kwc-OWXEXe-FNFY6c > div.VfPpkd-wzTsW > div > div.VfPpkd-T0kwCb > div > div:nth-child(2) > button'
    
    save_btn=f'#yDmH0d > div.VfPpkd-Sx9Kwc.cC1eCc.UDxLd.PzCPDd.OHihnb.VfPpkd-Sx9Kwc-OWXEXe-FNFY6c > div.VfPpkd-wzTsW > div > div.VfPpkd-T0kwCb > div > div:nth-child(3) > button'

    highlight='/html/body/c-wiz[3]/div/c-wiz/div[2]/div/div[1]/div[2]/div[2]/div'

    try:
        driver.find_element(By.XPATH, highlight)
        wait_and_click(driver, (By.CSS_SELECTOR, b1))
    except NoSuchElementException:
        wait_and_click(driver, (By.CSS_SELECTOR, b2))

    crt=len(tickers)
    for ticker,qty,date,price in zip(tickers, qtys, dates,prices): 
        wait_and_send_keys(driver,(By.CSS_SELECTOR,search_stock),ticker)
        wait_and_click(driver,(By.CSS_SELECTOR,choice_1))
        wait_and_send_keys(driver,(By.CSS_SELECTOR,qty_btn),qty)
        wait_and_send_keys(driver,(By.CSS_SELECTOR,date_btn),date)
        if price is not None:
            wait_and_send_keys(driver,(By.CSS_SELECTOR,price_btn),price)
        if crt<=1:
            wait_and_click(driver,(By.CSS_SELECTOR,save_btn))         
        else :
            wait_and_click(driver,(By.CSS_SELECTOR,addMore_btn))
            time.sleep(1)
            crt-=1
    print("Execution completed :\t add_stocks_to_basket function...")

def sell_stocks_from_basket(driver,tickers, dates1, prices1, qtys, dates2, prices2):
    base_1_xpath='/html/body/c-wiz[2]/div/c-wiz/div[2]/div/div[2]/div/div/c-wiz/div/div'

    crt1=3

    base_2_xpath=f'{base_1_xpath}/div[{crt1}]/div/div/a/div/div'

    print(base_2_xpath)

def open_gf_basket(driver,bname):
    driver.get(url2)
    name='#yDmH0d > c-wiz.zQTmif.SSPGKf.ccEnac > div > div.e1AOyf > div > div > div.OFJocd > c-wiz:nth-child(1) > div > div > ul > li'

    n=len(driver.find_elements(By.CSS_SELECTOR,name))
    time.sleep(1)
    for i in range(1,n+1):
        path=f'#yDmH0d > c-wiz.zQTmif.SSPGKf.ccEnac > div > div.e1AOyf > div > div > div.OFJocd > c-wiz:nth-child(1) > div > div > ul > li:nth-child({i}) > a > div.ZS4fKc > span > div.Ej6xqf'
        e=driver.find_element(By.CSS_SELECTOR,path).text
        if(e==bname):
            wait_and_click(driver,(By.CSS_SELECTOR,path))
            time.sleep(2)
            print("Execution completed :\t open_gf_basket function...")
            return 1
    print(f"No basket with name :\t {bname}")
    driver.quit()
    print("Execution completed :\t open_gf_basket function...")
    return 0

def export_holdings_to_gf(tickers,qtys,dates,prices,driver,basketName='Holdings',external_call=False):
    if external_call is True:
        try:
            service = Service(executable_path=chrome_driver_path)
            driver = webdriver.Chrome(service=service,options=chrome_options)
        except Exception as e:
            print(e)
    gf_open(driver,cookies_path)
    create_gf_basket(driver,basketName)
    open_gf_basket(driver,basketName)
    add_stocks_to_basket(driver,tickers,qtys,dates,prices)

def load_json_file(filepath):
    """Load JSON data from a file."""
    try:
        with open(filepath, 'r') as file:
            return json.load(file)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        print(f"Error decoding JSON from file: {filepath}")
        return {}
    except Exception as e:
        print(f"An error occurred while reading: {e}")
        return {}

def save_json_file(filepath, data):
    """Save data to a JSON file."""
    try:
        with open(filepath, 'w') as file:
            json.dump(data, file, indent=4)
        print("Data updated and written to file successfully.")
    except IOError:
        print(f"Error writing to file: {filepath}")
    except Exception as e:
        print(f"An error occurred while writing: {e}")

def fetch_symbol_from_api(isin):
    """Fetch symbol from API using ISIN."""
    response = requests.get(f'https://groww.in/v1/api/search/v3/query/global/st_p_query?page=0&query={isin}&size=10&web=true')
    response_json = response.json()
    return response_json['data']['content'][0].get('nse_scrip_code', 'N/A')

def convert_date_format(date_str):
    """Convert date from 'DD-MM-YYYY' to 'MM/DD/YY'."""
    try:
        date_obj = datetime.strptime(date_str.strip(), '%d-%m-%Y')
        return date_obj.strftime('%m/%d/%y')
    except ValueError:
        return date_str

def extract_data_from_excel(filename):
    """
    Extracts data from an Excel file and returns a list of buy and sell trades.
    """
    try:
        df = pd.read_excel(filename, skiprows=2)
        tickers = df['ISIN'].tolist()
        qtys = df['Quantity'].tolist()
        bdates = df['Buy date(DD-MM-YYYY)'].tolist()
        cprices = df['Buy price'].tolist()
        sdates = df['Sell date(DD-MM-YYYY)'].tolist()
        sprices = df['Sell price'].tolist()

        map_ISIN_to_symbol = load_json_file('./ISIN_TO_SYMBOL.json')

        bdates = [convert_date_format(date) for date in bdates]
        sdates = [convert_date_format(date) if not pd.isna(date) else '' for date in sdates]

        t1, q, d1, p1 = [], [], [], []
        t2, d2, p2 = [], [], []

        for i in range(len(tickers)):
            isin = tickers[i]
            if isin not in map_ISIN_to_symbol:
                symbol = fetch_symbol_from_api(isin)
                map_ISIN_to_symbol[isin] = symbol

            t1.append(isin)
            q.append(qtys[i])
            d1.append(bdates[i])
            p1.append(cprices[i])

            if sdates[i]:
                t2.append(map_ISIN_to_symbol[isin])
                d2.append(sdates[i])
                p2.append(sprices[i])

        buy_trades = [t1, q, d1, p1]
        sell_trades = [t2, d1, p1, q, d2, p2]

        save_json_file('./ISIN_TO_SYMBOL.json', map_ISIN_to_symbol)

        return [buy_trades, sell_trades]

    except FileNotFoundError:
        print(f"Error: The file '{filename}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__=='__main__':
    service = Service(executable_path=chrome_driver_path)
    driver = webdriver.Chrome(service=service,options=chrome_options)
    a=extract_data_from_excel("data.xlsx")
    buy_trades=a[0]
    sell_trades=a[1]
    gf_open(driver,cookies_path)
    create_gf_basket(driver,"portfolio 1")
    open_gf_basket(driver,'portfolio 1')
    add_stocks_to_basket(driver,buy_trades[0],buy_trades[1],buy_trades[2],buy_trades[3])
    # sell_stocks_from_basket(driver,sell_trades[0],sell_trades[1],sell_trades[2],sell_trades[3],sell_trades[4],sell_trades[5])
    driver.quit()
