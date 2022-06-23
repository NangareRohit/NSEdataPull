import json
import os
from selenium import webdriver
import pandas as pd
import requests
import xlwings as xw
headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'

}

def get_session_cookies():
    plat_exe = "chromedriver"
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    plat_exe_path = os.path.join(PROJECT_ROOT, plat_exe)
    print(plat_exe_path)
    driver = webdriver.Chrome(executable_path=plat_exe_path)
    driver.get("https://www.nseindia.com")
    cookies = driver.get_cookies()
    cookie_dic = {}
    with open(os.path.join(os.path.dirname(__file__), "cookies"), "w") as line:
        for cookie in cookies:
            cookie_dic[cookie['name']] = cookie['value']
        line.write(json.dumps(cookie_dic))
    driver.quit()
    return cookie_dic

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    underlying_sp = "HDFC"
    url = 'https://www.nseindia.com/api/option-chain-equities?symbol={0}'.format(underlying_sp)
    cookies = get_session_cookies()
    session = requests.session()
    for cookie in cookies:
        if cookie in ['bm_sv', 'nseappid', 'nsit']:
            session.cookies.set(cookie, cookies[cookie])

    r = session.get(url, headers=headers, timeout=10).json()
    ce_data = pd.DataFrame([data['CE'] for data in r['filtered']['data'] if "CE" in data])
    sheet["A1"].value = ce_data




@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("nsedatapull.xlsm").set_mock_caller()
    main()
