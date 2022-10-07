# %%
import sys
import os
import shutil
from json import load, dumps
from datetime import date
from glob import glob

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service

from webdriver_manager.microsoft import EdgeChromiumDriverManager

def latest_download_file():
    os.chdir(savepath)
    while True:
        try:
            return glob('*.XLS')[-1]
        except:
            pass

def allclose():
    original_window = driver.current_window_handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            driver.close()
    driver.switch_to.window(original_window)

if __name__ == '__main__':

    last = glob(r'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易檢核表\TCRI\20*.XLS')[-1]
    last = last.split('\\')[-1]

    if last.replace('.XLS', '') == date.today().strftime("%Y%m%d"):
        sys.exit()

    if not os.path.exists('download_tmp'):
        os.makedirs('download_tmp')
    savepath = os.path.join(os.getcwd(), 'download_tmp')
    PDF_savepath = r'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易明細\券商利關人(暫放)'

    options = Options()

    appState = {
        'recentDestinations': [{
            'id': 'Save as PDF',
            'origin': 'local',
            'account': ''
        }],
        'selectedDestinationId': 'Save as PDF',
        'version': 2
    }
    prefs = {
        'printing.print_preview_sticky_settings.appState': dumps(appState), 
        'savefile.default_directory': PDF_savepath,
        "download.default_directory": savepath,
    }
    options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing')
    options.add_argument("disable-web-security")
    options.add_argument("disable-notifications")

    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options = options)
    driver.get("http://sinocloud.sph/Login.aspx")

    with open('pwd.json', 'r') as f:
        pwd = load(f)

    # 登入
    js = f"""document.getElementById("txtUserID_txtData").value="{pwd['User']}" """
    driver.execute_script(js)
    js = f"""document.getElementById("txtPassword_txtData").value="{pwd['Password']}" """
    driver.execute_script(js)

    button = driver.find_element('name', 'btnLogin')

    ActionChains(driver).click(button).perform()
    print('log in')

    allclose()

    # %%
    driver.get("http://sinocloud.sph/EIP/ProgFrame.aspx?ProgID=EIP00101")
    driver.switch_to.frame("frmMenu")

    button = driver.find_element("id",'ProgTreet1')

    ActionChains(driver).click(button).perform()

    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "home_13"))
    )
    driver.get('http://tej.bsp/html/f1/f1_default.htm') 
    # %%
    driver.get('http://tej.bsp/html/f2/f2.asp')
    driver.switch_to.frame("main")
    driver.switch_to.frame("option")
    button = driver.find_element('name', 'save')
    ActionChains(driver).click(button).perform()

    filepath = latest_download_file()
    os.rename(filepath, f'{date.today().strftime("%Y%m%d")}.XLS')
    filepath = f'{date.today().strftime("%Y%m%d")}.XLS'
    shutil.move(filepath, rf'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易檢核表\TCRI\{filepath}')
    os.chdir("..")
    shutil.rmtree(savepath)	