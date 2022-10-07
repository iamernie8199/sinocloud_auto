import pywinauto
import os
import shutil
from glob import glob
from json import load, dumps

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service

from webdriver_manager.microsoft import EdgeChromiumDriverManager


def latest_download_file():
    os.chdir(savepath)
    while True:
        try:
            return glob('*.pdf')[-1]
        except:
            pass

def allclose():
    original_window = driver.current_window_handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            driver.close()
    driver.switch_to.window(original_window)


def upload(path):
    button = driver.find_element('id', 'ctl00_ContentPlaceHolder1_ucUploadButton2_btnLaunch')
    ActionChains(driver).click(button).perform()

    original_window = driver.current_window_handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)

    button = driver.find_element('id', 'uploader_browse')
    ActionChains(driver).click(button).perform()

    app = pywinauto.Desktop()
    dlg = app["開啟"]
    dlg["檔案名稱(&N):):Edit"].SetText(path)
    dlg["開啟(&O)"].click()

    button = driver.find_element('class name', 'plupload_start')
    ActionChains(driver).click(button).perform()

    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "Util_txtDone"))
    )

    driver.close()
    driver.switch_to.window(original_window)

    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "alertify-log-success"))
    )

    button = driver.find_element('id', 'ctl00_ContentPlaceHolder1_btnQry')
    ActionChains(driver).click(button).perform()

    button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ucGridView1_gvMain_ctl01_btnExportPdf"))
    )
    ActionChains(driver).click(button).perform()

    filepath = latest_download_file()

    os.rename(filepath, path.split('\\')[-1].replace('txt', 'pdf'))
    filepath = path.split('\\')[-1].replace('txt', 'pdf')
    shutil.move(filepath, rf'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易明細\方向性名單\{filepath}')
    os.chdir("..")
    shutil.rmtree(savepath)	


with open('pwd.json', 'r') as f:
    pwd = load(f)

PDF_savepath = r'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易明細\券商利關人(暫放)'
if not os.path.exists('download_tmp'):
    os.makedirs('download_tmp')
savepath = os.path.join(os.getcwd(), 'download_tmp')

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

driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options = options)
driver.get("http://sinocloud.sph/Login.aspx")

# 登入
js = f"""document.getElementById("txtUserID_txtData").value="{pwd['User']}"; document.getElementById("txtPassword_txtData").value="{pwd['Password']}" """
driver.execute_script(js)

button = driver.find_element('name', 'btnLogin')

ActionChains(driver).click(button).perform()
print('log in')

allclose()

driver.get('http://sinocloud.sph/EIP/SSOLauncher.aspx?SSOID=STAKEHOLDER&SSOCaption=金控利害關係人管理&Targeturl=10.240.6.7/SOFun/SSORecvMode2.aspx&TargetFrame=ContPage&TargetPage=Stakeholder/Default.aspx')
element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "divBody"))
)
driver.get('http://10.240.6.7/Stakeholder/PS/PS1200.aspx')

js = 'document.getElementById("idWatermark").style.display = "none";'
driver.execute_script(js)

select = Select(driver.find_element('id', 'ctl00_ContentPlaceHolder1_qryRespectively_ddlSourceList'))
select.select_by_value("1")

upload(r'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易明細\方向性名單\量化批次查詢-名稱.txt')

select = Select(driver.find_element('id', 'ctl00_ContentPlaceHolder1_qryRespectively_ddlSourceList'))
select.select_by_value("0")

upload(r'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易明細\方向性名單\量化批次查詢-證號.txt')
