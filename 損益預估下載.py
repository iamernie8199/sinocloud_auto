import os
import shutil
import pandas as pd
from datetime import date
from json import load, dumps

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service

from webdriver_manager.microsoft import EdgeChromiumDriverManager


def allclose():
    original_window = driver.current_window_handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            driver.close()
    driver.switch_to.window(original_window)


def checknsave(table):
    btn = driver.find_element('id', table)
    ActionChains(driver).click(btn).perform()

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("Board"))

    # 查詢
    btn = driver.find_element('id', 'btnInquery')
    ActionChains(driver).click(btn).perform()

    # 查詢完成確認
    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
    alert.accept()

    # 下載
    btn = driver.find_element('id', 'btnReserve1')
    ActionChains(driver).click(btn).perform()

    # 關閉下載視窗
    allclose()

    driver.switch_to.frame("Menu")


with open('pwd.json', 'r') as f:
    pwd = load(f)

# 建立下個交易日資料夾
parent = rf'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\即時監控 XQ'
tomorrow = date.today() + pd.offsets.BDay(1)
if not os.path.exists(os.path.join(parent, tomorrow.strftime("%Y"), tomorrow.strftime("%Y%m%d"))):
    os.makedirs(os.path.join(parent, tomorrow.strftime("%Y"), tomorrow.strftime("%Y%m%d")))
savepath = os.path.join(parent, tomorrow.strftime("%Y"), tomorrow.strftime("%Y%m%d"))

# 複製底稿
shutil.copyfile(os.path.join(parent, date.today().strftime("%Y"), date.today().strftime("%Y%m%d"),
                             f'交易簿即時損益監控{date.today().strftime("%Y%m%d")}.xlsx'),
                os.path.join(savepath, f'交易簿即時損益監控{tomorrow.strftime("%Y%m%d")}.xlsx'))

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
    "download.default_directory": savepath,
}
options.add_experimental_option('prefs', prefs)
options.add_argument('--kiosk-printing')

driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
driver.get("http://sinocloud.sph/Login.aspx")

# 登入
js = f"""document.getElementById("txtUserID_txtData").value="{pwd['User']}"; document.getElementById("txtPassword_txtData").value="{pwd['Password']}" """
driver.execute_script(js)
button = driver.find_element('name', 'btnLogin')
ActionChains(driver).click(button).perform()

allclose()

# TSI
driver.get(
    'http://sinocloud.sph/EIP/SSOLauncher.aspx?SSOID=TSI&SSOCaption=TSI金融交易&TargetURL=10.16.1.56:898/SignOn.aspx')

WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("Menu"))
# 展開目錄 - 台幣股票
js = "TreeView_ToggleNode(trvMenu_Data,12,document.getElementById('trvMenun12'),'t',document.getElementById('trvMenun12Nodes'))"
driver.execute_script(js)
js = "TreeView_ToggleNode(trvMenu_Data,19,document.getElementById('trvMenun19'),'t',document.getElementById('trvMenun19Nodes'))"
driver.execute_script(js)

"""
st2002: 'trvMenut21'
st2004:'trvMenut23'
st2009: 'trvMenut26'
st2013: 'trvMenut29'
st2014; 'trvMenut30'
"""
for i in ['trvMenut21', 'trvMenut23', 'trvMenut26', 'trvMenut29', 'trvMenut30']:
    checknsave(i)
