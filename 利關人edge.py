# %%
import os
from json import load, dumps
from datetime import date

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
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


# 輸入PDF保存的路徑
PDF_savepath = r'\\10.5.20.211\資產負債管理部\權益證券科\期貨選擇權流程\利關人\券商'

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
    'savefile.default_directory': PDF_savepath
}
options.add_experimental_option('prefs', prefs)
options.add_argument('--kiosk-printing')

driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
driver.get("http://sinocloud.sph/Login.aspx")

with open('pwd.json', 'r') as f:
    pwd = load(f)

os.system(f"net use \\10.5.20.211  /user:{pwd['User']} {pwd['netpwd']}")

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
driver.get(
    'http://sinocloud.sph/EIP/SSOLauncher.aspx?SSOID=STAKEHOLDER&SSOCaption=金控利害關係人管理&Targeturl=10.240.6.7/SOFun/SSORecvMode2.aspx&TargetFrame=ContPage&TargetPage=Stakeholder/Default.aspx')
element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "divBody"))
)
driver.get('http://10.240.6.7/Stakeholder/PS/PS1100.aspx')

# %%
js = f"""document.getElementById("ctl00_ContentPlaceHolder1_qryMajorID_txtData").value="84309615" """
driver.execute_script(js)
js = f"""document.getElementById("ctl00_ContentPlaceHolder1_qryCName_txtData").value="永豐期貨股份有限公司" """
driver.execute_script(js)

js = 'document.getElementById("idWatermark").style.display = "none";'
driver.execute_script(js)

button = driver.find_element('id', 'ctl00_ContentPlaceHolder1_btnQry')
ActionChains(driver).click(button).perform()

driver.execute_script('window.print();')
os.rename(r'\\10.5.20.211\資產負債管理部\權益證券科\期貨選擇權流程\利關人\券商\金控利害關係人.pdf',
          rf'\\10.5.20.211\資產負債管理部\權益證券科\期貨選擇權流程\利關人\券商\{date.today().strftime("%Y%m%d")}_sinopac.pdf')
