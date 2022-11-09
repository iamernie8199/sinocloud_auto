# %%
import os
import shutil
import time
from datetime import date
from glob import glob
from json import load

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.microsoft import EdgeChromiumDriverManager

if not os.path.exists('download_tmp'):
    os.makedirs('download_tmp')
savepath = os.path.join(os.getcwd(), 'download_tmp')


# method to get the downloaded file name
def getDownLoadedFile(waitTime, type='name'):
    """
    For chrome
    """
    original_window = driver.current_window_handle
    driver.execute_script("window.open()")
    # switch to new tab
    driver.switch_to.window(driver.window_handles[-1])
    # chrome下載頁
    driver.get('chrome://downloads')
    endTime = time.time() + waitTime
    while True:
        try:
            # get downloaded stat
            state = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList').items")
            state = state[0]
            # check if COMPLETE (otherwise the script will keep waiting)
            if state['state'] == 'COMPLETE':
                driver.close()
                driver.switch_to.window(original_window)
                if type == 'name':
                    # return the file name once the download is completed
                    return state['fileName']
                elif type == 'path':
                    return state['filePath']
        except:
            pass
        time.sleep(1)
        if time.time() > endTime:
            break


def latest_download_file():
    os.chdir(savepath)
    while True:
        try:
            return glob('*.csv')[-1]
        except:
            pass


def allclose():
    original_window = driver.current_window_handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            driver.close()
    driver.switch_to.window(original_window)


def checknsave(table):
    button = driver.find_element('id', table)
    ActionChains(driver).click(button).perform()

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it("Board")
    )

    td = date.today().strftime("%Y/%m/%d")
    # rp1011查當天
    if table == 'trvMenut111':
        js = f"""document.getElementById("txtBeginDate").value="{td}" """
        driver.execute_script(js)
    else:
        tmp = {
            'trvMenut21': 'txtBeginDate',
            'trvMenut29': 'txtQueryDate'
        }
        js = f"""return document.getElementById("{tmp[table]}").value"""
        if driver.execute_script(js) == td:
            tmp = {
                'trvMenut21': 'TSI庫存股數表',
                'trvMenut29': 'TSI庫存天數'
            }
            last = glob(rf'//10.5.20.211/資產負債管理部/權益證券科/權益證券科-暫定架構/四大流程/交易簿/每日交易檢核表/{tmp[table]}/T*.csv')[-1]

            if last.split('\\')[-1][-12:-4] == date.today().strftime("%Y%m%d"):
                driver.switch_to.default_content()
                driver.switch_to.frame("Menu")
                return
            else:
                tmp = last.split('\\')[-1][3:-4]
                js = f"""document.getElementById("txtBeginDate").value="{tmp[:4]}/{tmp[4:6]}/{tmp[6:]}" """
                driver.execute_script(js)

    # 查詢
    button = driver.find_element('id', 'btnInquery')
    ActionChains(driver).click(button).perform()

    # 查詢完成確認
    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
    alert.accept()

    # 下載
    button = driver.find_element('id', 'btnReserve1')
    ActionChains(driver).click(button).perform()

    # 關閉下載視窗
    allclose()

    # filepath = getDownLoadedFile(5, type='path')
    filepath = latest_download_file()
    last = filepath.split('\\')[-1]
    if table == 'trvMenut21':
        filename = f'T庫存{str(today)}.csv'
    elif table == 'trvMenut29':
        filename = f'T庫存天數{str(today)}.csv'
    elif table == 'trvMenut111':
        filename = f'利關人{str(today)}.csv'

    os.rename(filepath, filepath.replace(last, filename))
    filepath = filepath.replace(last, filename)
    last = filepath.split('\\')[-1]

    tmp = {
        'trvMenut21': 'TSI庫存股數表',
        'trvMenut29': 'TSI庫存天數',
        'trvMenut111': 'TSI利關人報表'
    }
    shutil.move(filepath, rf'\\10.5.20.211\資產負債管理部\權益證券科\權益證券科-暫定架構\四大流程\交易簿\每日交易檢核表\{tmp[table]}\{last}')

    driver.switch_to.frame("Menu")


today = date.today().strftime("%Y%m%d")
options = Options()
prefs = {
    "download.default_directory": savepath,
}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
driver.get("http://sinocloud.sph/Login.aspx")

with open('pwd.json', 'r') as f:
    pwd = load(f)

# 登入
js = f"""document.getElementById("txtUserID_txtData").value="{pwd['User']}";document.getElementById("txtPassword_txtData").value="{pwd['Password']}" """
driver.execute_script(js)

button = driver.find_element('name', 'btnLogin')
ActionChains(driver).click(button).perform()

# TSI
driver.get(
    'http://sinocloud.sph/EIP/SSOLauncher.aspx?SSOID=TSI&SSOCaption=TSI金融交易&TargetURL=10.16.1.56:898/SignOn.aspx')

WebDriverWait(driver, 10).until(
    EC.frame_to_be_available_and_switch_to_it("Menu")
)
# 展開目錄 - 台幣股票
js = "TreeView_ToggleNode(trvMenu_Data,12,document.getElementById('trvMenun12'),'t',document.getElementById('trvMenun12Nodes'));TreeView_ToggleNode(trvMenu_Data,19,document.getElementById('trvMenun19'),'t',document.getElementById('trvMenun19Nodes'))"
driver.execute_script(js)
# 展開目錄 - 通用報表
js = "TreeView_ToggleNode(trvMenu_Data,107,document.getElementById('trvMenun107'),'t',document.getElementById('trvMenun107Nodes'))"
driver.execute_script(js)

# 關閉其他視窗
allclose()
driver.switch_to.frame("Menu")

# st2002:trvMenut21/st2013:trvMenut29/rp1011:trvMenut111
for t in ['trvMenut21', 'trvMenut29', 'trvMenut111']:
    checknsave(t)
