from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
from pushsafer import init, Client
import os
import platform
from win32com.client import Dispatch
import urllib.request
from pathlib import Path
from zipfile import ZipFile
from win10toast import ToastNotifier
# download driver from https://chromedriver.chromium.org/downloads
# then copy the driver to the folder that you run this script
# install selenium by typing pip install selenium in console
# install pushsafer by typing pip install python-pushsafer in console
def clear():
    if platform.system() == 'Linux':
        os.system('clear')
        return 'chromedriver_mac64.zip'
    elif platform.system() == 'Darwin':
        os.system('clear')
        return 'chromedriver_linux64.zip'
    else:
        os.system('cls')
        return 'chromedriver_win32.zip'

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

os_name = clear()
if __name__ == "__main__":
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    # print(version)

if int(version[0:2]) == 89:
    url = 'https://chromedriver.storage.googleapis.com/89.0.4389.23/' + str(os_name)
elif int(version[0:2]) == 88:
    url = 'https://chromedriver.storage.googleapis.com/88.0.4324.96/' + str(os_name)
elif int(version[0:2]) == 90:
    url = 'https://chromedriver.storage.googleapis.com/90.0.4430.24/' + str(os_name)
else:
    print('This script needs at least Chrome 88, please update your Google Chrome.')

if not os.path.isfile(str(Path.cwd() / 'chromedriver.exe')):
    urllib.request.urlretrieve(url, filename='zipped.zip')
    with ZipFile('zipped.zip', 'r') as zipObj:
       # Extract all the contents of zip file in current directory
       zipObj.extractall()
else:
    print('You already have the chromedriver')

if platform.system() == 'Linux' or platform.system() == 'Darwin':
    options = webdriver.ChromeOptions()
    options.binary_location = input('Enter Chrome application binary (usually in Applications/Google Chrome.app/Contents/MacOS/Google Chrome): ')
    chrome_driver_binary = input('Enter Chromedriver binary location: ')
    driver = webdriver.Chrome(chrome_driver_binary, chrome_options=options)
else:
    driver = webdriver.Chrome()
##
plzList = str(input('Enter postcodes of vaccination centers to search: \nPlease use comma to separate: e.g. 12345, 67890\t: ')).replace(' ','').split(',')
age = str(input('Please enter your age: '))

##
cookie_count = 0
toast = ToastNotifier()
ind = 0

while True:
    # driver.get('https://impfterminradar.de/') #
    clear()
    plz = plzList[ind]
    ind += 1
    if ind == len(plzList)-1:
        ind = 0
    print('Checking',plz)
    driver.get('https://001-iz.impfterminservice.de/impftermine/service?plz='+plz)
    driver.refresh()
    # driver.minimize_window()
    time.sleep(1)
    try:
        para = driver.find_element_by_xpath('/html/body/section/div[2]/div/p').text
    except:
        para = ''
    if 'Nachfrage ' in para:
        print('Waiting room in',plz)
        pass
    else:
        try:
            wartung = driver.find_element_by_xpath('/html/body/div/div[2]/h1').text
        except:
            wartung = ''
            pass
        if 'Wartungsarbeiten' in wartung:
            pass
        else:
            try:
                button = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[2]/div/div/label[2]/span')
                ActionChains(driver).click(button).perform()
            except:
                continue
            if cookie_count == 0:
                try:
                    cookies = driver.find_element_by_xpath('/html/body/app-root/div/div/div/div[2]/div[2]/div/div[2]/a')
                    ActionChains(driver).click(cookies).perform()
                    cookie_count += 1
                except:
                    continue
            time.sleep(5)
            try:
                check_quota = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/div/div').text
            except:
                check_quota = ''
                pass
            if 'keine' in check_quota:
                print('No quota in', plz, 'Trying next one')
            else:
                try: ## /html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/app-corona-vaccination-no/form/div[3]/div/div
                    button2 = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/app-corona-vaccination-no/form/div[1]/div/div/label[1]/span')
                    ActionChains(driver).click(button2).perform()
                    time.sleep(0.3)
                    age_inp = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/app-corona-vaccination-no/form/div[3]/div/div/input')
                    age_inp.send_keys(age)
                    prufung = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/app-corona-vaccination-no/form/div[4]/button')
                    ActionChains(driver).click(prufung).perform()
                    time.sleep(0.3)
                    check2 = driver.find_element_by_xpath('/html/body/app-root/div/app-page-its-login/div/div/div[2]/app-its-login-user/div/div/app-corona-vaccination/div[3]/div/div/div/div[2]/div/app-corona-vaccination-no/form/div[2]').text
                    if 'keine freien Termine' in check2:
                        continue
                    else:
                        toast.show_toast("Quota available","In Vaccine center "+plz+" quota available",duration=5)
                        time.sleep(6000)
                except:
                    time.sleep(6000)
                    continue
