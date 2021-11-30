"""
1.  Selenium: abrir o moodle
2.  Selenium: ir até a página da matéria
3.  Selenium: for para clicar em todos os cards
4.  BSP:buscar no elemento img-text se existe um elemento pdf 
5.  Selenium: Se existir clicar no elemento e baixar o arquivo
"""
import subprocess
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException
from bs4 import BeautifulSoup
import time

chrome_version = subprocess.check_output(
    r'wmic datafile where name="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" get Version /value', shell=True)
print(chrome_version.decode('utf-8').strip())

driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://moodle.pucrs.br/')
time.sleep(3)

#1.  Selenium: abrir o moodle
inputElement = driver.find_element(By.ID, "login_username")
inputElement.send_keys(str(input('Entre com seu usuário: ')))
inputElementPass = driver.find_element(By.ID, "login_password")
inputElementPass.send_keys(str(input('Entre com sua senha: ')))
inputElement.send_keys(Keys.ENTER)
time.sleep(5)

#2.  Selenium: ir até a página da matéria
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
i = 0
for elem in soup.find_all('span', {'class': 'multiline'}):
    if (elem.text).rstrip().lstrip()[0].isdigit(): #pegar somente cards que iniciam em numeral: disciplinas
        driver.find_elements(By.CLASS_NAME, 'multiline')[1].click()
        time.sleep(5)

        # Close pop-up
        close = driver.find_elements(By.CSS_SELECTOR, '#page-course-view-tiles > div:nth-child(20) > div.modal.moodle-has-zindex.show > div > div > div.modal-footer > button.btn.btn-secondary')
        if len(close) > 0:
            close[0].click()
            time.sleep(4)

        # 3.  Selenium: for para clicar em todos os cards
        cards = driver.find_elements(By.CLASS_NAME, 'tile-clickable')
        for j in range(len(cards)):
            # open card
            cards[j+1].click()
            time.sleep(4)

            #4.  BSP:      buscar no elemento img-text se existe um elemento pdf 
            card_html = driver.page_source
            card_soup = BeautifulSoup(card_html, 'html.parser')
            pdfs = driver.find_elements(By.CLASS_NAME, 'activityicon')
            img_src = card_soup.find_all('img', {'class': 'activityicon'})
            for k in range(len(img_src)):
                if 'pdf' in img_src[k]['src']:
                    try:
                        pdfs[k].click()
                        time.sleep(3)
                        if len(driver.find_elements(By.CLASS_NAME, 'modal')) > 0:
                            modal_html = driver.page_source
                            modal_soup = BeautifulSoup(modal_html, 'html.parser')
                            close_btn = modal_soup.find_all('a', {'class': 'button_download'})
                            driver.find_element(By.ID, close_btn[-1]['id']).click()
                            time.sleep(3)
                            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            time.sleep(3)
                    except (ElementClickInterceptedException, ElementNotInteractableException):
                        pass
            
            # close card
            card_close = card_soup.find_all('span', {'class': 'closesectionbtn'})
            for l in range(len(card_close)):
                try:
                    driver.find_element(By.ID, card_close[l]['id']).click()
                except (ElementClickInterceptedException, ElementNotInteractableException):
                    pass
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(5)

        # home
        driver.find_element(By.CLASS_NAME, 'has-logo').click()
        time.sleep(5)
    i += 1
    break
