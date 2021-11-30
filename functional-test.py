"""
1.  Selenium: abrir o moodle
2.  Selenium: ir até a página da matéria
3.  Selenium: for para clicar em todos os cards
4.  BSP:      buscar no elemento img-text se existe um elemento cmoppt ou pdf 
5.  Selenium: Se existir clicar no elemento e baixar o arquivo

"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
import time

driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://moodle.pucrs.br/')
time.sleep(3)
inputElement = driver.find_element(By.ID, "login_username")
inputElement.send_keys(str(input('Entre com seu usuário: ')))
inputElementPass = driver.find_element(By.ID,"login_password")
inputElementPass.send_keys(str(input('Entre com sua senha: ')))
inputElement.send_keys(Keys.ENTER)
time.sleep(5)
        