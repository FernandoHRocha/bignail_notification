from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import time

sel_delay = 0.2

def enterField(self, path, text):
    field = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    field.send_keys(text)
    time.sleep(sel_delay)

def enterFieldElement(element, text):
    element.send_keys(text)
    time.sleep(sel_delay)

def buttonClick(self, path):
    button = WebDriverWait(self.sel_driver,1).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    button.click()

def switchFrame(self, path):
    self.sel_driver.switch_to.default_content()
    frame = WebDriverWait(self.sel_driver,10).until(expected_conditions.frame_to_be_available_and_switch_to_it((By.XPATH,path)))

def mouseHover(self, path):
    clickable = WebDriverWait(self.sel_driver,10).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    hover = ActionChains(self.sel_driver).move_to_element(clickable)
    hover.perform()

def getElement(self, path):
    el = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    return el

def getElements(self, path):
    el = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_all_elements_located((By.XPATH,path)))
    return el

def newWindowClick(self, path):
    action = ActionChains(self.sel_driver).key_down(Keys.SHIFT)
    action.perform()
    action = ActionChains(self.sel_driver).click(path)
    action.perform()
    action = ActionChains(self.sel_driver).key_up(Keys.SHIFT)
    action.perform()

def fechar_popup(self):
    sel_mainWindow = self.sel_driver.window_handles[0]
    time.sleep(0.5)
    sel_windowToClose = self.sel_driver.window_handles[1]
    self.sel_driver.switch_to.window(sel_windowToClose)
    self.sel_driver.close()
    self.sel_driver.switch_to.window(sel_mainWindow)
