from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import dados
import time

sel_delay = 0.2
sel_wait = 0.2

def enterField(self, path, text):
    field = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    field.send_keys(text)
    time.sleep(sel_delay)

def enterFieldElement(element, text):
    element.send_keys(text)
    time.sleep(sel_delay)

def clicar_xpath(self, path):
    button = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    button.click()

def trocar_frame(self, path):
    self.sel_driver.switch_to.default_content()
    frame = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.frame_to_be_available_and_switch_to_it((By.XPATH,path)))

def sobrepor_mouse(self, path):
    clickable = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    hover = ActionChains(self.sel_driver).move_to_element(clickable)
    hover.perform()

def obter_elemento_xpath(self, path):
    el = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    return el

def obter_elemento_id(self, id):
    el = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.presence_of_element_located((By.ID,id)))
    return el

def obter_elementos_xpath(self, path):
    el = WebDriverWait(self.sel_driver,sel_wait).until(expected_conditions.presence_of_all_elements_located((By.XPATH,path)))
    return el

def clicar_nova_janela_xpath(self, path):
    elemento = self.sel_driver.find_element_by_xpath(path)
    href = elemento.get_attribute('href')
    if(href == None):
        href = elemento.find_element_by_xpath('./a').get_attribute('href')
    self.sel_driver.execute_script("window.open('"+href+"');")

def clicar_nova_janela_elemento(self, elemento):
    href = elemento.get_attribute('href')
    if(href == None):
        href = elemento.find_element_by_xpath('./a').get_attribute('href')
    self.sel_driver.execute_script("window.open('"+href+"');")

def fechar_popup(self):
    self.sel_mainWindow = self.sel_driver.window_handles[0]
    time.sleep(sel_delay)
    sel_windowToClose = self.sel_driver.window_handles[1]
    self.sel_driver.switch_to.window(sel_windowToClose)
    self.sel_driver.close()
    self.sel_driver.switch_to.window(self.sel_mainWindow)

def configurar_webdriver(self):
    print('Brainpro Tecnologia - Automação por Fernando H Rocha')
    options = webdriver.ChromeOptions()
    options.add_argument ('--headless')
    options.add_argument('--log-level=3')
    options.add_argument('--disable-notifications')
    self.sel_driver = webdriver.Chrome("chromedriver.exe", options=options)
    #self.sel_driver.maximize_window()
    endereco_comprasnet=dados.pregao_address
    self.sel_driver.get(endereco_comprasnet)

def coletar_credenciais_acessar_sistema(self):
    login_comprasnet=dados.pregao_account
    senha_comprasnet=dados.pregao_pass
    clicar_xpath(self,'//*[@id="card0"]/div/div/div/div[2]/button')
    enterField(self,'//*[@id="txtLogin"]',login_comprasnet)
    enterField(self,'//*[@id="txtSenha"]', senha_comprasnet)
    clicar_xpath(self,'//*[@id="card0"]/div/div/div[2]/div[4]/button[2]')
    fechar_popup(self)
    print('Logado no sistema ComprasNet')

def acessar_menu_comprasnet(self):
    while (True):
        trocar_frame(self,'/html/frameset/frame[1]')
        sobrepor_mouse(self,'/html/body/div[2]/div[1]')
        trocar_frame(self,'/html/frameset/frameset/frame')
        time.sleep(sel_delay)
        try:
            clicar_xpath(self,'/html/body/div[2]/div[4]')
            break
        except:
            self.sel_driver.refresh()
            time.sleep(sel_delay)
            fechar_popup(self)