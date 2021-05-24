
from sys import warnoptions
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from twilio.rest import Client
import dados
import time
import xlrd
from xlutils.copy import copy
import xlwt
import os

#CONFIGURAÇÃO DE VARIAVEIS
#TWILIO
twilio_account = dados.twilio_account
twilio_pass = dados.twilio_pass
twilio_from_number=dados.twilio_from_number
twilio_to_number=dados.twilio_to_number
#PREGAO
pregao_account=dados.pregao_account
pregao_pass=dados.pregao_pass
pregao_address=dados.pregao_address
#pregao_sheet=xlrd.open_workbook('notificacoes.xlsx')
#SELENIUM
sel_driver = webdriver.Chrome("chromedriver.exe")
sel_driver.maximize_window()
sel_driver.get(pregao_address)
sel_delay=0.2
#client = Client(twilio_account, twilio_pass)
#message = client.messages.create(body='Hello there!', from_=twilio_from_number, to=twilio_to_number)
#print(message.sid)

def sel_enterField(path, text):
    field = WebDriverWait(sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    field.send_keys(text)
    time.sleep(sel_delay)
def sel_enterFieldElement(element, text):
    element.send_keys(text)
    time.sleep(sel_delay)
def sel_buttonClick(path):
    button = WebDriverWait(sel_driver,10).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    button.click()
def sel_switchFrame(path):
    sel_driver.switch_to.default_content()
    frame = WebDriverWait(sel_driver,10).until(expected_conditions.frame_to_be_available_and_switch_to_it((By.XPATH,path)))
def sel_mouseHover(path):
    clickable = WebDriverWait(sel_driver,10).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    hover = ActionChains(sel_driver).move_to_element(clickable)
    hover.perform()
def sel_getElement(path):
    el = WebDriverWait(sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    return el
def sel_getElements(path):
    el = WebDriverWait(sel_driver,10).until(expected_conditions.presence_of_all_elements_located((By.XPATH,path)))
    return el

#ACESSAR O SISTEMA
sel_buttonClick('//*[@id="card0"]/div/div/div/div[2]/button')
sel_enterField('//*[@id="txtLogin"]',pregao_account)
sel_enterField('//*[@id="txtSenha"]', pregao_pass)
sel_buttonClick('//*[@id="card0"]/div/div/div[2]/div[4]/button[2]')
print('LOGADO COM SUCESSO')
#FECHAR POPUP
sel_mainWindow = sel_driver.window_handles[0]
time.sleep(0.5)
sel_windowToClose = sel_driver.window_handles[1]
sel_driver.switch_to.window(sel_windowToClose)
sel_driver.close()
sel_driver.switch_to.window(sel_mainWindow)
#frame1 = sel_getElement('//*[@id="corpo"]/frame')
#ACESSAR SEÇÃO DE PREGÕES EM ANDAMENTO
sel_switchFrame('/html/frameset/frame[1]')
sel_mouseHover('/html/body/div[2]/div[1]')
sel_switchFrame('//*[@id="corpo"]/frame')
time.sleep(0.2)
sel_buttonClick('/html/body/div[1]/div[4]')
sel_buttonClick('/html/body/div[1]/ul/li[4]/a')

sel_switchFrame('/html/frameset/frameset/frame')
table = sel_getElements('/html/body/div[1]/table/tbody/tr[5]/td[2]/table/tbody/*')
link = []
nail = []
for index in range(0,len(table)):
    if( index == 0 ):
        print('cabeçalho')
    else:
        linha = table[index].find_elements_by_xpath('./*')
        link.append(linha[0])
        aux_nail=[]
        aux_nail.append(linha[1].text)
        aux_nail.append(linha[2].text)
        aux_nail.append(linha[3].text)
        nail.append(aux_nail)

def sel_lerPregao(link, sheet):
    link.click()
    sel_buttonClick('/html/body/div[1]/table/tbody/tr[6]/td[2]/input')
    sel_mainWindow = sel_driver.window_handles[0]
    sel_windowToClose = sel_driver.window_handles[1]
    sel_driver.switch_to.window(sel_windowToClose)
    table = sel_getElements('/html/body/table[2]/tbody/*')
    sh = copy(excel_read.get_sheet(sheet))
    for index in range(0,len(table)):
        if(index<20):
            msg = table[index].find_elements_by_xpath('./*')
            sh.write(index, 0, msg[0].text)
            sh.write(index, 1, msg[1].text)
            print('linha')
    sh.save('notifa.xls')
    print('arquivo salvo ')
    sel_driver.close()
    sel_driver.switch_to.window(sel_mainWindow)
    sel_buttonClick('/html/body/div[1]/table/tbody/tr[4]/td[2]/input[2]')

excel_read = xlrd.open_workbook('notifa.xls')
excel_write = copy(excel_read)
for index in range(0,len(nail)):
    criar = True
    sheet_name=str(nail[index][0]+nail[index][1])
    for sheet in range(0,len(excel_read.sheets())):
        if(excel_read.sheet_by_index(sheet).name == sheet_name):
            criar = False
            if(link[index].text == 'Acompanhar'):
                sel_lerPregao(link[index], sheet)
    if(criar):
        excel_write.add_sheet(sheet_name)
        excel_write.save('notifa.xls')

excel_write.save('notifa.xls')