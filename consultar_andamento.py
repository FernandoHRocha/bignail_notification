from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import sel_operacoes_comum as sel
import openpyxl
import dados
import time
import sys
import os
from openpyxl import Workbook

excel_read=''

class andamento():
    def iniciar(self):
        self.pregao_planilha()
        print("Varredura iniciada.")
        sel.configurar_webdriver(self)
        sel.coletar_credenciais_acessar_sistema(self)
        sel.acessar_menu_comprasnet(self)
        self.verificar_anexos()

    def pregao_planilha(self):
        pregao_path = dados.pregao_path
        try:
            self.excel_read = openpyxl.load_workbook(pregao_path)
        except:
            print("A planilha de controle de mensagens foi corrompida.")
            if os.path.exists(pregao_path):
                os.remove(pregao_path)
            wb = openpyxl.Workbook().save(pregao_path)
            self.excel_read = openpyxl.load_workbook(pregao_path)
            print("Planilha de controle recriada.")
    
    def verificar_anexos(self):
        sel.newWindowClick(self,'/html/body/div[1]/ul/li[11]/a')
        time.sleep(0.2)
        sel_windowToClose = sel_driver.window_handles[1]
        self.sel_driver.switch_to.window(sel_windowToClose)
        table = sel.getElements('/html/body/div[1]/form/table/tbody/tr[5]/td[2]/table/tbody/*')
        del table[0]
        if(table[0].text == "No momento não existem pregões para enviar anexos."):
            print('Sem envio de anexos pendente')
        else:
            for rows in table:
                print('\nATENÇÃO É PRECISO ENVIAR ANEXOS PARA: ')
                info = rows.find_elements_by_xpath('./*')
                print("Pregão: "+info[1].text +"\nUASG: "+ info[2].text + "\nÓrgão: "+ info[3].text)
                print('\n')
        self.sel_driver.close()
        self.sel_driver.switch_to.window(self.sel_mainWindow)
        print('certo')
    
    def varrer_pregoes(self):
        #sel_switchFrame('/html/frameset/frameset/frame')
        #sel_buttonClick('/html/body/div[1]/ul/li[4]/a')

        return

bot = andamento()
bot.iniciar()
exit()


print("Varredura iniciada.")

def pregao_planilha():
    try:
        excel_read = openpyxl.load_workbook(pregao_path)
    except:
        print("A planilha de controle de mensagens foi corrompida.")
        if os.path.exists(pregao_path):
            os.remove(pregao_path)
        wb = openpyxl.Workbook().save(pregao_path)
        excel_read = openpyxl.load_workbook(pregao_path)
    return excel_read

pregao_path = dados.pregao_path
excel_read = pregao_planilha()

#PREGAO
identidade = dados.identidade
pregao_account=dados.pregao_account
pregao_pass=dados.pregao_pass
pregao_address=dados.pregao_address
link = []
nail = []
#SELENIUM
options = webdriver.ChromeOptions()
options.add_argument ('--headless')
options.add_argument('--log-level=3')
options.add_argument('--disable-notifications')
sel_driver = webdriver.Chrome("chromedriver.exe", options=options)
sel_driver.maximize_window()
sel_driver.get(pregao_address)
sel_delay=0.2

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
def sel_newWindowClick(path):
    action = ActionChains(sel_driver).key_down(Keys.SHIFT)
    action.perform()
    action = ActionChains(sel_driver).click(path)
    action.perform()
    action = ActionChains(sel_driver).key_up(Keys.SHIFT)
    action.perform()
def sel_consultarAnexo():
    table = sel_getElements('/html/body/div[1]/form/table/tbody/tr[5]/td[2]/table/tbody/*')
    del table[0]
    if(table[0].text == "No momento não existem pregões para enviar anexos."):
        print('Sem envio de anexos pendente')
    else:
        for rows in table:
            print('\nATENÇÃO É PRECISO ENVIAR ANEXOS PARA: ')
            info = rows.find_elements_by_xpath('./*')
            print("Pregão: "+info[1].text +"\nUASG: "+ info[2].text + "\nÓrgão: "+ info[3].text)
            print('\n')
def sel_refreshTable():
    sel_switchFrame('/html/frameset/frameset/frame')
    table = sel_getElements('/html/body/div[1]/table/tbody/tr[5]/td[2]/table/tbody/*')
    del table[0]
    def_nail = []
    for index in table:
        linha = index.find_elements_by_xpath('./*')
        aux_nail=[]
        aux_nail.append(linha[1].text)
        aux_nail.append(linha[2].text)
        aux_nail.append(linha[3].text)
        aux_nail.append(linha[0])
        def_nail.append(aux_nail)
    return def_nail

def sel_lerPregao(linka, sheet, exist):
    sel_mainWindow = sel_driver.window_handles[0]
    sel_newWindowClick(linka)
    time.sleep(1)
    sel_windowToClose = sel_driver.window_handles[1]
    sel_driver.switch_to.window(sel_windowToClose)
    sel_buttonClick('/html/body/div[1]/table/tbody/tr[6]/td[2]/input')
    sel_driver.close()
    sel_windowToClose = sel_driver.window_handles[1]
    sel_driver.switch_to.window(sel_windowToClose)
    table = sel_getElements('/html/body/table[2]/tbody/*')
    sh = excel_read[sheet]
    if(exist):
        msg = table[0].find_elements_by_xpath('./*')
        if(sh.cell(row=1, column=1).value == msg[0].text):
            print('_________________________ Processo: '+sheet+"\nSem novas atividades.\n")
        else:
            for index in range(0,len(table)):
                if(index<50):
                    msg = table[index].find_elements_by_xpath('./*')
                    cl = sh.cell(row=index+1, column=1)
                    cl.value = msg[0].text
                    cl = sh.cell(row=index+1, column=2)
                    cl.value = msg[1].text
                    if identidade in msg[1].text:
                        print('_________________________ Processo: '+sheet)
                        print('Mensagem referenciando a sua indentidade foi enviada.\n'+msg[0].text+' expressando:\n'+msg[1].text)
                if(index==0):
                    print('_________________________ Processo: '+sheet+" "+ msg[0].text + "\n" + msg[1].text+"\n")
    else:
        for index in range(0,len(table)):
            if(index<50):
                msg = table[index].find_elements_by_xpath('./*')
                cl = sh.cell(row=index+1, column=1)
                cl.value = msg[0].text
                cl = sh.cell(row=index+1, column=2)
                cl.value = msg[1].text
                if identidade in msg[1].text:
                    print('_________________________ Processo: '+sheet)
                    print('Mensagem referenciando a sua indentidade foi enviada.\n'+msg[0].text+' expressando:\n'+msg[1].text)
            if(index==0):
                    print('_________________________ Processo: '+sheet+" "+ msg[0].text + "\n" + msg[1].text+"\n")
    excel_read.save(pregao_path)
    sel_driver.close()
    sel_driver.switch_to.window(sel_mainWindow)

#ACESSAR O SISTEMA
sel_buttonClick('//*[@id="card0"]/div/div/div/div[2]/button')
sel_enterField('//*[@id="txtLogin"]',pregao_account)
sel_enterField('//*[@id="txtSenha"]', pregao_pass)
sel_buttonClick('//*[@id="card0"]/div/div/div[2]/div[4]/button[2]')
print('Logado no sistema ComprasNet')
#FECHAR POPUP
sel_mainWindow = sel_driver.window_handles[0]
time.sleep(0.5)
sel_windowToClose = sel_driver.window_handles[1]
sel_driver.switch_to.window(sel_windowToClose)
sel_driver.close()
sel_driver.switch_to.window(sel_mainWindow)
#ACESSAR SEÇÃO DE PREGÕES EM ANDAMENTO
sel_switchFrame('/html/frameset/frame[1]')
sel_mouseHover('/html/body/div[2]/div[1]')
time.sleep(0.2)
sel_switchFrame('//*[@id="corpo"]/frame')
time.sleep(0.2)
try:
    sel_buttonClick('/html/body/div[1]/div[4]')
except:
    sel_switchFrame('/html/frameset/frame[1]')
    sel_mouseHover('/html/body/div[2]/div[1]')
    sel_switchFrame('//*[@id="corpo"]/frame')
    sel_buttonClick('/html/body/div[1]/div[4]')
time.sleep(0.2)
#CONSULTAR ANEXOS PENDENTES
sel_mainWindow = sel_driver.window_handles[0]
sel_newWindowClick(sel_getElement('/html/body/div[1]/ul/li[11]/a'))
time.sleep(0.2)
sel_windowToClose = sel_driver.window_handles[1]
sel_driver.switch_to.window(sel_windowToClose)
sel_consultarAnexo()
sel_driver.close()
sel_driver.switch_to.window(sel_mainWindow)
sel_switchFrame('/html/frameset/frameset/frame')
sel_buttonClick('/html/body/div[1]/ul/li[4]/a')

nail = sel_refreshTable()
print('No momento '+ str(len(nail)) +' pregões estão ativos.\n')

#COMPARAR INFORMAÇÕES COM A PLANILHA
for index in range(0,len(nail)):
    sel_refreshTable()
    criar = True
    sheet_name=str(nail[index][0]+" | "+nail[index][1])
    for sheet in range(0,len(excel_read.sheetnames)):
        if(excel_read.worksheets[sheet].title == sheet_name):
            criar = False
            if(nail[index][3].text == 'Acompanhar'):
                try:
                    sel_lerPregao(nail[index][3], sheet_name, True)
                except:
                    print('Ops, aconteceu o erro ', sys.exc_info()[0])
                    print('Infelizmente não foi possível realizar a consulta ao pregão: nº ' + nail[index][0]+" | UASG: " + nail[index][1]+"\n")
    if(criar):
        excel_read.create_sheet(sheet_name)
        excel_read.save(pregao_path)
        if(nail[index][3].text == 'Acompanhar'):
            for sheet in range(0,len(excel_read.sheetnames)):
                if(excel_read.worksheets[sheet].title == sheet_name):
                    sel_lerPregao(nail[index][3], sheet_name, False)

print('Varredura finalizada.')
sel_driver.quit()
excel_read.save(pregao_path)
excel_read.close()
inp = input('>')
sys.exit()