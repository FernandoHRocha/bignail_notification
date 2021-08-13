from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from openpyxl import Workbook
import openpyxl
import dados
import time
import os
import dados_cotacao
import conversor

sel_delay=0.5

def sel_enterField(self, path, text):
    field = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    field.send_keys(text)
    time.sleep(sel_delay)
def sel_enterFieldElement(element, text):
    element.send_keys(text)
    time.sleep(sel_delay)
def sel_buttonClick(self, path):
    button = WebDriverWait(self.sel_driver,5).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    button.click()
def sel_switchFrame(self, path):
    self.sel_driver.switch_to.default_content()
    frame = WebDriverWait(self.sel_driver,10).until(expected_conditions.frame_to_be_available_and_switch_to_it((By.XPATH,path)))
def sel_mouseHover(self, path):
    clickable = WebDriverWait(self.sel_driver,10).until(expected_conditions.element_to_be_clickable((By.XPATH,path)))
    hover = ActionChains(self.sel_driver).move_to_element(clickable)
    hover.perform()
def sel_getElement(self, path):
    el = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_element_located((By.XPATH,path)))
    return el
def sel_getElements(self, path):
    el = WebDriverWait(self.sel_driver,10).until(expected_conditions.presence_of_all_elements_located((By.XPATH,path)))
    return el
def sel_newWindowClick(self, path):
    action = ActionChains(self.sel_driver).key_down(Keys.SHIFT)
    action.perform()
    action = ActionChains(self.sel_driver).click(path)
    action.perform()
    action = ActionChains(self.sel_driver).key_up(Keys.SHIFT)
    action.perform()
def sel_fechar_popup(self):
    sel_mainWindow = self.sel_driver.window_handles[0]
    time.sleep(0.5)
    sel_windowToClose = self.sel_driver.window_handles[1]
    self.sel_driver.switch_to.window(sel_windowToClose)
    self.sel_driver.close()
    self.sel_driver.switch_to.window(sel_mainWindow)


class Robo:

    def iniciar(self):
        self.configurar_webdriver()
        self.coletar_credenciais_acessar_sistema()
        self.acessar_menu_comprasnet()
        #self.ler_planilha_cotacao()
        self.fechar_webdriver()
    
    def configurar_webdriver(self):
        self.options = webdriver.ChromeOptions()
        #self.options.add_argument ('--headless')
        self.options.add_argument('--log-level=3')
        self.options.add_argument('--disable-notifications')
        self.sel_driver = webdriver.Chrome("chromedriver.exe", options=self.options)
        self.sel_driver.maximize_window()
        endereco_comprasnet=dados.pregao_address
        self.sel_driver.get(endereco_comprasnet)

    def coletar_credenciais_acessar_sistema(self):
        login_comprasnet=dados.pregao_account
        senha_comprasnet=dados.pregao_pass
        sel_buttonClick(self,'//*[@id="card0"]/div/div/div/div[2]/button')
        sel_enterField(self,'//*[@id="txtLogin"]',login_comprasnet)
        sel_enterField(self,'//*[@id="txtSenha"]', senha_comprasnet)
        sel_buttonClick(self,'//*[@id="card0"]/div/div/div[2]/div[4]/button[2]')
        sel_fechar_popup(self)
        print('Logado no sistema ComprasNet')

    def acessar_menu_comprasnet(self):
        while (True):
            sel_switchFrame(self,'/html/frameset/frame[1]')
            sel_mouseHover(self,'/html/body/div[2]/div[1]')
            sel_switchFrame(self,'/html/frameset/frameset/frame')
            time.sleep(0.2)
            try:
                sel_buttonClick(self,'/html/body/div[2]/div[4]')
                break
            except:
                self.sel_driver.refresh()
                sel_fechar_popup(self)

    def ler_planilha_cotacao(self):

        return

    def fechar_webdriver(self):
        choose = input('>')
        self.sel_driver.close()
        return

start = Robo()
start.iniciar()

def register():
    print('Inserir a planilha de planejamento na pasta que abriu, com o nome COTACAO.xlsx.')
    print('Quando estiver pronto pressione ENTER para continuar.')
    path = 'C:/Fernando/LOJA/outros/twilio/bignail_notification'
    path = os.path.realpath(path)
    os.startfile(path)
    choose = input('>')
    uasg = str(dados_cotacao.import_uasg())
    numero = str(dados_cotacao.import_numero())
    sel_buttonClick('/html/body/div[1]/ul/li[1]/a')
    sel_buttonClick('/html/body/div[1]/ul/li[1]/span/a[1]')
    sel_enterField('/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input', uasg)
    sel_enterField('/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input', numero)
    sel_buttonClick('/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[7]/td/input[3]')
    itens = dados_cotacao.import_items()
    table = sel_getElements('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/form/table/tbody/tr[2]/td/table/tbody')
    
def disputar(pregao, cotados):
    
    return
    
def fight():
    itens_disputa=[]
    print('Inserir a planilha de planejamento com o nome COTACAO.xlsx.')
    print('1 - Abrir a pasta.')
    print('ENTER - Seguir sem abrir.')
    choose = input('>')
    if(choose == '1'):
        path = 'C:/Fernando/LOJA/outros/twilio/bignail_notification'
        path = os.path.realpath(path)
        os.startfile(path)
    pregao_numero = str(dados_cotacao.import_numero())
    pregao_uasg = str(dados_cotacao.import_uasg())
    print('Participar do pregão Nº '+pregao_numero+' UASG: '+pregao_uasg)
    sel_switchFrame('//*[@id="corpo"]/frame')
    sel_buttonClick('/html/body/div[1]/ul/li[2]/a')
    #sel_buttonClick('/html/body/div[1]/ul/li[2]/a')
    table = sel_getElement('/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[3]/td[2]/table/tbody')
    rows = table.find_elements_by_xpath('./*')
    infos = rows[1].find_elements_by_xpath('./*')
    if(len(infos) == 1):
        print('Nenhum pregão está em fase de disputa de lances no momento.')
    else:
        for row in rows:
            if(row != rows[0]):
                print('Temos 1 pregão em andamento do orgão: ' + infos[3].text + '\nnº ' + infos[1].text + ' UASG: ' + infos[2].text)
                print('1 - Entrar na disputa')
                print('2 - Ver o próximo')
                choose = input('>')
                if(choose == '1'):
                    itens_cotados = dados_cotacao.import_items()
                    infos[0].click()
                    sel_mainWindow = sel_driver.window_handles[0]
                    sel_fightWindow = sel_driver.window_handles[1]
                    sel_driver.switch_to.window(sel_fightWindow)
                    time.sleep(0.5)
                    try:
                        warning = sel_getElement('/html/body/modal-container')
                        sel_buttonClick('/html/body/modal-container/div/div/app-dialog-confirmacao/div/div/div[3]/div/div/button')
                        print('O pregão está encerrado e deve ser acessado pelo página de acompanhamento.')
                    except:
                        print('O pregão está em andamento.')
                    finally:
                        print('Os itens disputados estão:')
                    aguardando = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/ul/li[1]/a')
                    disputa = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/ul/li[2]/a')
                    encerrados = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/ul/li[3]')
                    print(aguardando.text + "\n" +disputa.text + "\n"+ encerrados.text)
                    if(aguardando.text != 'Aguardando disputa'):
                        aguardando.click()
                        table = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[1]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div')
                        items = table.find_elements_by_xpath('./div')
                        print(str(len(items))+' aguardando disputa: ')
                        for item in items:
                            nome = item.find_elements_by_xpath('./div')
                            #print(nome.text)
                    if(disputa.text != 'Em disputa'):
                        disputa.click()
                        #DESVINCULAR
                        table = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[2]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div')
                        items = table.find_elements_by_xpath('./div')
                        print(str(len(items))+' em disputa de lances: ')
                        for item in items:
                            item_aux = []
                            numero = item.find_elements_by_xpath('./div/div/div/div/span')[0].text
                            nome = item.find_elements_by_xpath('./div/div/div/div/span')[1].text
                            melhor_valor = item.find_element_by_xpath('./div[2]/div/div[2]/div/div/div[2]/div[1]').text
                            meu_valor = item.find_element_by_xpath('./div[2]/div/div[2]/div/div/div[2]/div[2]').text
                            tempo = item.find_element_by_xpath('./div/div[2]/div/div[2]').text
                            entrada_dados = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/input').text
                            intervalo = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[2]').text
                            melhor_valor = conversor.monetario(melhor_valor)
                            meu_valor = conversor.monetario(meu_valor)
                            tempo = conversor.temporizador(tempo)
                            intervalo = conversor.intervalo(intervalo, melhor_valor)
                            item_aux.append(numero)
                            item_aux.append(nome)
                            item_aux.append(melhor_valor)
                            item_aux.append(meu_valor)
                            item_aux.append(tempo)
                            item_aux.append(entrada_dados)
                            item_aux.append(intervalo)
                            itens_disputa.append(item_aux)
                            #disputar(itens_disputa,itens_cotados)
                        print(itens_cotados)
                        print(itens_disputa)
                    if(encerrados.text != 'Encerrados'):
                        encerrados.click()
                        table = sel_getElement('/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[3]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div')
                        items = table.find_elements_by_xpath('./div')
                        print(str(len(items))+' encerrados: ')

#function_dict={'1':register,'2':fight}



#function_dict[choose]()