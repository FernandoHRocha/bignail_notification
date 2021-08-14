from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from openpyxl import Workbook
import sel_operacoes_comum as sel
import conversor
import openpyxl
import dados
import time
import os

sel_delay=0.5
sel_driver = ''

class ComprasNet:#LEVA A APLICAÇÃO ATÉ UM LUGAR EM COMUM DENTRO DO COMPRASNET E MOSTRA AS OPÇÕES DE OPERAÇÕES

    def iniciar(self):
        self.configurar_webdriver()
        self.coletar_credenciais_acessar_sistema()
        self.acessar_menu_comprasnet()
        self.oferecer_opcoes()
    
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
        sel.buttonClick(self,'//*[@id="card0"]/div/div/div/div[2]/button')
        sel.enterField(self,'//*[@id="txtLogin"]',login_comprasnet)
        sel.enterField(self,'//*[@id="txtSenha"]', senha_comprasnet)
        sel.buttonClick(self,'//*[@id="card0"]/div/div/div[2]/div[4]/button[2]')
        sel.fechar_popup(self)
        print('Logado no sistema ComprasNet')

    def acessar_menu_comprasnet(self):
        while (True):
            sel.switchFrame(self,'/html/frameset/frame[1]')
            sel.mouseHover(self,'/html/body/div[2]/div[1]')
            sel.switchFrame(self,'/html/frameset/frameset/frame')
            time.sleep(0.2)
            try:
                sel.buttonClick(self,'/html/body/div[2]/div[4]')
                break
            except:
                self.sel_driver.refresh()
                sel.fechar_popup(self)

    def oferecer_opcoes(self):
        class_dict={'1':Registrar,'2':Disputar}
        print('Escolha o modo de operação')
        print('1 - Registrar proposta.')
        print('2 - Participar da disputa de lances')
        escolha = input('>')
        global sel_driver
        sel_driver = self.sel_driver
        class_dict[escolha].iniciar(class_dict[escolha])

class Registrar:#REGISTRA O PREGÃO REFERENTE AO ARQUIVO DE COTAÇÃO

    def iniciar(self):
        self.sel_driver = sel_driver
        self.ler_planilha_cotacao(self)
        self.acessar_cadastro(self)
        return
    
    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = wb.cell(2,1).value
        self.uasg = wb.cell(2,2).value
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Planilha1']
        itens=[]
        for row in range(2,wb.max_row):
            rowItens=[]
            colunas_interesse=[1,2,3,4,9,13]
            colunas_monetarias =[3,9]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    rowItens.append(round(wb.cell(row,col).value,2))
                else:
                    rowItens.append(wb.cell(row,col).value)
            itens.append(rowItens)

    def acessar_cadastro(self):
        sel.buttonClick(self,'/html/body/div[1]/ul/li[1]/a')
        sel.buttonClick(self,'/html/body/div[1]/ul/li[1]/span')
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input',self.uasg)
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input',self.pregao)
        sel.buttonClick(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[7]/td/input[3]')
        sel.buttonClick(self,'/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]/a')

    def registrar_proposta(self):
        return

class Disputar:#DISPUTA OS PREÇOS DO PREGÃO REFERENTE AO ARQUIVO DE COTAÇÃO

    def iniciar(self):
        self.sel_driver = sel_driver
        self.pasta_cotacao(self)
        self.ler_planilha_cotacao(self)
        return

    def pasta_cotacao(self):
        print('A planilha de cotação já está na pasta? O nome do arquivo deve ser "COTACAO.xlsx"')
        print('1 - Abrir a pasta.')
        print('2 - A planilha já está na pasta.')
        escolha = input('>')
        if(escolha == '1'):
            path = 'C:/Fernando/LOJA/outros/twilio/bignail_notification'
            path = os.path.realpath(path)
            os.startfile(path)

    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = wb.cell(2,1).value
        self.uasg = wb.cell(2,2).value
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Planilha1']
        itens=[]
        for row in range(2,wb.max_row):
            rowItens=[]
            colunas_interesse=[1,2,3,4,9]
            colunas_monetarias =[3,9]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    rowItens.append(round(wb.cell(row,col).value,2))
                else:
                    rowItens.append(wb.cell(row,col).value)
            itens.append(rowItens)
        print(itens)
        

    def disputar_lances(self):
        return

start = ComprasNet()
start.iniciar()

#start = Disputar()
#start.ler_planilha_cotacao()

#start = Registrar()
#start.ler_planilha_cotacao()
    
def disputa():
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


def fechar_webdriver(self):
    choose = input('>')
    self.sel_driver.close()