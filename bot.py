import sel_operacoes_comum as sel
import conversor
import openpyxl
import time
import os

sel_delay=0.5
sel_driver = ''

def abrir_pasta_cotacao():
    print('A planilha de cotação já está na pasta? O nome do arquivo deve ser "COTACAO.xlsx"')
    print('1 - Abrir a pasta.')
    print('2 - A planilha já está na pasta.')
    escolha = input('>')
    if(escolha == '1'):
        path = 'C:/Fernando/LOJA/outros/twilio/bignail_notification'
        path = os.path.realpath(path)
        os.startfile(path)

class ComprasNet:#LEVA A APLICAÇÃO ATÉ UM LUGAR EM COMUM DENTRO DO COMPRASNET E MOSTRA AS OPÇÕES DE OPERAÇÕES

    def iniciar(self):
        sel.configurar_webdriver(self)
        sel.coletar_credenciais_acessar_sistema(self)
        sel.acessar_menu_comprasnet(self)
        self.oferecer_opcoes()

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
        abrir_pasta_cotacao()
        self.ler_planilha_cotacao(self)
        self.acessar_cadastro(self)
        self.registrar_proposta(self)
        return
    
    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = wb.cell(2,1).value
        self.uasg = wb.cell(2,2).value
        self.abertura = wb.cell(2,3).value
        self.hora = wb.cell(2,4).value
        self.inserir_orgao = wb.cell(2,5).value
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Planilha1']
        itens=[]
        for row in range(2,wb.max_row):
            rowItens=[]
            colunas_interesse=[1,2,3,4,5]
            colunas_monetarias =[4]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    rowItens.append(round(wb.cell(row,col).value,2))
                else:
                    rowItens.append(wb.cell(row,col).value)
            itens.append(rowItens)
        self.itens = itens

    def acessar_cadastro(self):
        sel.clicar_xpath(self,'/html/body/div[1]/ul/li[1]/a')
        sel.clicar_xpath(self,'/html/body/div[1]/ul/li[1]/span')
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input',self.uasg)
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input',self.pregao)
        sel.clicar_xpath(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[7]/td/input[3]')
        sel.clicar_xpath(self,'/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]/a')

    def registrar_proposta(self):
        tabela = sel.obter_elementos_xpath(self,'/html/body/center/table[2]/tbody/tr[4]/td/center[2]/table/tbody/tr')
        del tabela[0]
        print(len(tabela))
        item_registrar = []
        if(len(tabela)==40):
            for n in range(0,len(tabela)):
                item_auxiliar = []
                #item = tabela[n].find_element_by_class('tex3b')
                #print(item)
        return

class Disputar:#DISPUTA OS PREÇOS DO PREGÃO REFERENTE AO ARQUIVO DE COTAÇÃO

    def iniciar(self):
        self.sel_driver = sel_driver
        abrir_pasta_cotacao()
        self.ler_planilha_cotacao(self)
        if(self.abrir_disputa(self)):
            self.reconhecer_disputa(self)
        #self.extrair_relatorio(self)
        return

    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = str(wb.cell(2,1).value)
        self.uasg = str(wb.cell(2,2).value)
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Planilha1']
        self.itens=[]
        for row in range(2,wb.max_row):
            rowItens=[]
            colunas_interesse=[1,4,5,10]
            colunas_monetarias =[4,10]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    rowItens.append(round(wb.cell(row,col).value,2))
                else:
                    rowItens.append(wb.cell(row,col).value)
            self.itens.append(rowItens)
        print(self.itens)     

    def abrir_disputa(self):
        sel.clicar_xpath(self,'/html/body/div[1]/ul/li[2]/a')
        tabela = sel.obter_elementos_xpath(self,'/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[3]/td[2]/table/tbody/tr')
        del tabela[0]
        if(len(tabela[0].find_elements_by_xpath('./*'))>1):
            for linha in tabela:
                colunas = linha.find_elements_by_xpath('./td')
                codpregao = colunas[1].text
                coduasg = colunas[2].text
                if((codpregao == self.pregao) and (coduasg == self.uasg)):
                    colunas[0].click()
            time.sleep(5)
            self.janela_pregoes = self.sel_driver.window_handles[0]
            self.janela_disputa = self.sel_driver.window_handles[1]
            self.sel_driver.switch_to.window(self.janela_disputa)
            return True
        else:
            print('Não foram encontradas disputas em andamento.')
            return False
        
    def reconhecer_disputa(self):
        self.modo_disputa = sel.obter_elemento_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[4]/div[1]/app-identificacao-compra/div/span').text
        self.navegacao_itens = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/ul/li')
        for botao in self.navegacao_itens:
            print(botao.text)
        if(self.navegacao_itens[1].text != 'Em disputa'):
            self.navegacao_itens[1].click()
            itens_em_disputa = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[2]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div/div')
            print(len(itens_em_disputa))
            for item in itens_em_disputa:
                codigo_item = str(item.find_element_by_xpath('./div[1]/div[1]/div[1]/div[1]/span[1]').text)
                melhor_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[1]').text)
                nosso_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[2]').text)
                print(codigo_item)
                print(melhor_valor)
                print(nosso_valor)
                #CONSEGUE LER DURANTE A DISPUTA ABERTA E DISPURA FECHADA NÃO CONVOCADO
            return
        return
    
    def extrair_relatorio(self):
        self.navegacao_itens[2].click()
        itens_encerrados = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[3]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div/div')
        self.resultado_por_item = []
        for item in itens_encerrados:
            aux_item = []
            aux_item.append(str(item.find_elemet_by_xpath('./div[1]/div[1]/div[1]/div').text))
            
        return

bot = ComprasNet()
bot.iniciar()

#bot = Disputar()
#bot.ler_planilha_cotacao()

#bot = Registrar()
#bot.iniciar()
exit()
    
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