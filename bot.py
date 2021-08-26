import sel_operacoes_comum as sel
import openpyxl
import time
import os

sel_delay=0.5
sel_driver = ''

def abrir_pasta():
    path = 'C:/Fernando/LOJA/outros/twilio/bignail_notification'
    path = os.path.realpath(path)
    os.startfile(path)

def oferecer_abrir_pasta():#OFERECE A OPÇÃO DE ABRIR A PASTA QUE DEVERÁ CONTER AS PLANILHAS CONTROLE E COTAÇÃO
    print('A planilha de cotação já está na pasta? O nome do arquivo deve ser "COTACAO.xlsx"')
    print('1 - Abrir a pasta.')
    print('2 - A planilha já está na pasta.')
    escolha = input('>')
    if(escolha == '1'):
        abrir_pasta()
        print('Aguardando para continuar.')
        escolha = input('>')

def converter_texto_para_decimal(texto):
    texto = float(texto.replace("R$","").strip().replace(".","").replace(",","."))
    return "{:.2f}".format(texto)

def converter_intervalo_lances(intervalo):
    padrao1 = 'Intervalo mínimo entre lances: R$ '
    if padrao1 in intervalo:
        intervalo = intervalo.replace(padrao1,"")
        return converter_texto_para_decimal(intervalo)

def converter_tempo_restante(tempo):
    tempo = tempo.split(":")
    return int(tempo[1])+(int(tempo[0])*60)

def enviar_lance(self, item, valor):
    valor = str(valor).replace('.',',')
    sel.enterFieldElement(item['input'],valor)
    item['botao_confirma'].click()
    sel.obter_elemento_xpath(self,'/html/body/modal-container/div/div/app-dialog-confirmacao/div/div/div[3]/div/div[2]/button').click()
    print('lance enviado')
    return

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
        oferecer_abrir_pasta()
        self.ler_planilha_cotacao(self)
        self.acessar_cadastro(self)
        self.identificar_pagina_registro(self)
        return
    
    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = str(wb.cell(2,1).value)
        self.uasg = str(wb.cell(2,2).value)
        self.abertura = str(wb.cell(2,3).value)
        self.hora = str(wb.cell(2,4).value)
        self.inserir_orgao = str(wb.cell(2,5).value)
        wb = openpyxl.load_workbook('COTACAO.xlsx', data_only=True)['Planilha1']
        itens=[]
        for row in range(2,wb.max_row):
            auxiliar_item_linha=[]
            colunas_interesse=[1,2,3,4,5]
            colunas_monetarias =[4]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    valor = str(round(wb.cell(row,col).value,2))
                    if(len(valor.split('.'))<2):
                        valor = valor + '.00'
                    elif(len(valor.split('.')[1])<2):
                        valor = valor.split('.')[0]+'.'+valor.split('.')[1]+'0'
                    auxiliar_item_linha.append(str(valor).replace('.',','))
                else:
                    auxiliar_item_linha.append(str(wb.cell(row,col).value))
            valor = str(round(round(wb.cell(row,4).value,2)*int(wb.cell(row,5).value),2))
            if(len(valor.split('.'))<2):
                    valor = valor + '.00'
            elif(len(valor.split('.')[1])<2):
                valor = valor.split('.')[0]+'.'+valor.split('.')[1]+'0'
            auxiliar_item_linha.append(str(valor).replace('.',','))
            itens.append(auxiliar_item_linha)
        self.itens_cotacao = itens
        #print(self.itens_cotacao)
        # self.itens_cotacao[0] IDENTIFICADOR
        # self.itens_cotacao[1] DESCRIÇÃO
        # self.itens_cotacao[2] MARCA
        # self.itens_cotacao[3] VALOR UNITÁRIO
        # self.itens_cotacao[4] QUANTIDADE OFERTADA
        # self.itens_cotacao[5] QUANTIDADE TOTAL

        while(True):
            print('O cadatro será realizado para o pregão: '+self.pregao+' do uasg: '+self.uasg)
            print('1 - Continuar.')
            print('2 - Abrir pasta de cotação para alterar a planilha de cotação.')
            choose = input('>')
            if(choose == '1'):
                break
            elif(choose == '2'):
                abrir_pasta()
                input('>')
                self.ler_planilha_cotacao()
                break

    def acessar_cadastro(self):
        sel.clicar_xpath(self,'/html/body/div[1]/ul/li[1]/a')
        sel.clicar_xpath(self,'/html/body/div[1]/ul/li[1]/span')
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/input',self.uasg)
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input',self.pregao)
        sel.clicar_xpath(self,'/html/body/form/table/tbody/tr[2]/td/table[2]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[7]/td/input[3]')
        sel.clicar_xpath(self,'/html/body/table/tbody/tr[2]/td/table[2]/tbody/tr[2]/td[2]/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[1]/a')

    def identificar_pagina_registro(self):
        tabela = sel.obter_elementos_xpath(self,'/html/body/center/table[2]/tbody/tr[4]/td/center[2]/table/tbody/tr')
        del tabela[0]
        item_registrar = []
        if(len(tabela)==40):#PODEM HAVER CONFIGURAÇÕES DIFERENTES

            for item_cotado in self.itens_cotacao:
                #ITENS EM COTAÇÃO EM ORDEM CRESCENTE DE IDENTIFICAÇÃO DO ITEM
                #PARA CADA ITEM EM COTAÇÃO, SERÁ BUSCADO O ITEM NA PÁGINA DE RESGISTRO E MUDANDO DE PÁGINA ATÉ ENCONTRAR O ITEM E REMOVE-LO DA LISTA
                item_pendente = True
                while(item_pendente):
                    tabela = sel.obter_elementos_xpath(self,'/html/body/center/table[2]/tbody/tr[4]/td/center[2]/table/tbody/tr')
                    del tabela[0]
                    itens = []
                    while(True):#ARRANJA OS CAMPOS DOS ITENS EM APENAS UM ELEMENTO DENTRO DA LISTA
                        item_aux = []
                        item_aux.append(tabela.pop(0))
                        item_aux.append(tabela.pop(0))
                        item_aux.append(tabela.pop(0))
                        item_aux.append(tabela.pop(0))
                        itens.append(item_aux)
                        if(len(tabela)<1):
                            break

                    item_registrar = []
                    for item_registrar in itens:#PROCURA ITEM A ITEM
                        identificador = item_registrar[0].find_element_by_css_selector('.tex3b').text#IDENTIFICAÇÃO DO ITEM
                        if(identificador == item_cotado[0]):
                            item_pendente = False
                            self.preencher_item_registrar(item_cotado = item_cotado, item_registrar = item_registrar)
                    if(item_pendente):#CASO O ITEM NÃO TENHA SIDO ENCONTRADO, PROSSEGUIMOS PARA A PRÓXIMA PÁGINA
                        sel.obter_elemento_id(self,'proximas').click()

    def preencher_item_registrar(item_registrar,item_cotado):
        for entrada in item_registrar[0].find_elements_by_tag_name('input'):#ENTRADAS DE DADOS REFERENTES A QUANTIDADE E VALORES
            if(entrada.is_displayed() and entrada.is_enabled()):
                if(str(entrada.get_attribute('name')) == 'qtdOfertada'):
                    sel.enterFieldElement(entrada,item_cotado[4])
                if(str(entrada.get_attribute('name')) == 'valorunit'):
                    sel.enterFieldElement(entrada,item_cotado[3])
                if(str(entrada.get_attribute('name')) == 'valorprp'):
                    sel.enterFieldElement(entrada,item_cotado[5])
                #print(entrada.get_attribute('name'))
                #qtdOfertada
                #valorunit
                #valorprp
        for entrada in item_registrar[1].find_elements_by_tag_name('input'):#ENTRADAS DE DADOS REFERENTES A MARCA MODELO
            if(entrada.is_displayed() and entrada.is_enabled()):
                if(str(entrada.get_attribute('name')) == 'MarcaFornec'):
                    sel.enterFieldElement(entrada,item_cotado[2])
                if(str(entrada.get_attribute('name')) == 'FabriFornec'):
                    sel.enterFieldElement(entrada,item_cotado[2])
                if(str(entrada.get_attribute('name')) == 'ModVerFornec'):
                    sel.enterFieldElement(entrada,item_cotado[2])
                #print(entrada.get_attribute('name'))
                #MarcaFornec
                #FabriFornec
                #ModVerFornec
        for entrada in item_registrar[2].find_elements_by_tag_name('textarea'):#ENTRADAS DE DADOS REFERENTES A DESCRIÇÃO
            if(entrada.is_displayed() and entrada.is_enabled()):
                if(str(entrada.get_attribute('name')) == 'DescrFornec'):
                    sel.enterFieldElement(entrada,item_cotado[1])
                #print(entrada.get_attribute('name'))
                #DescrFornec


class Disputar:#DISPUTA OS PREÇOS DO PREGÃO REFERENTE AO ARQUIVO DE COTAÇÃO

    def iniciar(self):
        self.sel_driver = sel_driver
        oferecer_abrir_pasta()
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
                if((codpregao == self.pregao) and (coduasg == self.uasg)):#SOMENTE ABRIRÁ A DISPUTA CASO CORRESPONDA AO PREGÃO E AO UASG
                    colunas[0].click()
            time.sleep(5)
            self.janela_pregoes = self.sel_driver.window_handles[0]
            self.janela_disputa = self.sel_driver.window_handles[1]
            self.sel_driver.switch_to.window(self.janela_disputa)
            return True
        else:
            print('Não foram encontradas disputas em andamento.')
            return False
        
    def reconhecer_disputa(self):#COLOCAR UMA ESTRUTURA DE REPETIÇÃO PARA CONTINUAR ATÉ QUE A DISPUTA ENCERRE
        self.modo_disputa = sel.obter_elemento_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[4]/div[1]/app-identificacao-compra/div/span').text
        self.navegacao_itens = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/ul/li')
        
        while(True):
            if(self.navegacao_itens[0].text != 'Aguardando' or self.navegacao_itens[1].text != 'Em disputa'):
                if(self.navegacao_itens[1].text != 'Em disputa'):
                    self.navegacao_itens[1].click()
                    self.reconhecer_itens_disputa(self)#COMEÇAR O CICLO DE DISPUTA DE LANCES
                elif(self.navegacao_itens[0].text != 'Aguardando'):
                    aguardando = True
                    while(aguardando):
                        if(self.navegacao_itens[1].text != 'Em disputa'):
                            aguardando=False
                        else:
                            time.sleep(5)
            else:
                break
        

    def reconhecer_itens_disputa(self):
        while(True):
            itens_em_disputa = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[2]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div/div')
            itens_reconhecidos = []
            for item in itens_em_disputa:
                item_disputa = {}
                codigo_item = str(item.find_element_by_xpath('./div[1]/div[1]/div[1]/div[1]/span[1]').text)
                atual_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[1]').text)
                nosso_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[2]').text)
                tempo_restante = str(item.find_element_by_xpath('./div[1]/div[2]/div/div[2]/span/span').text)
                intervalo_lances = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/span/small').text)
                input = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/input')
                botao_confirma = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/div/button/u')
                menu_lances = item.find_element_by_xpath('./div[2]/div[2]/div/app-botao-icone/span/button/i')
                item_disputa = {
                    'item':codigo_item,
                    'atual_valor':atual_valor,
                    'nosso_valor':nosso_valor,
                    'tempo_restante':tempo_restante,
                    'intervalo_lances':converter_intervalo_lances(intervalo_lances),
                    'input':input,
                    'botao_confirma':botao_confirma,
                    'menu_lances':menu_lances,
                }
                itens_reconhecidos.append(item_disputa)
            self.itens_reconhecidos = itens_reconhecidos

    def decidir_lance(self,item):#adicionar modo de disputa fechada
        for cotado in self.itens:
            if(str(cotado[0]) == str(item['item'])):
                melhor = converter_texto_para_decimal(item['melhor_valor'])
                nosso = converter_texto_para_decimal(item['nosso_valor'])
                tempo = converter_tempo_restante(item['tempo_restante'])
                intervalo = converter_intervalo_lances(item['intervalo_lances'])
                print(melhor)
                print(nosso)
                print(tempo)
                if(tempo < 120):
                    if( melhor > nosso):
                        valor = melhor - intervalo
                        #enviar_lance(self,item,valor)
                    print('Nosso mínimo é R$ '+ nosso)
                    print('Preço atual é R$ ' + melhor)
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
#bot.ler_planilha_cotacao()

def fechar_webdriver(self):
    choose = input('>')
    self.sel_driver.close()

fechar_webdriver(bot)