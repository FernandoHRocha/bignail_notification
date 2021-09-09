import sel_operacoes_comum as sel
import openpyxl
import time
import os
import sys
import traceback

sel_delay=0.5
sel_driver = ''
caminho_pasta = 'C:/Fernando/LOJA/outros/twilio/bignail_notification/COTACAO'
arquivo_planilha = 'COTACAO.xlsx'
arquivo_proposta = 'PROPOSTA.pdf'
arquivo_documentacao = 'DOCUMENTACAO.rar'
endereco_planilha = caminho_pasta + '/' + arquivo_planilha
endereco_proposta = caminho_pasta + '/' + arquivo_proposta
endereco_documentacao = caminho_pasta + '/' + arquivo_documentacao
itens_aguardando_disputa = 'Aguardando disputa'
itens_fase_disputa = 'Em disputa'
itens_disputa_encerrados = 'Encerrados'
item_etapa_aberta = 'Etapa aberta'
item_etapa_fechada = 'Etapa fechada'

def abrir_pasta():
    path = os.path.realpath(caminho_pasta)
    os.startfile(path)

def oferecer_abrir_pasta():#OFERECE A OPÇÃO DE ABRIR A PASTA QUE DEVERÁ CONTER AS PLANILHAS CONTROLE E COTAÇÃO
    print('O nome dos arquivos devem ser:\nPlanilha de cotação "'+arquivo_planilha+'"')
    print('Proposta assinada "'+arquivo_proposta+'"')
    print('Documentação compactada "'+arquivo_documentacao+'"')
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

def converter_intervalo_minimo(intervalo):
    padrao1 = 'Intervalo mínimo entre lances: R$ '
    padrao2 = 'Não há intervalo mínimo entre lances'
    if padrao1 in intervalo:
        intervalo = intervalo.replace(padrao1,"")
        return converter_texto_para_decimal(intervalo)
    elif padrao2 in intervalo:
        return 0.10

def converter_tempo_restante(tempo):
    tempo = tempo.split(":")
    return int(tempo[1])+(int(tempo[0])*60)

class ComprasNet:#LEVA A APLICAÇÃO ATÉ UM LUGAR EM COMUM DENTRO DO COMPRASNET E MOSTRA AS OPÇÕES DE OPERAÇÕES

    def iniciar(self):
        sel.configurar_webdriver(self,True)
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
        self.submeter_docs = True
        if(len(self.itens_cotacao)>0):
            print('iniciando processo de registro de item')
            while(True):
                print('Temos ',len(self.itens_cotacao),' item(s) para registrar.')
                self.encontrar_proximo_item_cotado(self,self.itens_cotacao[0],self.identificar_pagina_registro(self))
                if(len(self.itens_cotacao)==0):
                    break
        else:
            print('A planilha não possui itens cotados, ou não foi possível a sua identificação.')
        self.aceitar_insercao_item(self,False)
    
    def ler_planilha_cotacao(self):
        wb = openpyxl.load_workbook('./COTACAO/COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = str(wb.cell(2,1).value)
        self.uasg = str(wb.cell(2,2).value)
        self.abertura = str(wb.cell(2,3).value)
        self.hora = str(wb.cell(2,4).value)
        self.inserir_orgao = str(wb.cell(2,5).value)
        wb = openpyxl.load_workbook('./COTACAO/COTACAO.xlsx', data_only=True)['Planilha1']
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
        sel.trocar_frame(self,'/html/frameset/frameset/frame')
        tabela = sel.obter_elementos_xpath(self,'/html/body/center/table[2]/tbody/tr[4]/td/center[2]/table/tbody/tr')
        del tabela[0]
        item_registrar = []
        novo_grupo = True
        itens_pagina = []
        aux_item =[]
        for tr in tabela:
            if(novo_grupo):
                if(tr.find_element_by_xpath('./td[2]').text != ''):
                    aux_item.append(tr)
                    novo_grupo = False
                else:
                    aux_item.append(tr)
                    novo_grupo = False
            else:
                if(tr.find_element_by_xpath('./td[2]').text == ''):
                    aux_item.append(tr)
                else:
                    novo_grupo = True
                    itens_pagina.append(aux_item)
                    aux_item = []
                    aux_item.append(tr)
        itens_pagina.append(aux_item)
        return itens_pagina

    def encontrar_proximo_item_cotado(self, item_cotado, itens_pagina):
        print('Procurando inserir o item ',item_cotado[0])
        for item_pagina in itens_pagina:
            sel.trocar_frame(self,'/html/frameset/frameset/frame')
            if(item_pagina[0].find_element_by_xpath('./td[2]').text == item_cotado[0]):
                self.preencher_item_registrar(self,item_cotado = self.itens_cotacao.pop(0), item_pagina = item_pagina)
                if(self.submeter_docs):
                    self.submeter_documentacao(self)
                    self.submeter_docs = False
                break
            if(item_pagina == itens_pagina[(len(itens_pagina)-1)]):
                self.aceitar_insercao_item(self,True)
    
    def submeter_documentacao(self):
        sel.clicar_xpath(self,'/html/body/center/table[2]/tbody/tr[3]/td/table/tbody/tr[4]/td/table/tbody/tr/td/center/input[1]')
        tabela = sel.obter_elementos_xpath(self,'//table[@id="declaraEdital"]//following-sibling::table')
        for linha in tabela:
            linha.find_element_by_xpath('./tbody/tr/td[2]/input').click()
        #DOCUMENTACAO
        janela_atual = self.sel_driver.window_handles[0]
        sel.obter_elemento_id(self,'incluiAnexo').click()
        time.sleep(sel_delay)
        janela_documento = self.sel_driver.window_handles[1]
        self.sel_driver.switch_to.window(janela_documento)
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input',endereco_documentacao)
        sel.clicar_xpath(self,'/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input')
        sel.aceitar_alerta(self.sel_driver)
        self.sel_driver.switch_to.window(janela_atual)
        sel.trocar_frame(self,'/html/frameset/frameset/frame')
        while(True):
            time.sleep(sel_delay)
            if(len(self.sel_driver.window_handles) == 1):
                break
        #PROPOSTA
        sel.obter_elemento_id(self,'incluiAnexoP').click()
        time.sleep(sel_delay)
        janela_documento = self.sel_driver.window_handles[1]
        self.sel_driver.switch_to.window(janela_documento)
        sel.enterField(self,'/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/input',endereco_proposta)
        sel.clicar_xpath(self,'/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td/input')
        sel.aceitar_alerta(self.sel_driver)
        while(True):
            time.sleep(sel_delay)
            if(len(self.sel_driver.window_handles) == 1):
                break
        self.sel_driver.switch_to.window(janela_atual)
        sel.trocar_frame(self,'/html/frameset/frameset/frame') 
        time.sleep(1)
        sel.enter_alerta(self.sel_driver)

    def preencher_item_registrar(self,item_pagina,item_cotado):
        for linha in item_pagina:
            for entrada in linha.find_elements_by_tag_name('input'):#ENTRADAS DE DADOS REFERENTES A QUANTIDADE, VALORES, MODELO
                if(entrada.is_displayed() and entrada.is_enabled()):
                    if(str(entrada.get_attribute('name')) == 'qtdOfertada'):
                        sel.enterFieldElement(entrada,item_cotado[4])
                    if(str(entrada.get_attribute('name')) == 'valorunit'):
                        sel.enterFieldElement(entrada,item_cotado[3])
                    if(str(entrada.get_attribute('name')) == 'valorprp'):
                        sel.enterFieldElement(entrada,item_cotado[5])
                    if(str(entrada.get_attribute('name')) == 'MarcaFornec'):
                        sel.enterFieldElement(entrada,item_cotado[2])
                    if(str(entrada.get_attribute('name')) == 'FabriFornec'):
                        sel.enterFieldElement(entrada,item_cotado[2])
                    if(str(entrada.get_attribute('name')) == 'ModVerFornec'):
                        sel.enterFieldElement(entrada,item_cotado[2])
            for entrada in linha.find_elements_by_tag_name('textarea'):#ENTRADAS DE DADOS REFERENTES A DESCRIÇÃO
                if(entrada.is_displayed() and entrada.is_enabled()):
                    if(str(entrada.get_attribute('name')) == 'DescrFornec'):
                        sel.enterFieldElement(entrada,item_cotado[1])
        sel.obter_elemento_id(self,'declaraEdital').find_element_by_xpath('./tbody/tr/td[2]/input').click()

    def aceitar_insercao_item(self, continuar):
        sel.obter_elemento_id(self,'incluir').click()
        time.sleep(1)
        sel.aceitar_alerta(self.sel_driver)
        if(continuar):
            try:
                sel.aceitar_alerta(self.sel_driver)
            except:
                time.sleep(2)
            sel.trocar_frame(self,'/html/frameset/frameset/frame')
            sel.obter_elemento_id(self,'proximas').click()
        else:
            print('Seus itens foram inseridos com sucesso.')

class Disputar:#DISPUTA OS PREÇOS DO PREGÃO REFERENTE AO ARQUIVO DE COTAÇÃO

    def iniciar(self):
        self.sel_driver = sel_driver
        oferecer_abrir_pasta()
        self.ler_planilha_cotacao(self)
        if(self.abrir_disputa(self)):
            self.reconhecer_disputa(self)
        return

    def ler_planilha_cotacao(self):#OBTEM UMA LISTA DE VALORES NUMÉRICOS PARA CADA ITEM COTADO
        wb = openpyxl.load_workbook('./COTACAO/COTACAO.xlsx', data_only=True)['Controle']
        self.pregao = str(wb.cell(2,1).value)
        self.uasg = str(wb.cell(2,2).value)
        wb = openpyxl.load_workbook('./COTACAO/COTACAO.xlsx', data_only=True)['Planilha1']
        self.itens=[]
        for row in range(2,wb.max_row):
            rowItens=[]
            colunas_interesse=[1,4,5,14]
            colunas_monetarias =[4,14]
            for col in colunas_interesse:
                if(col in colunas_monetarias):
                    rowItens.append(round(wb.cell(row,col).value,2))
                else:
                    rowItens.append(wb.cell(row,col).value)
            self.itens.append(rowItens)

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
        
        finalizado = False
        try:
            sel.clicar_xpath(self,'/html/body/modal-container/div/div/app-dialog-confirmacao/div/div/div[3]/div/div/button')
            finalizado = True
        except:
            pass
        
        if(finalizado == False):
            while(True):#CICLO DE REPETIÇÃO ATÉ QUE NÃO HAJA ITENS EM DISPUTA E AGUARDANDO
                if(self.navegacao_itens[0].text != itens_aguardando_disputa or self.navegacao_itens[1].text != itens_fase_disputa):
                    if(self.navegacao_itens[1].text != itens_fase_disputa):
                        self.navegacao_itens[1].click()
                        self.reconhecer_itens_disputa(self)#COMEÇAR O CICLO DE DISPUTA DE LANCES
                    elif(self.navegacao_itens[0].text != itens_aguardando_disputa):
                        self.navegacao_itens[0].click()
                        aguardando = True
                        while(aguardando):#CICLO DE ESPERA ATÉ QUE ALGUM ITEM AGUARDANDO ENTRE EM FASE DE LANCES
                            if(self.navegacao_itens[1].text != itens_fase_disputa):
                                aguardando=False
                            else:
                                time.sleep(5)
                else:
                    self.extrair_relatorio(self)
                    break
        else:
            self.extrair_relatorio(self)

    def reconhecer_itens_disputa(self):#CHAMA A FUNÇÃO DE ENVIO DE LANCES BASEADO NA ETAPA E TEMPO RESTANTE DE CADA ITEM
        itens_em_disputa = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[2]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div/div')
        sel.obter_elemento_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[2]/div/app-disputa-fornecedor-itens/div/div/div[2]/div/div[1]/button').click()
        for item in itens_em_disputa:#CONFERIR QUAL O ESTADO DA DISPUTA (ETAPA ABERTA - ETAPA FECHADA - FINALIZADO)
            try:
                item_disputa = {}
                etapa_disputa = str(item.find_element_by_xpath('./div[1]/div[1]/div[2]').text)
                if(etapa_disputa.find(item_etapa_aberta) >= 0):
                    codigo_item = str(item.find_element_by_xpath('./div[1]/div[1]/div[1]/div[1]/span[1]').text)
                    atual_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[1]').text)
                    nosso_valor = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[2]').text)
                    tempo_restante = str(item.find_element_by_xpath('./div[1]/div[2]/div/div[2]/span/span').text)
                    intervalo_lances = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/span/small').text)
                    input = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/input')
                    botao_confirma = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/div/button/u')
                    menu_lances = item.find_element_by_xpath('.//*[@title="Mostrar propostas/lances do item"]')
                    item_disputa = {
                        'webelement':item,
                        'item':codigo_item,
                        'etapa':item_etapa_aberta,
                        'atual_valor':atual_valor,
                        'nosso_valor':nosso_valor,
                        'tempo_restante': converter_tempo_restante(tempo_restante),
                        'intervalo_lances':converter_intervalo_minimo(intervalo_lances),
                        'input':input,
                        'botao_confirma':botao_confirma,
                        'menu_lances':menu_lances,
                    }
                    if(item_disputa['tempo_restante']>5 and item_disputa['tempo_restante']<350):
                        print('decidir disputa do item: ',item_disputa['item'])
                        self.decidir_lance(self,item_disputa)
                elif(etapa_disputa.find(item_etapa_fechada)>=0):#CONFIRMAR QUAIS DADOS QUE ESTÃO SENDO ENCONTRADOS
                    codigo_item = str(item.find_element_by_xpath('./div[1]/div[1]/div[1]/div[1]/span[1]').text)
                    intervalo_lances = str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/span/small').text)
                    input = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/input')
                    botao_confirma = item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[2]/div[2]/div/div[1]/div/button/u')
                    item_disputa = {
                        'item':codigo_item,
                        'etapa':item_etapa_fechada,
                        'atual_valor':atual_valor,
                        'nosso_valor':nosso_valor,
                        'tempo_restante':tempo_restante,
                        'intervalo_lances':converter_intervalo_minimo(intervalo_lances),
                        'input':input,
                        'botao_confirma':botao_confirma,
                    }
            except:
                type, val, tb = sys.exc_info()
                traceback.clear_frames(tb)
                raise type(val).with_traceback(tb)

    def decidir_lance(self,item):#DIFERENCIAR A ESTRATEGIA DE LANCE COM A ETAPA DO ITEM
        for cotado in self.itens:
            if(str(cotado[0]) == str(item['item'])):
                if(item['etapa'] == item_etapa_aberta):
                    atual = converter_texto_para_decimal(item['atual_valor'])
                    nosso = converter_texto_para_decimal(item['nosso_valor'])
                    intervalo = float(item['intervalo_lances'])
                    item['menu_lances'].click()
                    time.sleep(sel_delay)
                    try:#ABRIR O TABELA COM OS MELHORES LANCES
                        item['webelement'].find_element_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/ul/li[2]/a').click()
                        time.sleep(sel_delay)
                        item['webelement'].find_element_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/ul/li[2]/a').click()
                    except:
                        item['webelement'].find_element_by_xpath('./div[2]/div[2]/div/app-botao-icone/span/button/i').click()
                        time.sleep(sel_delay)
                        item['webelement'].find_element_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/ul/li[2]/a').click()
                    time.sleep(sel_delay)
                    linhas_tabela = item['webelement'].find_elements_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/div/p-tabpanel[2]/div/app-melhores-valores/div/div/table/tbody/*')
                    valores = []
                    while(True):
                        time.sleep(sel_delay)
                        linhas_tabela = item['webelement'].find_elements_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/div/p-tabpanel[2]/div/app-melhores-valores/div/div/table/tbody/*')
                        if(len(linhas_tabela) > 0):
                            break
                    for n in range(0,len(linhas_tabela)):
                        aux_valor = linhas_tabela[n].find_element_by_xpath('./td[2]').text
                        valores.append(float(converter_texto_para_decimal(aux_valor)))
                        print(aux_valor)
                        if((valores[n] > cotado[3]) and (valores[n] < float(nosso))):
                            if((valores[n-1] < cotado[3]) and (valores[n]-1 < float(nosso))):
                                lance = valores[n]-intervalo
                                while(lance>float(nosso)-intervalo):
                                    lance -= 0.05
                                print('Item ',item['item'],' -> R$ ',lance)
                                #enviar_lance(self,item,valor)
                elif(item['etapa'] == item_etapa_fechada):
                    atual = converter_texto_para_decimal(item['atual_valor'])
                    nosso = converter_texto_para_decimal(item['nosso_valor'])
                    intervalo = converter_intervalo_minimo(item['intervalo_lances'])
                    if(cotado[3] < float(atual)):
                        print('Nosso preço está mais baixo que o atual para o item ', item['item'])
                        if(float(nosso) > (float(atual)*1.1)):
                            print('Dar lance de R$ ', str(float(nosso)*1.1),' para o item ',item['item'])
                            item['menu_lances'].click()
                            return
                        item['menu_lances'].click()#ABRIR E FECHAR MENU MELHORES LANCES#########################################
                        return
        item['menu_lances'].click()

    def enviar_lance(self, item, valor):
        valor = str(valor).replace('.',',')
        sel.enterFieldElement(item['input'],valor)
        item['botao_confirma'].click()
        sel.obter_elemento_xpath(self,'/html/body/modal-container/div/div/app-dialog-confirmacao/div/div/div[3]/div/div[2]/button').click()
        print('lance enviado')
        return

    def extrair_relatorio(self):#TESTAR TESTAR TESTAR
        print('TESTE - Extraiir relatório')
        self.navegacao_itens[2].click()
        itens_encerrados = sel.obter_elementos_xpath(self,'/html/body/app-root/div/div/div/app-cabecalho-disputa-fornecedor/div[5]/div[2]/app-disputa-fornecedor/div/p-tabview/div/div/p-tabpanel[3]/div/app-disputa-fornecedor-itens/div/p-dataview/div/div[2]/div/div')
        print('Itens para relatório: ',len(itens_encerrados))
        resultado_por_item = []
        for item in itens_encerrados:
            aux_item = {}
            aux_item = {
                'item':item.find_element_by_xpath('./div[1]/div[1]/div/div/span[1]').text,
                'melhor_valor' : str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[1]').text),
                'nosso_valor' : str(item.find_element_by_xpath('./div[2]/div[1]/div[2]/div/div[1]/div[2]/div[2]').text),
                }
            item.find_element_by_xpath('./div[2]/div[2]/div/app-botao-icone/span/button/i').click()
            item.find_element_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/ul/li[2]/a').click()
            time.sleep(sel_delay)
            linhas_tabela = item.find_elements_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/div/p-tabpanel[2]/div/app-melhores-valores/div/div/table/tbody/tr')
            while(True):
                time.sleep(sel_delay)
                linhas_tabela = item.find_elements_by_xpath('./div[3]/app-listagem-propostas-lances-item/p-tabview/div/div/p-tabpanel[2]/div/app-melhores-valores/div/div/table/tbody/tr')
                if(len(linhas_tabela)>0):
                    break
            colocacao = 1
            for linha in linhas_tabela:
                if (linha.find_element_by_xpath('./td[2]').text == aux_item['nosso_valor']):
                    break
                else:
                    colocacao += 1
            sel.clicar_subelemento(self.sel_driver, item, './div[2]/div[2]/div/app-botao-icone/span/button/i')
            aux_item['colocacao'] = colocacao
            resultado_por_item.append(aux_item)
        print(resultado_por_item)

bot = ComprasNet()
bot.iniciar()

def fechar_webdriver(self):
    choose = input('>')
    self.sel_driver.close()

fechar_webdriver(bot)