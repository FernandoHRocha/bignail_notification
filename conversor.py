

def monetario(valor):
    result = ''.join(valor.split())[:-2].replace('.','').replace('R$','').replace(',','.')
    return float(result)

def temporizador(valor):
    valor = valor.rsplit(':')
    result = int(valor[0])*60+int(valor[1])
    return result

def intervalo(valor, base):
    valor = valor.replace('Intervalo mínimo entre lances: ','')
    result=''
    if (valor.find('%') < 0):
        result = monetario(valor)
    else:
        valor = ''.join(valor.split())[:-2].replace('.','').replace('%','').replace(',','.')
        print(valor)
        #não funcionando da maneira esperada
        #result = int(valor)*base
    return result