# Robo completo para atuação em processos de licitação

# Logar no sistema do compras net
    Entrar no site do compras net
    Inserir as credenciais da empresa
    Expor lista dos modos de operação
        Cadastrar proposta
        Disputar lances
        Buscar atualizações

# (MODO DISPUTA - ABERTO) Leitura da cotação
    Abrir a pasta que irá conter a planilha de cotação
    Ler a planilha de cotação > variaveis globais
        número do pregão
        código uasg
        item
        preço unitário inicial
        preço unitário mínimo
        quantidade

# (MODO DISPUTA - ABERTO) Acessar a disputa referente a planilha de cotação
    Comparar os pregões em disputa com os dados fornecidos da cotação
        número do pregão
        código uasg
    Acessar a disputa referente a planilha de cotação
    Identificar o modo de disputa
        Aberto
        Aberto/Fechado

# (MODO DISPUTA - ABERTO) Loop de disputa
    Coletar informações dos itens em disputa
        Número do item
        Tempo restante da disputa
        Melhor Valor
        Meu valor
        Intervalo mínimo entre lances > variaveis globais
        Lances entre o valor mínimo de cotação
    Decisão baseada no tempo restante de disputa de item para dar lance
    Enviar novo lance

# (MODO DISPUTA - ABERTO) Relatório de disputa
    Coletar as informações
        Item
        Melhor Valor
        Meu valor
        Classificação
    Armazenar as informações em planilha
        Pregão
        Uasg
        Item
        Classificação
        Melhor Valor
        Meu Valor

