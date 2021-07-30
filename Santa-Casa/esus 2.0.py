# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 13:12:28 2020

@author: Thales
Script que compara a base do sistema eSUS Notifica com o sistema
Assessor Público de uso municipal. Tem como entrada uma exportação
do sistema eSUS Notifica ('lista total esus.xlsx') e a exportação
dos agravos no sistema Assessor Público ('lista total assessor.xls') e tem
como resultado uma planilha com todos os agravos que devem ser
inseridos no Assessor, classificados entre Positivos, Negativos,
Suspeitos e Monitoramento.
"""


import pandas as pd
from datetime import timedelta
from datetime import datetime
import numpy as np
from timeit import default_timer as timer
import os

teste = None

def formataCpfCns(colunaCpf, colunaCns):
    for index, value in colunaCpf.items():
        if(len(str(colunaCns[index])) != 15):
            colunaCns.at[index] = None
        if(pd.isnull(value)):
            colunaCpf.at[index] = None
            continue
        value = value.strip()
        value = value.replace(".", "")
        value = value.replace("-", "")
        while len(value) < 11:
            value = "0" + value
        colunaCpf.at[index] = str(value)
        # if(len(colunaCpf[index]) < 14):
        #     colunaCpf[index] = value[:3] + '.' + value[3:6] + "." + value[6:9] + "-" + value[9:]
    return colunaCpf, colunaCns

def acharId(row, tabela, tabelaEsus):
    Cpf = row.CPF
    Cns = row.CNS
    Nome = row[paramColunaNomeEsus]
    Dn = row[paramColunaDataNascEsus]
    #row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus]
    #print("TESTE - DN param " + str(Dn) + "\nPrimeira data da tabela: " + str(tabela['Data Nasc.'][0]))
    if(type(Cpf) is float):
        print("ACHEI UM CPF FLOAT: " + str(Cpf) + "\nNome: " + str(Nome))
        Cpf = None
    if(Cpf is not None):
        filtro1 = tabela['CPF'] == Cpf.strip()
    else:
        filtro1 = False
    
    if(type(Cns) is float):
        Cns = None
    if(Cns is not None):
        filtro2 = tabela['CNS'] == Cns.strip()
    else:
        filtro2 = False
        
    if(Nome is not None and Dn is not None):
        filtro3 = tabela['Paciente'] == Nome.strip()
        filtro4 = tabela['Data Nasc.'] == Dn
    else:
        filtro3 = False
        filtro4 = False
    global teste
    notifs = tabela.where(filtro1 | filtro2 | (filtro3 & filtro4)).dropna(how='all')
    teste = notifs
    Ids = notifs["COd. Paciente"].drop_duplicates().tolist()
    for idAssessor in Ids:
        if tabelaEsus.at[row.Index, "IDS ASSESSOR"] == None:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = str(int(idAssessor))
        else:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = tabelaEsus.at[row.Index, "IDS ASSESSOR"] + ", " + str(int(idAssessor))
    return notifs

def acharUnidade(row):
    cnesUnidades = { #Fazer verificação pra email da Drogaria São Paulo
    "3142531": "HOSPITAL SAO JORGE",
    "2052970": "LABORATORIO DE ANALISES CLINICAS ESTIMA BORGES SC",
    "3549208": "LABORATORIO SUZUKI",
    "5124700": "LABORAMAS LABORATORIO DE ANALISES CLINICAS",
    "7909837": "INSTITUTO RESPIRE VACINAS E INFUSOES",
    "9344209": "POSTO DE COLETA SAO FRANCISCO",
    "9567771": "POSTO DE COLETA SAO FRANCISCO",
    "9371508": "SAUDE MED MULTICLINICA ESPECIALIZADA",
    "6316344": "CENTRO DIAGNOSTICO MEDICO ASSOCIADO LTDA",
    "419907": "DROGA RAIA"
    }
    
    emailsUnidades = {
    "yaah_cardoso@hotmail.com": "DROGA RAIA",
    "maicon.douglas19@hotmail.com": "DROGARIA SAO PAULO",
    "robertamap@yahoo.com.br": "DROGA RAIA",
    "barretos2@dpsp.com.br": "DROGARIA SAO PAULO"
    }
    if(not pd.isnull(row[paramColunaCnesEsus])):
        return cnesUnidades[str(int(row[paramColunaCnesEsus]))]
    else:
        return emailsUnidades[row[paramColunaEmailEsus]]
    

def mascaraDataEsus(tabelaEsus):
    if tabelaEsus.empty:
        return tabelaEsus
    colunasData = ['Data da NotificaCAo', 'Data de Nascimento', 'Data do inIcio dos sintomas']
    for i in colunasData:
        #print("Entrei em mascaraDataEsus com i = " + i)
        tabelaEsus[i] = pd.to_datetime(tabelaEsus[i])
        tabelaEsus[i] = tabelaEsus[i].dt.strftime('%d/%m/%Y')
    return tabelaEsus

def limpaUtf8(tabela): #Função que limpa os caracteres bugados do UTF-8 do arquivo de exportação do eSUS
    tabela.replace(to_replace={"Ã§": "C", "Ã‡": "C", "Ãµ": "O", "Ã³": "O", "Ã´": "O", "Ã\”": "O", "Ã­": "I", "Ãº": "U", "Ãš": "U", "ÃŠ": "E", "Ãª": "E", "Ã©": "E", "Ã‰": "E", "ÃG": "IG", "Ãƒ": "A", "Ã£": "A", "Ã¡": "A", "Ã¢": "A"}, inplace=True, regex=True)
    tabela.replace(to_replace={"Ã": "A"}, inplace=True, regex=True)
    tabela.rename(columns={"NÃºmero da NotificaÃ§Ã£o": "NUmero da NotificaCAo", "Nome Completo da MÃ£e": "Nome Completo da MAe", "NÃºmero (ou SN para Sem NÃºmero)": "NUmero", "RaÃ§a/Cor": "RaCa/Cor", "Data da NotificaÃ§Ã£o": "Data da NotificaCAo", "Resultado (PCR/RÃ¡pidos)": "Resultado (PCR/RApidos)", "MunicÃ­pio de ResidÃªncia": "MunicIpio de ResidEncia", "Data do inÃ­cio dos sintomas": "Data do inIcio dos sintomas", "ClassificaÃ§Ã£o Final": "ClassificaCAo Final", "EvoluÃ§Ã£o Caso": "EvoluCAo Caso", "Teste SorolÃ³gico": "Teste SorolOgico"}, inplace=True)
    return tabela

def limpaAcentos(tabela): #Função que limpa os caracteres especiais da tabela do Assessor
    tabela.replace(to_replace={"Á": "A", "á": "A", "Â": "A", "â": "A", "Ã": "A", "ã": "A",
                               "É": "E", "é": "E", "Ê": "E", "ê": "E",
                               "Í": "I", "í": "I", "Î": "I", "î": "I",
                               "Ó": "O", "ó": "O", "Ô": "O", "ô": "O", "Õ": "O", "õ": "O",
                               "Ú": "U", "ú": "U", "Û": "U", "û": "U",
                               "Ç": "C", "ç": "C"}, inplace=True, regex=True)
    tabela.rename(columns={"Cód. Paciente": "COd. Paciente", "Data da Notificação": "Data da NotificaCAo", "Situação": "SituaCAo"}, inplace=True)    
    return tabela

def appendTabelaAuxiliar(tabela, row, motivo):
    #aux = {"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Data Not.": row[paramColunaDataNotifEsus], "Notificante CNES": row[paramColunaCnesEsus], "Notificante Email": row[paramColunaEmailEsus], "Situacao": row[paramColunaSituacaoEsus]}
    aux = {"NUmero da NotificaCAo": row[paramColunaNumNotEsus], "Nome Completo": row[paramColunaNomeEsus], "CPF": row[paramColunaCpfEsus], "CNS": row[paramColunaCnsEsus], "Data de Nascimento": row[paramColunaDataNascEsus], "Nome Completo da MAe": row[paramColunaNomeMaeEsus], "Sexo": row[paramColunaSexoEsus], "Telefone de Contato": row[paramColunaTelContatoEsus], "Telefone Celular": row[paramColunaTelCelEsus], "CEP": row[paramColunaCepEsus], "Logradouro": row[paramColunaLogradouroEsus], "NUmero": row[paramColunaNumEndEsus], "Complemento": row[paramColunaCompEndEsus], "Bairro": row[paramColunaBairroEndEsus], "RaCa/Cor": row[paramColunaRacaEsus], "Data da NotificaCAo": row[paramColunaDataNotifEsus], "Data do inIcio dos sintomas": row[paramColunaDataSintomasEsus], "Notificante CNES": row[paramColunaCnesEsus], "Notificante Email": row[paramColunaEmailEsus], "Notificante Nome Completo": row[paramColunaNotifNomeEsus], "IDS ASSESSOR": row[paramColunaIdsEsus]}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

def formataImpressao(tabela, tipo):
    #tabela = mascaraDataEsus(tabela)
    tabela["NUmero da NotificaCAo"] = tabela["NUmero da NotificaCAo"].astype(str)
    if tipo == "CORRETO":        
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas", "Nome Completo da MAe", "Sexo", "Telefone de Contato", "Telefone Celular", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "Notificante CNES", "Notificante Email", "Notificante Nome Completo"]]
    elif tipo == "INCORRETO":
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas",  "Nome Completo da MAe", "Sexo", "Telefone de Contato", "Telefone Celular", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "Notificante CNES", "Notificante Email", "Notificante Nome Completo", "Motivo"]]
    else:
        return tabela

def descobreSituacao(tabela, row):
    if row[paramColunaResultadoPCREsus] == "Positivo" or (pd.isnull(row[paramColunaResultadoPCREsus]) and (row[paramColunaResultadoTotaisEsus] == "Reagente" or row[paramColunaResultadoIgaEsus] == "Reagente" or row[paramColunaResultadoPCREsus] == "Reagente")):
        tabela.at[row.Index, "SITUACAO"] = "POSITIVO"
        return "POSITIVO"
    if row[paramColunaResultadoPCREsus] == "Negativo" or (pd.isnull(row[paramColunaResultadoPCREsus]) and (row[paramColunaResultadoTotaisEsus] == "NAo Reagente" or row[paramColunaResultadoIgaEsus] == "NAo Reagente" or row[paramColunaResultadoPCREsus] == "NAo Reagente")):
        tabela.at[row.Index, "SITUACAO"] = "NEGATIVO"
        return "NEGATIVO"
    if(pd.isnull(row[paramColunaResultadoPCREsus]) and pd.isnull(row[paramColunaResultadoTotaisEsus]) and pd.isnull(row[paramColunaResultadoIgaEsus]) and pd.isnull(row[paramColunaResultadoPCREsus])):
        if(row[paramColunaClassifFinalEsus] != "Descartado"):
            tabela.at[row.Index, "SITUACAO"] = "SUSPEITO"
            return "SUSPEITO"
        else:
            tabela.at[row.Index, "SITUACAO"] = "DESCARTADO"
            return "DESCARTADO"
    tabela.at[row.Index, "SITUACAO"] = "SEM SITUACAO"
    return "SEM SITUACAO"

def trataSuspeito(row, notifAssessor, tabela, tabelaIncorreta):
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO")
        tabela.drop(row.Index, inplace=True)
        return
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        appendTabelaAuxiliar(tabelaIncorreta, row, "SUSPEITO EM ABERTO")
        tabela.drop(row.Index, inplace=True)
        return
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):                
                appendTabelaAuxiliar(tabelaIncorreta, row, "NEGATIVO DENTRO DO PERIODO")
                tabela.drop(row.Index, inplace=True)
                return
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):                
                appendTabelaAuxiliar(tabelaIncorreta, row, "DESCARTADO DENTRO DO PERIODO")
                tabela.drop(row.Index, inplace=True)
                return

def trataPositivo(row, notifAssessor, tabela, tabelaIncorreta):
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO")
        tabela.drop(row.Index, inplace=True)
        return

def trataNegativo(row, notifAssessor, tabela, tabelaIncorreta):
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO")
        tabela.drop(row.Index, inplace=True)
        return   
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                appendTabelaAuxiliar(tabelaIncorreta, row, "NEGATIVO DENTRO DO PERIODO")
                tabela.drop(row.Index, inplace=True)
                return
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                appendTabelaAuxiliar(tabelaIncorreta, row, "DESCARTADO DENTRO DO PERIODO")
                tabela.drop(row.Index, inplace=True)
                return
        
def converteFiltro(tabela, lblCnes, lblEmail, listaFiltro):
    filtroFinal = pd.Series(dtype=object)
    for filtroCnes in listaFiltro["cnes"]:
        if filtroFinal.empty:
            filtroFinal = tabela[lblCnes] == filtroCnes
        else:
            filtroFinal = (filtroFinal) | (tabela[lblCnes] == filtroCnes)
    
    for filtroEmail in listaFiltro["email"]:
        if filtroFinal.empty:
            filtroFinal = tabela[lblEmail] == filtroEmail
        else:
            filtroFinal = (filtroFinal) | (tabela[lblEmail] == filtroEmail)
    
    return filtroFinal

paramDiasAtras = 16 #Parâmetro que define quantos dias atrás ele considera na planilha de suspeitos e monitoramento (ex: notificações de até X dias atrás serão analisadas, antes disso serão ignoradas)
paramDiasNegativo = 5 #Parâmetro que define quantos dias atrás uma notificacao de negativo deve ser considerada "a mesma" notificacao. Acima desse parâmetro deve ser considerada uma nova notificacao (agora também usada para descartados)
paramColunaNomeEsus = 30 ##Parâmetro que define qual é o index da coluna que possui o nome completo do paciente no eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaDataNascEsus = 25 ##Parâmetro que define qual é o index da coluna que possui a data de nascimento do paciente no eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaDataNotifEsus = 54 #Parâmetro que define qual é o index da coluna que possui a data de notificacao do eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaCnesEsus = 67
paramColunaEmailEsus = 69

paramColunaResultadoPCREsus = 5
paramColunaEstadoTesteEsus = 8
paramColunaClassifFinalEsus = 10
paramColunaResultadoTotaisEsus = 11
paramColunaResultadoIgaEsus = 12
paramColunaResultadoIgmEsus = 13
paramColunaSituacaoEsus = 72
paramColunaIdsEsus = 73
paramColunaNumNotEsus = 1
paramColunaCpfEsus = 22
paramColunaCnsEsus = 24
paramColunaNomeMaeEsus = 30
paramColunaSexoEsus = 19
paramColunaTelContatoEsus = 17
paramColunaTelCelEsus = 29
paramColunaCepEsus = 33
paramColunaLogradouroEsus = 34
paramColunaNumEndEsus = 35
paramColunaCompEndEsus = 36
paramColunaBairroEndEsus = 37
paramColunaRacaEsus = 38
paramColunaDataSintomasEsus = 64
paramColunaNotifNomeEsus = 70

paramColunaDataNotifAssessor = 11 #Parâmetro igual ao paramColunaDataNotifEsus porém para o Assessor
paramDataAtual = datetime.today() #Parâmetro que define qual é a data atual para o script fazer a comparacao dos dias para trás (padrão: datetime.today() = data atual do sistema)

print("Data de hoje: " + paramDataAtual.strftime("%d/%m/%Y"))
print("Data de " + str(paramDiasAtras) + " dias atrás: " + (paramDataAtual - timedelta(days=paramDiasAtras)).strftime("%d/%m/%Y"))

start = timer()

tabelaTotalEsus = pd.read_excel("lista total esus.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_, 'Nome Completo': np.unicode_}).sort_values(by="Data da NotificaÃ§Ã£o", ignore_index=True) #Lista total das notificações do eSus
tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS'] = formataCpfCns(tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS']) #Formata as colunas de CNS e CPF conforme a regra de negócio
tabelaTotalEsus['Nome Completo'] = tabelaTotalEsus['Nome Completo'].str.upper() #Transforma a coluna de nome em Caixa Alta
tabelaTotalEsus = limpaUtf8(tabelaTotalEsus) #Limpa os caracteres que vem "bugados" do eSUS por conta de acentos e 'ç'
tabelaTotalEsus = tabelaTotalEsus.where(tabelaTotalEsus['MunicIpio de ResidEncia'] == "Barretos").dropna(how='all') #Filtra apenas pelo município de residência Barretos
tabelaTotalEsus = tabelaTotalEsus.where(tabelaTotalEsus['EvoluCAo Caso'] != "Cancelado").dropna(how='all') #Retira as notificações canceladas
tabelaTotalEsus = tabelaTotalEsus.where((tabelaTotalEsus['Teste SorolOgico'] != "Anticorpos Totais") & (tabelaTotalEsus['Teste SorolOgico'] != "IgG, Anticorpos Totais") & (tabelaTotalEsus['Teste SorolOgico'] != "Anticorpos Totais, IgG")).dropna(how='all')
tabelaTotalEsus.insert(len(tabelaTotalEsus.columns), "SITUACAO", None)
tabelaTotalEsus.insert(len(tabelaTotalEsus.columns), "IDS ASSESSOR", None)

listaFiltroStc = {"cnes": [2092611], "email": []}
filtroStc = converteFiltro(tabelaTotalEsus, 'Notificante CNES', 'Notificante Email', listaFiltroStc)
listaFiltroPio = {"cnes": [2090236], "email": []}
filtroPio = converteFiltro(tabelaTotalEsus, 'Notificante CNES', 'Notificante Email', listaFiltroPio)
listaFiltroHans = {"cnes": [9662561], "email": []}
filtroHans = converteFiltro(tabelaTotalEsus, 'Notificante CNES', 'Notificante Email', listaFiltroHans)
listaFiltroExternas = {"cnes": [9344209, 2052970, 9567771, 9371508, 3549208, 7909837, 3142531, 5124700, 6316344, 419907, 2074176],
                  "email": ["maicon.douglas19@hotmail.com", "yaah_cardoso@hotmail.com", "robertamap@yahoo.com.br", "barretos2@dpsp.com.br", "marcela@ltaseguranca.com.br"]}
filtroExternas = converteFiltro(tabelaTotalEsus, 'Notificante CNES', 'Notificante Email', listaFiltroExternas)
listaFiltroInternas = {"cnes": [2035731 , 2048736 , 2048744 , 2053314 , 2053062 , 2061473 , 2064081 , 2064103 , 2093642 , 2093650 , 2784572 , 2784580 , 2784599 , 5562325 , 5562333 , 7035861 , 7122217 , 7565577 , 7585020 , 2053306],
                  "email": ["machado.vl@bol.com.br" , "marcioli@bol.com.br" , "marcosnascimento32@outlook.com"]}
filtroInternas = converteFiltro(tabelaTotalEsus, 'Notificante CNES', 'Notificante Email', listaFiltroInternas)


tabelaTotalEsus = tabelaTotalEsus.where((filtroStc) | (filtroPio) | (filtroHans) | (filtroExternas) | (filtroInternas)).dropna(how='all')
qtdNotif = len(tabelaTotalEsus)
porcAtual = 0

endReadEsus = timer()
print("Terminei de ler a tabela do eSUS aos " + str(endReadEsus - start) + " segundos.")

#Puxa a tabela do Assessor
tabelaTotalAssessor = pd.read_excel("lista total assessor.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by=["Cód. Paciente", "Data da Notificação"]) #Lê a tabela do Assessor e classifica por ID do paciente e Data da Notificação
tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS'] = formataCpfCns(tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS']) #Formata as colunas de CPF e CNS baseado na regra de negócio
tabelaTotalAssessor = limpaAcentos(tabelaTotalAssessor) #Limpa acentos e 'ç' da tabela do Assessor

endReadAssessor = timer()
print("Terminei de ler a tabela do Assessor aos " + str(endReadAssessor - start) + " segundos.")

tabelaTotalEsusIncorreta = []
countLinhas = 0
porcentil = round(qtdNotif / 100)

for row in tabelaTotalEsus.itertuples():
    notifAssessor = acharId(row, tabelaTotalAssessor, tabelaTotalEsus)
    situacao = descobreSituacao(tabelaTotalEsus, row)
    if situacao == "POSITIVO":
        trataPositivo(row, notifAssessor, tabelaTotalEsus, tabelaTotalEsusIncorreta)
    elif situacao == "NEGATIVO":
        trataNegativo(row, notifAssessor, tabelaTotalEsus, tabelaTotalEsusIncorreta)
    elif situacao == "SEM SITUACAO":
        print("Achei uma notificação sem situação")        
    countLinhas += 1
    if countLinhas >= (porcAtual + 1) * porcentil:
        porcAtual += 1
        os.system('cls')
        porcTimer = timer()
        print("Concluido: " + str(porcAtual) + "% aos " + str(porcTimer - start) + " segundos.\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.")


endLaco = timer()
print("Terminei de processar o eSUS aos " + str(endLaco - start) + " segundos.")

tabelaTotalEsusIncorreta = pd.DataFrame(tabelaTotalEsusIncorreta)

with pd.ExcelWriter('Debug eSUS ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerDebug:
    tabelaTotalEsus.to_excel(writerDebug, "Corretos", index=False)
    tabelaTotalEsusIncorreta.to_excel(writerDebug, "Incorretos", index=False)

tabelaStc = tabelaTotalEsus.where(filtroStc).dropna(how='all')
tabelaPio = tabelaTotalEsus.where(filtroPio).dropna(how='all')
tabelaHans = tabelaTotalEsus.where(filtroHans).dropna(how='all')
tabelaExternas = tabelaTotalEsus.where(filtroExternas).dropna(how='all')
tabelaInternas = tabelaTotalEsus.where(filtroInternas).dropna(how='all')

filtroStcInc = converteFiltro(tabelaTotalEsusIncorreta, 'Notificante CNES', 'Notificante Email', listaFiltroStc)
filtroPioInc = converteFiltro(tabelaTotalEsusIncorreta, 'Notificante CNES', 'Notificante Email', listaFiltroPio)
filtroHansInc = converteFiltro(tabelaTotalEsusIncorreta, 'Notificante CNES', 'Notificante Email', listaFiltroHans)
filtroExternasInc = converteFiltro(tabelaTotalEsusIncorreta, 'Notificante CNES', 'Notificante Email', listaFiltroExternas)
filtroInternasInc = converteFiltro(tabelaTotalEsusIncorreta, 'Notificante CNES', 'Notificante Email', listaFiltroInternas)

tabelaStcIncorreta = tabelaTotalEsusIncorreta.where(filtroStcInc).dropna(how='all')
tabelaPioIncorreta = tabelaTotalEsusIncorreta.where(filtroPioInc).dropna(how='all')
tabelaHansIncorreta = tabelaTotalEsusIncorreta.where(filtroHansInc).dropna(how='all')
tabelaExternasIncorreta = tabelaTotalEsusIncorreta.where(filtroExternasInc).dropna(how='all')
tabelaInternasIncorreta = tabelaTotalEsusIncorreta.where(filtroInternasInc).dropna(how='all')

with pd.ExcelWriter('Santa Casa ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerStc:
    tabelaStcPos = tabelaStc.where(tabelaStc["SITUACAO"] == "POSITIVO").dropna(how='all')
    tabelaStcPos = formataImpressao(tabelaStcPos, "CORRETO")
    tabelaStcPos.to_excel(writerStc, "Positivos", index=False)    
    
    tabelaStcNeg = tabelaStc.where(tabelaStc["SITUACAO"] == "NEGATIVO").dropna(how='all')
    tabelaStcNeg = formataImpressao(tabelaStcNeg, "CORRETO")
    tabelaStcNeg.to_excel(writerStc, "Negativos", index=False)
        
    tabelaStcIncorreta = formataImpressao(tabelaStcIncorreta, "INCORRETO")
    tabelaStcIncorreta.to_excel(writerStc, "Incorretos", index=False)
    
with pd.ExcelWriter('Pio XII ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerPio:
    tabelaPioPos = tabelaPio.where(tabelaPio["SITUACAO"] == "POSITIVO").dropna(how='all')
    tabelaPioPos = formataImpressao(tabelaPioPos, "CORRETO")
    tabelaPioPos.to_excel(writerPio, "Positivos", index=False)
    
    tabelaPioNeg = tabelaPio.where(tabelaPio["SITUACAO"] == "NEGATIVO").dropna(how='all')
    tabelaPioNeg = formataImpressao(tabelaPioNeg, "CORRETO")
    tabelaPioNeg.to_excel(writerPio, "Negativos", index=False)
        
    tabelaPioIncorreta = formataImpressao(tabelaPioIncorreta, "INCORRETO")
    tabelaPioIncorreta.to_excel(writerPio, "Incorretos", index=False)
    
with pd.ExcelWriter('HANS ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerHans:
    tabelaHansPos = tabelaHans.where(tabelaHans["SITUACAO"] == "POSITIVO").dropna(how='all')
    tabelaHansPos = formataImpressao(tabelaHansPos, "CORRETO")
    tabelaHansPos.to_excel(writerHans, "Positivos", index=False)
    
    tabelaHansNeg = tabelaHans.where(tabelaHans["SITUACAO"] == "NEGATIVO").dropna(how='all')
    tabelaHansNeg = formataImpressao(tabelaHansNeg, "CORRETO")
    tabelaHansNeg.to_excel(writerHans, "Negativos", index=False)
        
    tabelaHansIncorreta = formataImpressao(tabelaHansIncorreta, "INCORRETO")
    tabelaHansIncorreta.to_excel(writerHans, "Incorretos", index=False)
    
with pd.ExcelWriter('Externas ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerExternas:
    tabelaExternasPos = tabelaExternas.where(tabelaExternas["SITUACAO"] == "POSITIVO").dropna(how='all')
    tabelaExternasPos = formataImpressao(tabelaExternasPos, "CORRETO")
    tabelaExternasPos.to_excel(writerExternas, "Positivos", index=False)
    
    tabelaExternasNeg = tabelaExternas.where(tabelaExternas["SITUACAO"] == "NEGATIVO").dropna(how='all')
    tabelaExternasNeg = formataImpressao(tabelaExternasNeg, "CORRETO")
    tabelaExternasNeg.to_excel(writerExternas, "Negativos", index=False)
        
    tabelaExternasIncorreta = formataImpressao(tabelaExternasIncorreta, "INCORRETO")
    tabelaExternasIncorreta.to_excel(writerExternas, "Incorretos", index=False)
    
with pd.ExcelWriter('Internas ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerInternas:
    tabelaInternasPos = tabelaInternas.where(tabelaInternas["SITUACAO"] == "POSITIVO").dropna(how='all')
    tabelaInternasPos = formataImpressao(tabelaInternasPos, "CORRETO")
    tabelaInternasPos.to_excel(writerInternas, "Positivos", index=False)
    
    tabelaInternasNeg = tabelaInternas.where(tabelaInternas["SITUACAO"] == "NEGATIVO").dropna(how='all')
    tabelaInternasNeg = formataImpressao(tabelaInternasNeg, "CORRETO")
    tabelaInternasNeg.to_excel(writerInternas, "Negativos", index=False)
        
    tabelaInternasIncorreta = formataImpressao(tabelaInternasIncorreta, "INCORRETO")
    tabelaInternasIncorreta.to_excel(writerInternas, "Incorretos", index=False)