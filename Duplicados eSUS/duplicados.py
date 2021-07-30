# -*- coding: utf-8 -*-
"""
Created on Wed Jul 28 12:30:06 2021

@author: Info
"""

import re
import pandas as pd
from datetime import datetime
from timeit import default_timer as timer
import numpy as np

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
    
    if(type(Nome) is float):
        Nome = None
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
    # if len(Ids) > 1:
    #     print("Achei um duplicado! Notificação = " + str(row[1]))
    for idAssessor in Ids:
        if tabelaEsus.at[row.Index, "IDS ASSESSOR"] == None:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = str(int(idAssessor))
        else:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = tabelaEsus.at[row.Index, "IDS ASSESSOR"] + ", " + str(int(idAssessor))
    return notifs

def formataImpressao(tabela, tipo):
    #tabela = mascaraDataEsus(tabela)
    tabela["NUmero da NotificaCAo"] = tabela["NUmero da NotificaCAo"].astype(str).str.replace(".0", "")
    if tipo == "CORRETO":        
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas", "Nome Completo da MAe", "Sexo", "Telefone de Contato", "Telefone Celular", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "Notificante CNES", "Notificante Email", "Notificante Nome Completo", "SITUACAO"]]
    elif tipo == "INCORRETO":
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas",  "Nome Completo da MAe", "Sexo", "Telefone de Contato", "Telefone Celular", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "Notificante CNES", "Notificante Email", "Notificante Nome Completo", "Motivo", "SITUACAO"]]
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
    aux = {"NUmero da NotificaCAo": row[paramColunaNumNotEsus], "Nome Completo": row[paramColunaNomeEsus], "CPF": row[paramColunaCpfEsus], "CNS": row[paramColunaCnsEsus], "Data de Nascimento": row[paramColunaDataNascEsus], "Nome Completo da MAe": row[paramColunaNomeMaeEsus], "Sexo": row[paramColunaSexoEsus], "Telefone de Contato": row[paramColunaTelContatoEsus], "Telefone Celular": row[paramColunaTelCelEsus], "CEP": row[paramColunaCepEsus], "Logradouro": row[paramColunaLogradouroEsus], "NUmero": row[paramColunaNumEndEsus], "Complemento": row[paramColunaCompEndEsus], "Bairro": row[paramColunaBairroEndEsus], "RaCa/Cor": row[paramColunaRacaEsus], "Data da NotificaCAo": row[paramColunaDataNotifEsus], "Data do inIcio dos sintomas": row[paramColunaDataSintomasEsus], "Notificante CNES": row[paramColunaCnesEsus], "Notificante Email": row[paramColunaEmailEsus], "Notificante Nome Completo": row[paramColunaNotifNomeEsus], "IDS ASSESSOR": tabelaTotalEsus.at[row.Index, "IDS ASSESSOR"], "SITUACAO": tabelaTotalEsus.at[row.Index, "SITUACAO"]}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

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

endReadEsus = timer()
print("Terminei de ler a tabela do eSUS aos " + str(endReadEsus - start) + " segundos.")

qtdNotif = len(tabelaTotalEsus)
porcAtual = 0
countLinhas = 0
porcentil = round(qtdNotif / 100)

#Puxa a tabela do Assessor
tabelaTotalAssessor = pd.read_excel("lista total assessor.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by=["Cód. Paciente", "Data da Notificação"]) #Lê a tabela do Assessor e classifica por ID do paciente e Data da Notificação
tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS'] = formataCpfCns(tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS']) #Formata as colunas de CPF e CNS baseado na regra de negócio
tabelaTotalAssessor = limpaAcentos(tabelaTotalAssessor) #Limpa acentos e 'ç' da tabela do Assessor

endReadAssessor = timer()
print("Terminei de ler a tabela do Assessor aos " + str(endReadAssessor - start) + " segundos.")

regex = re.compile('.*\,.*', re.IGNORECASE)
tabelaTotalEsusIncorreta = []

for row in tabelaTotalEsus.itertuples():
    notifAssessor = acharId(row, tabelaTotalAssessor, tabelaTotalEsus)
    situacao = descobreSituacao(tabelaTotalEsus, row)
    IdsAssessor = tabelaTotalEsus.at[row.Index, "IDS ASSESSOR"]
    if(pd.isnull(IdsAssessor) or not regex.match(IdsAssessor)):
        appendTabelaAuxiliar(tabelaTotalEsusIncorreta, row, "Não é duplicado")
        tabelaTotalEsus.drop(row.Index, inplace=True)
    # else:
    #     print("Realmente é duplicado! Notificacao = " + str(row[1]))
    countLinhas += 1
    if countLinhas >= (porcAtual + 1) * porcentil:
        porcAtual += 1
        porcTimer = timer()
        print("Concluido: " + str(porcAtual) + "% aos " + str(porcTimer - start) + " segundos.\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.")
        
tabelaTotalEsusIncorreta = pd.DataFrame(tabelaTotalEsusIncorreta)
        
with pd.ExcelWriter('IDs Duplicados eSUS ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerDebug:
    tabelaTotalEsus = formataImpressao(tabelaTotalEsus, "CORRETO")
    tabelaTotalEsusIncorreta = formataImpressao(tabelaTotalEsusIncorreta, "INCORRETO")
    tabelaTotalEsus.to_excel(writerDebug, "Corretos", index=False)
    tabelaTotalEsusIncorreta.to_excel(writerDebug, "Incorretos", index=False)