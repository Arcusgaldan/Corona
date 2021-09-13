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
import re

teste = None

def realmenteSuspeitos(tabelaAssessor):
    """
    Parameters
    ----------
    tabelaAssessor : DataFrame
        Tabela Assessor

    Returns
    -------
    tabelaTotal : DataFrame
        Tabela Assessor, porém com apenas os "Realmente Suspeitos", de acordo com a regra de negócio

    """
    print("Iniciando 'Realmente Suspeitos'...")
    tabelaTotal = tabelaAssessor
    tabelaSuspeitos = tabelaTotal.where((tabelaTotal["Data da Notificação"] >= datetime.strptime('01-06-2020', '%d-%m-%Y')) & (tabelaTotal["Situação"] == "SUSPEITA")).dropna(how='all').sort_values(by=["Cód. Paciente", "Data da Notificação"]).drop_duplicates("Cód. Paciente", ignore_index=True)
    startProcessamento = timer()
    porcAtual = 0
    somaVeloc = 0
    countLinhas = 0
    qtdNotif = len(tabelaSuspeitos)
    porcentil = round(qtdNotif / 100)
    for row in tabelaSuspeitos.itertuples(): #Passando por cada suspeito
        situacoes = tabelaTotal.where(tabelaTotal["Cód. Paciente"] == row[paramColunaIdAssessor]).dropna(how='all') #Pega apenas a situação e data de notificação de cada registro que o suspeito possui na lista total
        flag = False
        #print("Teste de situação: " + situacoes["Situação"].values)
        if "CONFIRMADO" in situacoes["Situação"].values: #Se houver um confirmado na lista de situações
            confirmadosAssessor = situacoes.where(situacoes["Situação"] == "CONFIRMADO").dropna(how='all')
            for rowConfirmado in confirmadosAssessor.itertuples(): #Passa por cada situação
                flag = False
                if rowConfirmado[paramColunaDataNotifAssessor] >= (row[paramColunaDataNotifAssessor] - timedelta(days=90)):
                    tabelaSuspeitos.drop(row.Index, inplace=True) #Tira o suspeito da lista de suspeitos
                    flag = True
                    break
            if not flag:
                if "NEGATIVO" in situacoes["Situação"].values: #Se houver um negativo nas situações
                    negativosAssessor = situacoes.where(situacoes["Situação"] == "NEGATIVO").dropna(how='all')
                    for rowNegativo in negativosAssessor.itertuples():                
                        if rowNegativo[paramColunaDataNotifAssessor] >= (row[paramColunaDataNotifAssessor] - timedelta(days=2)):                    
                            tabelaSuspeitos.drop(row.Index, inplace=True) #Tira o suspeito da lista de suspeitos
                            break
        
        countLinhas += 1
        if countLinhas >= (porcAtual + 1) * porcentil:
            porcAtual += 1
            os.system('cls')
            porcTimer = timer()
            tempoAtual = porcTimer - startProcessamento
            velocEstimada = countLinhas / tempoAtual
            somaVeloc += velocEstimada
            velocMedia = somaVeloc / porcAtual
            tempoEstimadoTotal = (qtdNotif - countLinhas) / velocEstimada
            stringTempoEstimado = segundoEmHora(tempoEstimadoTotal)
            print("REALMENTE SUSPEITOS\nConcluido: " + str(porcAtual) + "% aos " + "{:.0f}".format(tempoAtual) + " segundos (desde o início do laço).\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.\nVelocidade estimada: " + "{:.2f}".format(velocEstimada) + " linhas por segundo.\nVelocidade media: " + "{:.2f}".format(velocMedia) + " linhas por segundo.\nTempo estimado: " + stringTempoEstimado + " para conlusão\n\n")
    
    tabelaTotal = tabelaTotal.where(tabelaTotal["Situação"] != "SUSPEITA").dropna(how='all')
    tabelaTotal = pd.concat([tabelaTotal, tabelaSuspeitos])
    with pd.ExcelWriter('Total e Realmente Suspeitos ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:
        tabelaTotal.to_excel(writer, index=False)
    print("Total de " + str(len(tabelaTotal)) + " total e realmente suspeitos.")
    return tabelaTotal

def formataCpfCns(colunaCpf, colunaCns): 
    """
    Parameters
    ----------
    colunaCpf : Series
        Coluna da Tabela eSUS que possui os CPFs
    colunaCns : Series
        Coluna da Tabela eSUS que possui os CNSs

    Returns
    -------
    colunaCpf : Series
        Coluna da Tabela eSUS que possui os CPFs, formatada com quantidade de caracteres e formatação/máscara
    colunaCns : Series
        Coluna da Tabela eSUS que possui os CNSs, formatada com quantidade de caracteres

    """
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
    """
    Parameters
    ----------
    row : Itertuple
        Linha da Tabela eSUS
    tabela : DataFrame
        Tabela Assessor
    tabelaEsus : DataFrame
        Tabela eSUS

    Returns
    -------
    notifs : DataFrame
        Tabela Assessor recortada com apenas as notificações relativas à essa linha da Tabela eSUS (row)

    """
    Cpf = row.CPF
    Cns = row.CNS
    Nome = row[paramColunaNomeEsus]
    Dn = row[paramColunaDataNascEsus]
    
    if(type(Cpf) is not str): #Checa se o CPF é NaN
        Cpf = None
    if(Cpf is not None):
        filtro1 = tabela['CPF'] == Cpf.strip()
    else:
        filtro1 = False
    
    if(type(Cns) is not str): #Checa se o CNS é NaN
        Cns = None
    if(Cns is not None):
        filtro2 = tabela['CNS'] == Cns.strip()
    else:
        filtro2 = False
    
    if(type(Nome) is not str): #Checa se o Nome é NaN
        Nome = None
    if(Nome is not None and Dn is not None):
        filtro3 = tabela['Paciente'] == Nome.strip()
        filtro4 = tabela['Data Nasc.'] == Dn
    else:
        filtro3 = False
        filtro4 = False
    
    notifs = tabela.where(filtro1 | filtro2 | (filtro3 & filtro4)).dropna(how='all') #Dataframe das notificações desse paciente
    Ids = notifs["COd. Paciente"].drop_duplicates().tolist() #Lista dos IDs encontrados deste paciente
    # if 316914 in Ids:
    #     print("Encontrei o ID 316914!\nIndex: " + str(row.Index) + "\nCPF: " + (Cpf if Cpf is not None else "Vazio") + "\nCNS: " + (Cns if Cns is not None else "Vazio") + "\nNome: " + (Nome if Nome is not None else "Vazio") + "\nDN: " + (str(Dn) if Dn is not None else "Vazio") + "\nNotificações:\n")
    #     for codNotif in notifs["Código"].tolist():
    #         print(str(codNotif))
    #     print("")
    for idAssessor in Ids: #Adiciona os IDs encontrados na coluna "IDS ASSESSOR", concatenando com vírgula onde há mais de 01 ID encontrado
        if tabelaEsus.at[row.Index, "IDS ASSESSOR"] == None:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = str(int(idAssessor))
        else:
            tabelaEsus.at[row.Index, "IDS ASSESSOR"] = tabelaEsus.at[row.Index, "IDS ASSESSOR"] + ", " + str(int(idAssessor))
    return notifs

def mascaraDataEsus(tabelaEsus):
    """
    Parameters
    ----------
    tabelaEsus : DataFrame
        DataFrame representando a tabela eSUS.

    Returns
    -------
    tabelaEsus : DataFrame
        O mesmo DataFrame, porém com as colunas de data devidamente formatadas para impressão
    """
    if tabelaEsus.empty:
        return tabelaEsus
    colunasData = ['Data da NotificaCAo', 'Data de Nascimento', 'Data do inIcio dos sintomas'] #Os nomes de colunas que contém data
    for i in colunasData:
        #print("Entrei em mascaraDataEsus com i = " + i)
        tabelaEsus[i] = pd.to_datetime(tabelaEsus[i])
        tabelaEsus[i] = tabelaEsus[i].dt.strftime('%d/%m/%Y')
    return tabelaEsus

def limpaUtf8(tabela):
    """
    Parameters
    ----------
    tabela : DataFrame
        Tabela eSUS

    Returns
    -------
    tabela : DataFrame
        A mesma tabela recebida, porém com os caracteres "bugados" em UTF-8 substituídos por sua contraparte em ASCII

    """
    tabela.replace(to_replace={"Ã§": "C", "Ã‡": "C", "Ãµ": "O", "Ã³": "O", "Ã´": "O", "Ã\”": "O", "Ã­": "I", "Ãº": "U", "Ãš": "U", "ÃŠ": "E", "Ãª": "E", "Ã©": "E", "Ã‰": "E", "ÃG": "IG", "Ãƒ": "A", "Ã£": "A", "Ã¡": "A", "Ã¢": "A"}, inplace=True, regex=True) #Regras de substituição por caractere
    tabela.replace(to_replace={"Ã": "A"}, inplace=True, regex=True) #Mais regras de substituição
    tabela.rename(columns={"NÃºmero da NotificaÃ§Ã£o": "NUmero da NotificaCAo", "Nome Completo da MÃ£e": "Nome Completo da MAe", "NÃºmero (ou SN para Sem NÃºmero)": "NUmero", "RaÃ§a/Cor": "RaCa/Cor", "Data da NotificaÃ§Ã£o": "Data da NotificaCAo", "Resultado (PCR/RÃ¡pidos)": "Resultado (PCR/RApidos)", "MunicÃ­pio de ResidÃªncia": "MunicIpio de ResidEncia", "Data do inÃ­cio dos sintomas": "Data do inIcio dos sintomas", "ClassificaÃ§Ã£o Final": "ClassificaCAo Final", "EvoluÃ§Ã£o Caso": "EvoluCAo Caso", "Teste SorolÃ³gico": "Teste SorolOgico", "CNES NotificaÃ§Ã£o": "CNES NotificaCAo"}, inplace=True) #Renomeia as colunas substituindo os caracteres acima citados
    return tabela

def limpaAcentos(tabela):
    """
    Parameters
    ----------
    tabela : DataFrame
        Tabela Assessor

    Returns
    -------
    tabela : DataFrame
        A mesma tabela recebida, porém retirando acentos e cedilhas.

    """
    tabela.replace(to_replace={"Á": "A", "á": "A", "Â": "A", "â": "A", "Ã": "A", "ã": "A",
                               "É": "E", "é": "E", "Ê": "E", "ê": "E",
                               "Í": "I", "í": "I", "Î": "I", "î": "I",
                               "Ó": "O", "ó": "O", "Ô": "O", "ô": "O", "Õ": "O", "õ": "O",
                               "Ú": "U", "ú": "U", "Û": "U", "û": "U",
                               "Ç": "C", "ç": "C"}, inplace=True, regex=True) #Regras de substituição por caractere
    tabela.rename(columns={"Cód. Paciente": "COd. Paciente", "Data da Notificação": "Data da NotificaCAo", "Situação": "SituaCAo"}, inplace=True) #Renomeia as colunas substituindo os caracteres acima citados
    return tabela

def appendTabelaAuxiliar(tabelaIncorreta, row, motivo, tabelaEsus):
    """
    Parameters
    ----------
    tabela : List
        Lista de linhas (dict) incorretas, que posteriormente é convertida em DataFrame
    row : Itertuple
        A linha em questão da tabela eSUS a ser inserida na lista de linhas incorretas
    motivo : String
        Mensagem de motivo pelo qual está sendo considerada incorreta

    Returns
    -------
    None.

    """    
    aux = {"NUmero da NotificaCAo": row[paramColunaNumNotEsus], "Nome Completo": row[paramColunaNomeEsus], "CPF": row[paramColunaCpfEsus], "CNS": row[paramColunaCnsEsus], "Data de Nascimento": row[paramColunaDataNascEsus], "Nome Completo da MAe": row[paramColunaNomeMaeEsus], "Sexo": row[paramColunaSexoEsus], "Telefone 1": row[paramColunaTel1Esus], "Telefone 2": row[paramColunaTel2Esus], "CEP": row[paramColunaCepEsus], "Logradouro": row[paramColunaLogradouroEsus], "NUmero": row[paramColunaNumEndEsus], "Complemento": row[paramColunaCompEndEsus], "Bairro": row[paramColunaBairroEndEsus], "RaCa/Cor": row[paramColunaRacaEsus], "Data da NotificaCAo": row[paramColunaDataNotifEsus], "Data do inIcio dos sintomas": row[paramColunaDataSintomasEsus], "CNES NotificaCAo": row[paramColunaCnesEsus], "Notificante Email": row[paramColunaEmailEsus], "Notificante Nome Completo": row[paramColunaNotifNomeEsus], "SITUACAO": tabelaEsus.at[row.Index, "SITUACAO"], "IDS ASSESSOR": tabelaEsus.at[row.Index, "IDS ASSESSOR"]} #Criação de novo Dict a ser inserido na lista de incorretos
    if motivo is not None:
        aux['Motivo'] = motivo #Adiciona o motivo ao Dict
    tabelaIncorreta.append(aux) #Insere o Dict na lista

def formataImpressao(tabela, tipo):
    """
    Parameters
    ----------
    tabela : DataFrame
        Tabela eSUS
    tipo : String
        String que indica que tipo de tabela é, para saber como tratar. Valores: CORRETO ou INCORRETO

    Returns
    -------
    DataFrame
        A mesma DataFrame recebida, porém com as colunas reorganizadas e a coluna do Número de Notificação reformatada.

    """
    tabela["NUmero da NotificaCAo"] = tabela["NUmero da NotificaCAo"].astype(str)
    if tipo == "CORRETO": #Colunas em ordem caso seja tabela de Corretos
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas", "Nome Completo da MAe", "Sexo", "Telefone 1", "Telefone 2", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "CNES NotificaCAo", "Notificante Email", "Notificante Nome Completo"]]
    elif tipo == "INCORRETO": #Colunas em ordem caso seja tabela de Incorretos (ou seja, com adição do motivo ao final)
        return tabela[["NUmero da NotificaCAo", "IDS ASSESSOR", "Nome Completo", "Data de Nascimento", "CPF", "CNS", "Data da NotificaCAo", "Data do inIcio dos sintomas",  "Nome Completo da MAe", "Sexo", "Telefone 1", "Telefone 2", "CEP", "Logradouro", "NUmero", "Complemento", "Bairro", "RaCa/Cor", "CNES NotificaCAo", "Notificante Email", "Notificante Nome Completo", "SITUACAO", "Motivo"]]
    else: #A mesma tabela caso o parâmetro "Tipo" não tenha um dos valores preconizados
        return tabela

def descobreSituacao(tabela, row):
    """
    Parameters
    ----------
    tabela : DataFrame
        Tabela eSUS
    row : Itertuple
        Linha da Tabela eSUS em questão (a ser investigada)
        
    Description
    -----------
    Determina a situação da linha (row) e também adiciona este resultado à coluna SITUAÇÃO da Tabela eSUS

    Returns
    -------
    String
        Palavra-chave que determina qual é a situação da linha em questão (POSITIVO, NEGATIVO, SUSPEITO ou INCONCLUSIVO)

    """
    # if row[paramColunaTipoTesteEsus] == "RT-PCR" or row[paramColunaTipoTesteEsus] == "TESTE RAPIDO - ANTIGENO": #Checa se o tipo do teste é PCR ou TR-Antígeno
    #     if row[paramColunaResultadoPCREsus] == "Positivo" or ((row[paramColunaResultadoIgmEsus] == "Reagente" or row[paramColunaResultadoTotaisEsus] == "Reagente") and row[paramColunaAssintomaticoEsus] == "NAo"): #Se resultado PCR/Rápido for positivo, considera positivo. Se Resultado IGM ou Total for Positivo E Assintomático for Não, considera Positivo.
    #         tabela.at[row.Index, "SITUACAO"] = "POSITIVO"
    #         return "POSITIVO"
    #     elif row[paramColunaResultadoPCREsus] == "Negativo" or row[paramColunaResultadoIgmEsus] == "NAo Reagente" or row[paramColunaResultadoTotaisEsus] == "NAo Reagente": #Se qualquer das colunas de resultado forem Negativo ou Não Reagente, considera Negativo
    #         tabela.at[row.Index, "SITUACAO"] = "NEGATIVO"
    #         return "NEGATIVO"
    #     elif pd.isnull(row[paramColunaResultadoIgmEsus]) and pd.isnull(row[paramColunaResultadoTotaisEsus]): #Se Resultado IgM e Resultado Totais estiverem em branco, considera Suspeito. Para cair nessa condição, é porque Resultado PCR/Rapido com certeza está em branco (pois já testou para Positivo e Negativo)
    #         tabela.at[row.Index, "SITUACAO"] = "SUSPEITO"
    #         return "SUSPEITO"
    #     else: #Se nenhuma das condições acima forem verdadeiras, considera Inconclusivo
    #         tabela.at[row.Index, "SITUACAO"] = "INCONCLUSIVO"
    #         return "INCONCLUSIVO"
    # elif row[paramColunaTipoTesteEsus] == "TESTE RAPIDO - ANTICORPO" or pd.isnull(row[paramColunaTipoTesteEsus]): #Checa se o tipo de teste é TR-Anticorpo
    #     if (row[paramColunaResultadoPCREsus] == "Positivo" or row[paramColunaResultadoTotaisEsus] == "Reagente" or row[paramColunaResultadoIgmEsus] == "Reagente") and row[paramColunaAssintomaticoEsus] == "NAo": #Se qualquer uma das colunas de resultado for Positivo ou Reagente e Assintomático for Não, considera Positivo.
    #         tabela.at[row.Index, "SITUACAO"] = "POSITIVO"
    #         return "POSITIVO"
    #     elif row[paramColunaResultadoPCREsus] == "Negativo" or row[paramColunaResultadoIgmEsus] == "NAo Reagente" or row[paramColunaResultadoTotaisEsus] == "NAo Reagente": #Se qualquer uma das colunas de resultado for Negativo ou Não Reagente, considera Negativo
    #         tabela.at[row.Index, "SITUACAO"] = "NEGATIVO"
    #         return "NEGATIVO"
    #     elif pd.isnull(row[paramColunaResultadoPCREsus]) and pd.isnull(row[paramColunaResultadoIgmEsus]) and pd.isnull(row[paramColunaResultadoTotaisEsus]): #Se as 3 colunas de resultado estiverem em branco, considera Suspeito
    #         tabela.at[row.Index, "SITUACAO"] = "SUSPEITO"
    #         return "SUSPEITO"
    #     else: #Se nenhuma das condições acima forem verdadeiras, considera Inconclusivo.
    #         tabela.at[row.Index, "SITUACAO"] = "INCONCLUSIVO"
    #         return "INCONCLUSIVO"
    # else: #Se não for nenhum desses tipos de teste, considera Inconclusivo
    #     tabela.at[row.Index, "SITUACAO"] = "INCONCLUSIVO"
    #     return "INCONCLUSIVO"
    if retornaResultado(row[paramColunaResultadoPCREsus]) == "Positivo" or retornaResultado(row[paramColunaResultadoAntigenoEsus]) == "Positivo" or (row[paramColunaAssintomaticoEsus] == "NAo" and (retornaResultado(row[paramColunaResultadoIgaEsus]) == "Positivo" or retornaResultado(row[paramColunaResultadoSorologicoIgmEsus]) == "Positivo" or retornaResultado(row[paramColunaResultadoAnticorpoIgmEsus]) == "Positivo" or retornaResultado(row[paramColunaResultadoAnticorposTotaisEsus]) == "Positivo" or retornaResultado(row[paramColunaResultadoAnticorpoIgmEsus]) == "Positivo")):
        tabela.at[row.Index, "SITUACAO"] = "POSITIVO"
        return "POSITIVO"
    elif retornaResultado(row[paramColunaResultadoPCREsus]) == "Negativo" or retornaResultado(row[paramColunaResultadoAntigenoEsus]) == "Negativo" or (row[paramColunaAssintomaticoEsus] == "NAo" and (retornaResultado(row[paramColunaResultadoIgaEsus]) == "Negativo" or retornaResultado(row[paramColunaResultadoSorologicoIgmEsus]) == "Negativo" or retornaResultado(row[paramColunaResultadoAnticorpoIgmEsus]) == "Negativo" or retornaResultado(row[paramColunaResultadoAnticorposTotaisEsus]) == "Negativo" or retornaResultado(row[paramColunaResultadoAnticorpoIgmEsus]) == "Negativo")):
        tabela.at[row.Index, "SITUACAO"] = "NEGATIVO"
        return "NEGATIVO"
    elif (not pd.isnull(row[paramColunaResultadoTesteOutros1])) or (not pd.isnull(row[paramColunaResultadoTesteOutros2])) or (not pd.isnull(row[paramColunaResultadoTesteOutros3])):
        tabela.at[row.Index, "SITUACAO"] = "INCONCLUSIVO"
        return "INCONCLUSIVO"
    else:
        tabela.at[row.Index, "SITUACAO"] = "SUSPEITO"
        return "SUSPEITO"
        
def trataSuspeito(row, notifAssessor, tabela, tabelaIncorreta):
    """
    Parameters
    ----------
    row : Itertuple
        Linha da Tabela eSUS
    notifAssessor : DataFrame
        DataFrame com todas as notificações da tabela Assessor relativas à essa linha da tabela eSUS
    tabela : DataFrame
        Referência à tabela eSUS, usada para eliminar as linhas que já foram transcritas à tabela Incorreta.
    tabelaIncorreta : List
        Lista de Dicts, onde cada Dict é uma linha da tabela eSUS que foi considerada incorreta

    Returns
    -------
    None.

    """
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação confirmada no Assessor, não é suspeita.
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO", tabela)
        tabela.drop(row.Index, inplace=True)
        return
    if "SUSPEITA" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação suspeita no Assessor, não é nova suspeita
        appendTabelaAuxiliar(tabelaIncorreta, row, "SUSPEITO EM ABERTO", tabela)
        tabela.drop(row.Index, inplace=True)
        return
    if "NEGATIVO" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação negativa no Assessor, verifica o parâmetro de data
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all') #Lista de todas as notificações NEGATIVAS da Tabela Assessor
        for rowNegativo in negativosAssessor.itertuples(): #Laço para cada notificação negativa acima citada
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo): #Se a Data da Notificação no Assessor for (paramDiasNegativo) dias antes ou qualquer quantidade de dias depois da Notificação no eSUS, não considera suspeita
                appendTabelaAuxiliar(tabelaIncorreta, row, "NEGATIVO DENTRO DO PERIODO", tabela)
                tabela.drop(row.Index, inplace=True)
                return
    if "DESCARTADO" in notifAssessor["SituaCAo"].values: #Mesma regra de negativo
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):                
                appendTabelaAuxiliar(tabelaIncorreta, row, "DESCARTADO DENTRO DO PERIODO", tabela)
                tabela.drop(row.Index, inplace=True)
                return

def trataPositivo(row, notifAssessor, tabela, tabelaIncorreta):
    """
    Parameters
    ----------
    row : Itertuple
        Linha da Tabela eSUS
    notifAssessor : DataFrame
        DataFrame com todas as notificações da tabela Assessor relativas à essa linha da tabela eSUS
    tabela : DataFrame
        Referência à tabela eSUS, usada para eliminar as linhas que já foram transcritas à tabela Incorreta.
    tabelaIncorreta : List
        Lista de Dicts, onde cada Dict é uma linha da tabela eSUS que foi considerada incorreta

    Returns
    -------
    None.

    """
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação confirmada no Assessor, não é novo Confirmado
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO", tabela)
        tabela.drop(row.Index, inplace=True)
        return

def trataNegativo(row, notifAssessor, tabela, tabelaIncorreta):
    """
    Parameters
    ----------
    row : Itertuple
        Linha da Tabela eSUS
    notifAssessor : DataFrame
        DataFrame com todas as notificações da tabela Assessor relativas à essa linha da tabela eSUS
    tabela : DataFrame
        Referência à tabela eSUS, usada para eliminar as linhas que já foram transcritas à tabela Incorreta.
    tabelaIncorreta : List
        Lista de Dicts, onde cada Dict é uma linha da tabela eSUS que foi considerada incorreta

    Returns
    -------
    None.

    """
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação confirmada no Assessor, não é negativo
        appendTabelaAuxiliar(tabelaIncorreta, row, "CONFIRMADO", tabela)
        tabela.drop(row.Index, inplace=True)
        return   
    if "NEGATIVO" in notifAssessor["SituaCAo"].values: #Se houver alguma notificação negativo no Assessor, verifica parâmetro de data
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo): #Se a Data da Notificação no Assessor for (paramDiasNegativo) dias antes ou qualquer quantidade de dias depois da Notificação no eSUS, não considera negativo
                appendTabelaAuxiliar(tabelaIncorreta, row, "NEGATIVO DENTRO DO PERIODO", tabela)
                tabela.drop(row.Index, inplace=True)
                return
    if "DESCARTADO" in notifAssessor["SituaCAo"].values: #Mesma regra do negativo
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                appendTabelaAuxiliar(tabelaIncorreta, row, "DESCARTADO DENTRO DO PERIODO", tabela)
                tabela.drop(row.Index, inplace=True)
                return
    if "SUSPEITA" in notifAssessor["SituaCAo"].values: #Vê se tem suspeita aberta para esse negativo
        suspeitosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "SUSPEITA").dropna(how='all')
        for rowSuspeito in suspeitosAssessor.itertuples():
            dataNotifAssessor = rowSuspeito[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifEsus >= dataNotifAssessor:
                tabela.at[row.Index, "SITUACAO"] = "NEGATIVO C/ SUSPEITA"
                return
    tabela.at[row.Index, "SITUACAO"] = "NEGATIVO S/ SUSPEITA"
    return

def retornaResultado(stringResultado):
    stringsPositivo = ["DetectAvel", "Reagente"]
    stringsNegativo = ["NAo detectAvel", "NAo Reagente"]
    
    if stringResultado in stringsPositivo:
        return "Positivo"
    elif stringResultado in stringsNegativo:
        return "Negativo"
    else:
        return None

def converteFiltro(tabela, lblCnes, lblEmail, listaFiltro):
    """
    Parameters
    ----------
    tabela : DataFrame
        Tabela do eSUS (Correta ou Incorreta) a ser filtrada
    lblCnes : String
        String que representa rótulo da coluna que possui o CNES da unidade
    lblEmail : String
        String que representa rótulo da coluna que possui o e-mail da unidade
    listaFiltro : Dict
        Dict com os CNES e Emails a serem considerados de uma certa unidade ou grupo de unidades para o filtro

    Returns
    -------
    filtroFinal : Series
        Filtro a ser usado para filtrar todas as notificações da Tabela eSUS que se encaixam em ListaFiltro (filtro para ser usado com função .where())

    """
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

def segundoEmHora(segundos):
    """
    Parameters
    ----------
    segundos : Float
        Quantidade de segundos a serem convertidos em String de Tempo

    Returns
    -------
    stringTempo : String
        String representando os segundos recebidos em Hora:Minuto:Segundo

    """
    horas = int(segundos / 3600) if int(segundos / 3600) >= 10 else "0" + str(int(segundos / 3600))
    minutos = int((segundos % 3600) / 60) if int((segundos % 3600) / 60) >= 10 else "0" + str(int((segundos % 3600) / 60))
    segundos = int((segundos % 3600) % 60) if int((segundos % 3600) % 60) >= 10 else "0" + str(int((segundos % 3600) % 60))
    stringTempo = str(horas) + ":" + str(minutos) + ":" + str(segundos)
    return stringTempo

def printTabela(nomeArquivo, tabelaCorreta, tabelaIncorreta):
    """
    Parameters
    ----------
    nomeArquivo : String
        String que indica o nome do arquivo a ser gerado
    tabelaCorreta : DataFrame
        Tabela com os registros CORRETOS a serem divididos em Positivos, Negativos e Inconclusivos e posteriormente serem impressos em .xlsx
    tabelaIncorreta : DataFrame
        Tabela com os registros INCORRETOS a serem impressos no xlsx

    Returns
    -------
    None.

    """
    with pd.ExcelWriter(nomeArquivo + ' ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:
        tabelaPos = tabelaCorreta.where(tabelaCorreta["SITUACAO"] == "POSITIVO").dropna(how='all')
        tabelaPos = formataImpressao(tabelaPos, "CORRETO")
        tabelaPos.to_excel(writer, "Positivos", index=False)    
        
        tabelaNegCSus = tabelaCorreta.where(tabelaCorreta["SITUACAO"] == "NEGATIVO C/ SUSPEITA").dropna(how='all')
        tabelaNegCSus = formataImpressao(tabelaNegCSus, "CORRETO")
        tabelaNegCSus.to_excel(writer, "Negativos com Suspeita", index=False)
        
        tabelaNegSSus = tabelaCorreta.where(tabelaCorreta["SITUACAO"] == "NEGATIVO S/ SUSPEITA").dropna(how='all')
        tabelaNegSSus = formataImpressao(tabelaNegSSus, "CORRETO")
        tabelaNegSSus.to_excel(writer, "Negativos sem Suspeita", index=False)
        
        tabelaInconc = tabelaCorreta.where(tabelaCorreta["SITUACAO"] == "INCONCLUSIVO").dropna(how='all')
        tabelaInconc = formataImpressao(tabelaInconc, "CORRETO")
        tabelaInconc.to_excel(writer, "Inconclusivos", index=False)
            
        tabelaIncorreta = formataImpressao(tabelaIncorreta, "INCORRETO")
        tabelaIncorreta.to_excel(writer, "Incorretos", index=False)

paramDiasAtras = 16 #Parâmetro que define quantos dias atrás ele considera na planilha de suspeitos e monitoramento (ex: notificações de até X dias atrás serão analisadas, antes disso serão ignoradas)
paramDiasSuspeitos = 15 #Parâmetro que define quantos dias atrás uma notificacao de suspeita deve ser considerada "a mesma" notificacao. Acima desse parâmetro deve ser considerada uma nova notificacao
paramDiasNegativo = 15 #Parâmetro que define quantos dias atrás uma notificacao de negativo deve ser considerada "a mesma" notificacao. Acima desse parâmetro deve ser considerada uma nova notificacao (agora também usada para descartados)

paramColunaNumNotEsus = 1
paramColunaTel1Esus = 4
paramColunaSexoEsus = 6
paramColunaEstadoTesteEsus = 8
paramColunaCpfEsus = 9
paramColunaClassifFinalEsus = 10
paramColunaCnsEsus = 11
paramColunaDataNascEsus = 12
paramColunaNomeMaeEsus = 14
paramColunaTel2Esus = 16
paramColunaNomeEsus = 17
paramColunaCepEsus = 20
paramColunaLogradouroEsus = 21
paramColunaNumEndEsus = 22
paramColunaCompEndEsus = 23
paramColunaBairroEndEsus = 24
paramColunaRacaEsus = 25
paramColunaAssintomaticoEsus = 47
paramColunaDataNotifEsus = 48
paramColunaDataSintomasEsus = 50
paramColunaResultadoPCREsus = 72
paramColunaResultadoIgaEsus = 78
paramColunaResultadoSorologicoIgmEsus = 81
paramColunaResultadoAnticorposTotaisEsus = 87
paramColunaResultadoAnticorpoIgmEsus = 90
paramColunaResultadoAntigenoEsus = 96
paramColunaResultadoTesteOutros1 = 102
paramColunaResultadoTesteOutros2 = 108
paramColunaResultadoTesteOutros3 = 114
paramColunaCnesEsus = 121
paramColunaEmailEsus = 123
paramColunaNotifNomeEsus = 124
paramColunaSituacaoEsus = 126
paramColunaIdsEsus = 127


flagsUnidades = {}
flagsUnidades["stc"] = False
flagsUnidades["pio"] = False
flagsUnidades["hans"] = False
flagsUnidades["int"] = False
flagsUnidades["ext"] = False

mapOpcaoToUnidade = {}
mapOpcaoToUnidade["1"] = "stc"
mapOpcaoToUnidade["2"] = "pio"
mapOpcaoToUnidade["3"] = "hans"
mapOpcaoToUnidade["4"] = "int"
mapOpcaoToUnidade["5"] = "ext"

lblCnes = "CNES NotificaCAo"
lblEmail = "Notificante Email"

paramColunaIdAssessor = 2
paramColunaDataNotifAssessor = 11 #Parâmetro igual ao paramColunaDataNotifEsus porém para o Assessor
paramDataAtual = datetime.today() #Parâmetro que define qual é a data atual para o script fazer a comparacao dos dias para trás (padrão: datetime.today() = data atual do sistema)

print("Data de hoje: " + paramDataAtual.strftime("%d/%m/%Y"))
print("Data de " + str(paramDiasAtras) + " dias atrás: " + (paramDataAtual - timedelta(days=paramDiasAtras)).strftime("%d/%m/%Y"))

regexOpcoes = "\d[\,\s\d]*$"
regex = re.compile(regexOpcoes, re.IGNORECASE)

while True:
    print("1 - Santa Casa\n2 - Pio XII\n3 - HANS\n4 - Internas\n5 - Externas\n6 - Todas\n")
    print("Insira as opções desejadas, separadas por vírgula (ex: 1, 2, 3): ", end="")
    opcoes = input()
    
    if regex.match(opcoes):
        opcoes = opcoes.replace(' ', '')
        opcoes = opcoes.split(',')
        break
    else:
        print("Opções inválidas, por favor tente novamente.\n")

if "6" in opcoes:
    for key in flagsUnidades.keys():
        flagsUnidades[key] = True
else:
    for opcao in opcoes:
        flagsUnidades[mapOpcaoToUnidade[opcao]] = True
        
print("Executar Realmente Suspeitos? (S/N): ")
tiraSuspeitas = input().strip()

print("Iniciando Processamento...")
 
start = timer()

tabelaTotalEsus = pd.read_excel("lista total esus.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_, 'Nome Completo': np.unicode_}).sort_values(by="Data da NotificaÃ§Ã£o", ignore_index=True) #Lista total das notificações do eSus
tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS'] = formataCpfCns(tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS']) #Formata as colunas de CNS e CPF conforme a regra de negócio
tabelaTotalEsus['Nome Completo'] = tabelaTotalEsus['Nome Completo'].str.upper() #Transforma a coluna de nome em Caixa Alta
tabelaTotalEsus = limpaUtf8(tabelaTotalEsus) #Limpa os caracteres que vem "bugados" do eSUS por conta de acentos e 'ç'
tabelaTotalEsus = tabelaTotalEsus.where(tabelaTotalEsus['MunicIpio de ResidEncia'] == "Barretos").dropna(how='all') #Filtra apenas pelo município de residência Barretos
#tabelaTotalEsus = tabelaTotalEsus.where(tabelaTotalEsus['EvoluCAo Caso'] != "Cancelado").dropna(how='all') #Retira as notificações canceladas
#tabelaTotalEsus = tabelaTotalEsus.where((tabelaTotalEsus['Teste SorolOgico'] != "Anticorpos Totais") & (tabelaTotalEsus['Teste SorolOgico'] != "IgG, Anticorpos Totais") & (tabelaTotalEsus['Teste SorolOgico'] != "Anticorpos Totais, IgG")).dropna(how='all')
tabelaTotalEsus.insert(len(tabelaTotalEsus.columns), "SITUACAO", None)
tabelaTotalEsus.insert(len(tabelaTotalEsus.columns), "IDS ASSESSOR", None)

filtrosValidos = []

if flagsUnidades['stc']:
    listaFiltroStc = {"cnes": [2092611], "email": []}
    filtroStc = converteFiltro(tabelaTotalEsus, lblCnes, lblEmail, listaFiltroStc)
    filtrosValidos.append(filtroStc)
if flagsUnidades['pio']:
    listaFiltroPio = {"cnes": [2090236], "email": []}
    filtroPio = converteFiltro(tabelaTotalEsus, lblCnes, lblEmail, listaFiltroPio)
    filtrosValidos.append(filtroPio)
if flagsUnidades['hans']:
    listaFiltroHans = {"cnes": [9662561], "email": []}
    filtroHans = converteFiltro(tabelaTotalEsus, lblCnes, lblEmail, listaFiltroHans)
    filtrosValidos.append(filtroHans)
if flagsUnidades['ext']:
    listaFiltroExternas = {"cnes": [9344209, 2052970, 9567771, 9371508, 3549208, 7909837, 3142531, 5124700, 6316344, 419907, 2074176, 7653166],
                      "email": ["maicon.douglas19@hotmail.com", "yaah_cardoso@hotmail.com", "robertamap@yahoo.com.br", "barretos2@dpsp.com.br", "marcela@ltaseguranca.com.br"]}
    filtroExternas = converteFiltro(tabelaTotalEsus, lblCnes, lblEmail, listaFiltroExternas)
    filtrosValidos.append(filtroExternas)
if flagsUnidades['int']:
    listaFiltroInternas = {"cnes": [2035731, 2048736, 2048744, 2053314, 2053062, 2061473, 2064081, 2064103, 2093642, 2093650, 2784572, 2784580, 2784599, 5562325, 5562333, 7035861, 7122217, 7565577, 7585020, 2053306],
                      "email": ["machado.vl@bol.com.br", "marcioli@bol.com.br", "marcosnascimento32@outlook.com"]}
    filtroInternas = converteFiltro(tabelaTotalEsus, lblCnes, lblEmail, listaFiltroInternas)
    filtrosValidos.append(filtroInternas)

filtroUnidades = pd.Series(dtype=object)
for filtro in filtrosValidos:
    if filtroUnidades.empty:
        filtroUnidades = filtro
    else:
        filtroUnidades = (filtroUnidades) | (filtro)

tabelaTotalEsus = tabelaTotalEsus.where(filtroUnidades).dropna(how='all')
qtdNotif = len(tabelaTotalEsus)
porcAtual = 0

endReadEsus = timer()
print("Terminei de ler a tabela do eSUS aos " + str(endReadEsus - start) + " segundos.")

#Puxa a tabela do Assessor
tabelaTotalAssessor = pd.read_excel("lista total assessor.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by=["Cód. Paciente", "Data da Notificação"]) #Lê a tabela do Assessor e classifica por ID do paciente e Data da Notificação

if tiraSuspeitas != "N" and tiraSuspeitas != "n":
    tabelaTotalAssessor = realmenteSuspeitos(tabelaTotalAssessor)
    endRealmenteSuspeitos = timer()
    print("Terminei de tirar os 'Realmente Suspeitos' aos " + str(endRealmenteSuspeitos - start) + " segundos.")

tabelaTotalAssessor = limpaAcentos(tabelaTotalAssessor) #Limpa acentos e 'ç' da tabela do Assessor

endReadAssessor = timer()
print("Terminei de ler a tabela do Assessor aos " + str(endReadAssessor - start) + " segundos.")

tabelaTotalEsusIncorreta = []
countLinhas = 0
somaVeloc = 0
velocMedia = 0
porcentil = round(qtdNotif / 100)
startProcessamento = timer()

for row in tabelaTotalEsus.itertuples():
    situacao = descobreSituacao(tabelaTotalEsus, row)
    if situacao != "INCONCLUSIVO" and situacao != "SUSPEITO":
        notifAssessor = acharId(row, tabelaTotalAssessor, tabelaTotalEsus)
    
    if situacao == "POSITIVO":
        trataPositivo(row, notifAssessor, tabelaTotalEsus, tabelaTotalEsusIncorreta)
    elif situacao == "NEGATIVO":
        trataNegativo(row, notifAssessor, tabelaTotalEsus, tabelaTotalEsusIncorreta)
        
    countLinhas += 1
    if countLinhas >= (porcAtual + 1) * porcentil:
        porcAtual += 1
        os.system('cls')
        porcTimer = timer()
        tempoAtual = porcTimer - startProcessamento
        velocEstimada = countLinhas / tempoAtual
        somaVeloc += velocEstimada
        velocMedia = somaVeloc / porcAtual
        tempoEstimadoTotal = (qtdNotif - countLinhas) / velocEstimada
        stringTempoEstimado = segundoEmHora(tempoEstimadoTotal)
        print("Concluido: " + str(porcAtual) + "% aos " + "{:.0f}".format(tempoAtual) + " segundos (desde o início do laço).\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.\nVelocidade: " + "{:.2f}".format(velocEstimada) + " linhas por segundo.\nVelocidade media: " + "{:.2f}".format(velocMedia) + " linhas por segundo.\nTempo estimado: " + stringTempoEstimado + " para conlusão\n\n")


endLaco = timer()
print("Terminei de processar o eSUS aos " + str(endLaco - start) + " segundos.\nTotal de " + segundoEmHora(endLaco - start) + " de processamento.\nVelocidade média de " + "{:.2f}".format(velocMedia) + " linhas por segundo.")

tabelaTotalEsusIncorreta = pd.DataFrame(tabelaTotalEsusIncorreta)

with pd.ExcelWriter('Debug eSUS ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerDebug:
    tabelaTotalEsus.to_excel(writerDebug, "Corretos", index=False)
    tabelaTotalEsusIncorreta.to_excel(writerDebug, "Incorretos", index=False)

if flagsUnidades['stc']:
    tabelaStc = tabelaTotalEsus.where(filtroStc).dropna(how='all')
    filtroStcInc = converteFiltro(tabelaTotalEsusIncorreta, lblCnes, lblEmail, listaFiltroStc)
    tabelaStcIncorreta = tabelaTotalEsusIncorreta.where(filtroStcInc).dropna(how='all')
    printTabela("Santa Casa", tabelaStc, tabelaStcIncorreta)
if flagsUnidades['pio']:
    tabelaPio = tabelaTotalEsus.where(filtroPio).dropna(how='all')
    filtroPioInc = converteFiltro(tabelaTotalEsusIncorreta, lblCnes, lblEmail, listaFiltroPio)
    tabelaPioIncorreta = tabelaTotalEsusIncorreta.where(filtroPioInc).dropna(how='all')
    printTabela("Pio XII", tabelaPio, tabelaPioIncorreta)
if flagsUnidades['hans']:
    tabelaHans = tabelaTotalEsus.where(filtroHans).dropna(how='all')
    filtroHansInc = converteFiltro(tabelaTotalEsusIncorreta, lblCnes, lblEmail, listaFiltroHans)
    tabelaHansIncorreta = tabelaTotalEsusIncorreta.where(filtroHansInc).dropna(how='all')
    printTabela("HANS", tabelaHans, tabelaHansIncorreta)
if flagsUnidades['ext']:
    tabelaExternas = tabelaTotalEsus.where(filtroExternas).dropna(how='all')
    filtroExternasInc = converteFiltro(tabelaTotalEsusIncorreta, lblCnes, lblEmail, listaFiltroExternas)
    tabelaExternasIncorreta = tabelaTotalEsusIncorreta.where(filtroExternasInc).dropna(how='all')
    printTabela("Externas", tabelaExternas, tabelaExternasIncorreta)
if flagsUnidades['int']:
    tabelaInternas = tabelaTotalEsus.where(filtroInternas).dropna(how='all')
    filtroInternasInc = converteFiltro(tabelaTotalEsusIncorreta, lblCnes, lblEmail, listaFiltroInternas)
    tabelaInternasIncorreta = tabelaTotalEsusIncorreta.where(filtroInternasInc).dropna(how='all')
    printTabela("Internas", tabelaInternas, tabelaInternasIncorreta)