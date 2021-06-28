# -*- coding: utf-8 -*-
"""
Created on Tue Oct 13 17:02:38 2020

@author: Info
"""


import pandas as pd
from datetime import timedelta
from datetime import datetime
import numpy as np
import os

def formataCpfCns(colunaCpf, colunaCns):
    for index, value in colunaCpf.items():
        if(len(str(colunaCns[index])) != 15):
            colunaCns[index] = None
        if(pd.isnull(value)):
            colunaCpf[index] = None
            continue
        value = value.strip()
        value = value.replace(".", "")
        value = value.replace("-", "")
        while len(value) < 11:
            value = "0" + value
        colunaCpf[index] = str(value)
        # if(len(colunaCpf[index]) < 14):
        #     colunaCpf[index] = value[:3] + '.' + value[3:6] + "." + value[6:9] + "-" + value[9:]
    return colunaCpf, colunaCns

def acharId(Cpf, Cns, tabela):
    #print("TESTE - DN param " + str(Dn) + "\nPrimeira data da tabela: " + str(tabela['Data Nasc.'][0]))    
    #print("Entrei em acharId, meu tipo do CNS é ")
    #print(type(Cns))   
    if Cpf is None and Cns is None:    
        return None
    
    if(Cpf is not None):
        filtro1 = tabela['CPF'] == Cpf.strip()
    else:
        filtro1 = False
        
    if(Cns is not None):
        filtro2 = tabela['CNS'] == Cns.strip()
    else:
        filtro2 = False
        
    teste = tabela.where(filtro1 | filtro2).dropna(how='all')
    return teste   

def limpaAcentosAssessor(tabela): #Função que limpa os caracteres especiais da tabela do Assessor
    tabela.replace(to_replace={"Á": "A", "á": "A", "Â": "A", "â": "A", "Ã": "A", "ã": "A",
                               "É": "E", "é": "E", "Ê": "E", "ê": "E",
                               "Í": "I", "í": "I", "Î": "I", "î": "I",
                               "Ó": "O", "ó": "O", "Ô": "O", "ô": "O", "Õ": "O", "õ": "O",
                               "Ú": "U", "ú": "U", "Û": "U", "û": "U",
                               "Ç": "C", "ç": "C"}, inplace=True, regex=True)
    tabela.rename(columns={"Cód. Paciente": "COd. Paciente", "Data da Notificação": "Data da NotificaCAo", "Situação": "SituaCAo"}, inplace=True)    
    return tabela

def limpaAcentosGal(tabela): #Função que limpa os caracteres especiais da tabela do GAL
    tabela.replace(to_replace={"Á": "A", "á": "A", "Â": "A", "â": "A", "Ã": "A", "ã": "A",
                               "É": "E", "é": "E", "Ê": "E", "ê": "E",
                               "Í": "I", "í": "I", "Î": "I", "î": "I",
                               "Ó": "O", "ó": "O", "Ô": "O", "ô": "O", "Õ": "O", "õ": "O",
                               "Ú": "U", "ú": "U", "Û": "U", "û": "U",
                               "Ç": "C", "ç": "C"}, inplace=True, regex=True)
    tabela.rename(columns={"Requisição": "RequisiCAo", "Mun. Residência": "Mun. ResidEncia", "Coronavírus SARS-CoV2": "CoronavIrus SARS-CoV2"}, inplace=True)    
    return tabela

def limpaEspacosGal(tabela): #Função que limpa os espaços desnecessários da tabela do GAL
    colunas = ['Metodo', 'Resultado', "CoronavIrus SARS-CoV2"]
    for coluna in colunas:
        tabela[coluna] = tabela[coluna].str.strip()
    
    return tabela

def appendTabelaAuxiliar(tabela, row, motivo, timestamp=False):
    aux = {"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Data de Cadastro": row[paramColunaDataCadGal], "Requisitante": row.Requisitante, "Mun. Residencia": row[paramColunaMunResGal], "Resultado": row.Resultado, "CoronavIrus SARS-CoV2": row[paramColunaCoronavirusGal]}
    if motivo is not None:
        aux['Motivo'] = motivo
    if timestamp:
        aux['Data de Insercao'] = paramDataAtual
    tabela.append(aux)
    
def confereRepeticaoSemId(base, row):
    if base is None:
        return False
    if row.RequisiCAo in base["Requisicao"].values:
        return True
    return False
    

#TODO: Tirar a parte da hora da data atual para poder voltar o paramDiasAtras para 15
paramDiasAtras = 16 #Parâmetro que define quantos dias atrás ele considera na planilha de suspeitos e monitoramento (ex: notificações de até X dias atrás serão analisadas, antes disso serão ignoradas)
paramDiasNovaInfeccao = 14 #Parâmetro que define até quantos dias de diferença de outro agravo essa notificação pode ter para ser considerada a mesma infecção
paramDiasAnulaSuspeito = 5 #Parâmetro que define a quantos dias pode haver outra notificação com resultado para anular uma notificação suspeita nova
paramDiasDentroMonitoramento = 15 #Parâmetro que define até quantos dias atrás o paciente é considerado pelo monitoramento
paramDataAtual = datetime.today() #Parâmetro que define qual é a data atual para o script fazer a comparacao dos dias para trás (padrão: datetime.today() = data atual do sistema)

paramColunaDataCadGal = 18
paramColunaMunResGal = 7
paramColunaDataNotifAssessor = 11
paramColunaCoronavirusGal = 25
paramColunaDataLibGal = 20
paramColunaDataResultAssessor = 13
paramColunaStatusGal = 23

tabelaTotalGal = pd.read_excel("lista total gal.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_, 'Requisição': np.unicode_}).sort_values(by="Dt. Cadastro", ignore_index=True) #Lista total das notificações do GAL
tabelaTotalGal['CPF'], tabelaTotalGal['CNS'] = formataCpfCns(tabelaTotalGal['CPF'], tabelaTotalGal['CNS'])
tabelaTotalGal = limpaAcentosGal(tabelaTotalGal)
tabelaTotalGal = limpaEspacosGal(tabelaTotalGal)

filtroSemResultado1 = tabelaTotalGal['Status Exame'] == 'Resultado Liberado'
filtroSemResultado2 = tabelaTotalGal['Resultado'] == ''
filtroSemResultado3 = tabelaTotalGal['CoronavIrus SARS-CoV2'] == ''

tabelaGalSemResultado = tabelaTotalGal.where(filtroSemResultado1 & filtroSemResultado2 & filtroSemResultado3).dropna(how='all')

filtroPositivo1 = tabelaTotalGal['Resultado'] == 'DetectAvel'
filtroPositivo2 = tabelaTotalGal['CoronavIrus SARS-CoV2'] == 'DetectAvel'

tabelaGalPositivos = tabelaTotalGal.where(filtroPositivo1 | filtroPositivo2).dropna(how='all')

tabelaGalRecentes = tabelaTotalGal.where(tabelaTotalGal['Dt. Cadastro'] >= paramDataAtual - timedelta(days=paramDiasAtras)).dropna(how='all')

filtroNegativo1 = tabelaGalRecentes['Resultado'] == 'NAo DetectAvel'
filtroNegativo2 = tabelaGalRecentes['CoronavIrus SARS-CoV2'] == 'NAo DetectAvel'

tabelaGalNegativos = tabelaGalRecentes.where(filtroNegativo1 | filtroNegativo2).dropna(how='all')

tabelaGalNaoRealizados = tabelaGalRecentes.where(tabelaGalRecentes['Status Exame'] == 'Exame nAo-realizado').dropna(how='all')

filtroSuspeito1 = tabelaGalRecentes['Status Exame'] != 'Resultado Liberado'
filtroSuspeito2 = tabelaGalRecentes['Status Exame'] != 'Exame nAo-realizado'

tabelaGalSuspeitos = tabelaGalRecentes.where(filtroSuspeito1 & filtroSuspeito2).dropna(how='all')

tabelaTotalAssessor = pd.read_excel("lista total assessor.xls", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by=["Cód. Paciente", "Data da Notificação"]) #Lê a tabela do Assessor e classifica por ID do paciente e Data da Notificação
tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS'] = formataCpfCns(tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS']) #Formata as colunas de CPF e CNS baseado na regra de negócio
tabelaTotalAssessor = limpaAcentosAssessor(tabelaTotalAssessor) #Limpa acentos e 'ç' da tabela do Assessor

if(os.path.isfile('base repeticao sem dados.xlsx')):
    tabelaRepeticaoSemDados = pd.read_excel('base repeticao sem dados.xlsx', dtype={'Requisicao': np.unicode_})
else:
    tabelaRepeticaoSemDados = None

tabelaGalPositivosFalso = []
tabelaGalNegativosFalso = []
tabelaGalSuspeitosFalso = []
tabelaGalNaoRealizadosFalso = []
tabelaGalInconsistencias = []
baseRepeticaoSemId = []

for row in tabelaGalPositivos.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, tabelaTotalAssessor)
    if(notifAssessor is None):
        #baseRepeticaoSemId.append({"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Requisitante": row.Requisitante})
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalPositivosFalso, row, "Sem dados já inserido anteriormente")
            tabelaGalPositivos.drop(row.Index, inplace=True)
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Positivo", timestamp=True)
        continue
    
    if(row.CPF is None and row.CNS is not None and row.CNS[0] != '7'): #Verifica se tem apenas CNS e não é início 7 pra jogar pra base de repetição
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalPositivosFalso, row, "Apenas CadSUS sem ser início 7 e requisição já inserida anteriormente")
            tabelaGalPositivos.drop(row.Index, inplace=True)
            continue
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Positivo", timestamp=True)
    
    if notifAssessor.empty:
        #print("Positivo sem nada no Assessor, mun. res = " + row[paramColunaMunResGal])
        if(row[paramColunaMunResGal] == "BARRETOS"):
            appendTabelaAuxiliar(tabelaGalInconsistencias, row, "Barretense positivo sem nenhum agravo")
        continue
    
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        positivosAssessor = notifAssessor.where(notifAssessor['SituaCAo'] == 'CONFIRMADO').dropna(how='all')
        for rowPositivo in positivosAssessor.itertuples():
            flagIsWrong = False
            dataCadGal = row[paramColunaDataCadGal]
            dataNotifAssessor = rowPositivo[paramColunaDataNotifAssessor]
            dataLibGal = row[paramColunaDataLibGal]
            dataResultAssessor = rowPositivo[paramColunaDataResultAssessor]
            if (dataNotifAssessor >= dataCadGal - timedelta(days=paramDiasNovaInfeccao)) or (dataResultAssessor >= dataLibGal - timedelta(days=paramDiasNovaInfeccao)):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaGalPositivosFalso.append({"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Requisitante": row.Requisitante, "Motivo": "Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias"})
            appendTabelaAuxiliar(tabelaGalPositivosFalso, row, "Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias")
            tabelaGalPositivos.drop(row.Index, inplace=True)
            continue
    if not "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #tabelaGalInconsistencias.append({"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Requisitante": row.Requisitante, "Motivo": "Nao tem Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias nem Suspeita no Assessor"})
        appendTabelaAuxiliar(tabelaGalInconsistencias, row, "Nao tem Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias nem Suspeita no Assessor")
        
for row in tabelaGalNegativos.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, tabelaTotalAssessor)
    if(notifAssessor is None):
        #baseRepeticaoSemId.append({"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Requisitante": row.Requisitante})
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalNegativosFalso, row, "Sem dados já inserido anteriormente")
            tabelaGalNegativos.drop(row.Index, inplace=True)
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Negativo", timestamp=True)
        continue
    
    if(row.CPF is None and row.CNS is not None and row.CNS[0] != '7'): #Verifica se tem apenas CNS e não é início 7 pra jogar pra base de repetição
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalNegativosFalso, row, "Apenas CadSUS sem ser início 7 e requisição já inserida anteriormente")
            tabelaGalNegativos.drop(row.Index, inplace=True)
            continue
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Negativo", timestamp=True)
    
    if notifAssessor.empty:
        if(row[paramColunaMunResGal] == "BARRETOS"):
            appendTabelaAuxiliar(tabelaGalInconsistencias, row, "Barretense negativo sem nenhum agravo")
        appendTabelaAuxiliar(tabelaGalNegativosFalso, row, "Negativo sem notificacoes no Assessor")
        tabelaGalNegativos.drop(row.Index, inplace=True)
        continue
    
    if "NEGATIVO" in notifAssessor['SituaCAo'].values:
        negativosAssessor = notifAssessor.where(notifAssessor['SituaCAo'] == 'NEGATIVO').dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False
            dataCadGal = row[paramColunaDataCadGal]
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            if dataNotifAssessor >= dataCadGal - timedelta(days=paramDiasNovaInfeccao):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaGalPositivosFalso.append({"Requisicao": row.RequisiCAo, "Paciente": row.Paciente, "CNS": row.CNS, "CPF": row.CPF, "Requisitante": row.Requisitante, "Motivo": "Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias"})
            appendTabelaAuxiliar(tabelaGalNegativosFalso, row, "Negativo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias")
            tabelaGalNegativos.drop(row.Index, inplace=True)
            continue
    if not "SUSPEITA" in notifAssessor["SituaCAo"].values:
        appendTabelaAuxiliar(tabelaGalInconsistencias, row, "Nao tem Positivo dentro do periodo de " + str(paramDiasNovaInfeccao) + " dias nem Suspeita no Assessor")
        appendTabelaAuxiliar(tabelaGalNegativosFalso, row, "Negativo no GAL sem negativo equivalente no Assessor nem suspeita em aberto")
        tabelaGalNegativos.drop(row.Index, inplace=True)
        
for row in tabelaGalSuspeitos.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, tabelaTotalAssessor)
    if(notifAssessor is None):
        if confereRepeticaoSemId(tabelaRepeticaoSemDados, row):
            appendTabelaAuxiliar(tabelaGalSuspeitosFalso, row, "Sem dados já inserido anteriormente")
            tabelaGalSuspeitos.drop(row.Index, inplace=True)
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Suspeito", timestamp=True)
        continue
    
    if(row.CPF is None and row.CNS is not None and row.CNS[0] != '7'): #Verifica se tem apenas CNS e não é início 7 pra jogar pra base de repetição
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalSuspeitosFalso, row, "Apenas CadSUS sem ser início 7 e requisição já inserida anteriormente")
            tabelaGalSuspeitos.drop(row.Index, inplace=True)
            continue
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Suspeito", timestamp=True)
    
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #TODO: Ver se precisa verificar a quantidade de dias da suspeita pra inserir novamente ou não (por fins de monitoramento)
        appendTabelaAuxiliar(tabelaGalSuspeitosFalso, row, "Ja existe suspeita em aberto")
        tabelaGalSuspeitos.drop(row.Index, inplace=True)
        
for row in tabelaGalNaoRealizados.itertuples():
    notifGal = acharId(row.CPF, row.CNS, tabelaGalRecentes)
    if notifGal is None:
        if confereRepeticaoSemId(tabelaRepeticaoSemDados, row):
            appendTabelaAuxiliar(tabelaGalNaoRealizadosFalso, row, "Sem dados já inserido anteriormente")
            tabelaGalNaoRealizados.drop(row.Index, inplace=True)
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Nao realizado", timestamp=True)
        continue
    
    if(row.CPF is None and row.CNS is not None and row.CNS[0] != '7'): #Verifica se tem apenas CNS e não é início 7 pra jogar pra base de repetição
        if(confereRepeticaoSemId(tabelaRepeticaoSemDados, row)):
            appendTabelaAuxiliar(tabelaGalNaoRealizadosFalso, row, "Apenas CadSUS sem ser início 7 e requisição já inserida anteriormente")
            tabelaGalNaoRealizados.drop(row.Index, inplace=True)
            continue
        else:
            appendTabelaAuxiliar(baseRepeticaoSemId, row, "Nao realizado", timestamp=True)
            
    for rowGal in notifGal.itertuples():
        if(rowGal[paramColunaDataCadGal] >= row[paramColunaDataCadGal] and rowGal[paramColunaStatusGal] != "Exame nAo-realizado"):
            appendTabelaAuxiliar(tabelaGalNaoRealizadosFalso, row, "Foi feito um exame posterior a esse")
            tabelaGalNaoRealizados.drop(row.Index, inplace=True)
            break

baseRepeticaoSemId = pd.DataFrame(baseRepeticaoSemId)
if tabelaRepeticaoSemDados is None:
    baseRepeticaoSemId.to_excel("base repeticao sem dados.xlsx")
else:
    tabelaRepeticaoSemDados = tabelaRepeticaoSemDados.append(baseRepeticaoSemId, ignore_index=True)
    tabelaRepeticaoSemDados.to_excel("base repeticao sem dados.xlsx")
    
with pd.ExcelWriter('Gal Monitoramento ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerGalCorreto:
    tabelaGalPositivos.to_excel(writerGalCorreto, "Positivos")
    tabelaGalNegativos.to_excel(writerGalCorreto, "Negativos")
    tabelaGalSuspeitos.to_excel(writerGalCorreto, "Suspeitos")
    tabelaGalNaoRealizados.to_excel(writerGalCorreto, "Não Realizados")
    
with pd.ExcelWriter('Gal Monitoramento INCORRETO ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerGalIncorreto:
    tabelaGalPositivosFalso = pd.DataFrame(tabelaGalPositivosFalso)
    tabelaGalNegativosFalso = pd.DataFrame(tabelaGalNegativosFalso)
    tabelaGalSuspeitosFalso = pd.DataFrame(tabelaGalSuspeitosFalso)
    tabelaGalNaoRealizadosFalso = pd.DataFrame(tabelaGalNaoRealizadosFalso)
    
    tabelaGalPositivosFalso.to_excel(writerGalIncorreto, "Positivos")
    tabelaGalNegativosFalso.to_excel(writerGalIncorreto, "Negativos")
    tabelaGalSuspeitosFalso.to_excel(writerGalIncorreto, "Suspeitos")
    tabelaGalNaoRealizadosFalso.to_excel(writerGalIncorreto, "Não Realizados")
    
tabelaGalInconsistencias = pd.DataFrame(tabelaGalInconsistencias)
tabelaGalInconsistencias.to_excel("Gal Inconsistências " + paramDataAtual.strftime("%d.%m") + '.xlsx')
    