# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 14:58:29 2021

@author: Info
"""

import pandas as pd

def appendTabelaAuxiliar(tabela, row, motivo):
    aux = {"ID DISP.": row[paramColunaIdDispensacao], "DATA DISP.": row[paramColunaDataDispensacao], "COD UNIDADE": row[paramColunaCodUnidadeDispensacao], "UNIDADE": row[paramColunaNomeUnidadeDispensacao], "QTD DISP.": row[paramColunaQtdDispensacao], "ID. PACIENTE": row[paramColunaPacienteDispensacao]}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

paramColunaIdDispensacao = 1
paramColunaDataDispensacao = 2
paramColunaCodUnidadeDispensacao = 3
paramColunaNomeUnidadeDispensacao = 4
paramColunaQtdDispensacao = 5
paramColunaPacienteDispensacao = 6
paramColunaDataResultadoAssessor = 13

paramDiasResultado = 3

listaTotalAssessor = pd.read_excel("lista total assessor.xlsx")
filtroPositivo = listaTotalAssessor["Situação"] == "CONFIRMADO"
filtroNegativo = listaTotalAssessor["Situação"] == "NEGATIVO"
listaTotalAssessor = listaTotalAssessor.where(filtroPositivo | filtroNegativo).dropna(how='all')

listaTotalDispensacoes = pd.read_excel("lista total dispensacoes.xlsx", usecols="A:D,J,K")
listaTotalDispensacoes = listaTotalDispensacoes.where(listaTotalDispensacoes["ID DISP."] != "ID DISP.").dropna(how='all')
listaTotalDispensacoes["DATA DISP."] = pd.to_datetime(listaTotalDispensacoes["DATA DISP."], dayfirst=True)

listaDispensacoesIncorreto = []

for row in listaTotalDispensacoes.itertuples():
    agravos = listaTotalAssessor.where(listaTotalAssessor["Cód. Paciente"] == row[paramColunaPacienteDispensacao]).dropna(how='all')
    if agravos.empty:
        continue
    flagRemover = False
    for rowAgravo in agravos.itertuples():
        dataResultado = rowAgravo[paramColunaDataResultadoAssessor]
        dataTeste = row[paramColunaDataDispensacao]
        difDias = abs((dataResultado - dataTeste).days)
        print("Dispensação: " + str(row[paramColunaIdDispensacao]) + "\nAgravo: " + str(rowAgravo[1]) + "\nDiferença de dias (sem abs): " + str((dataResultado - dataTeste).days) + "\nDiferença de dias (com abs): " + str(difDias) + "\nIndex da dispensação: " + str(row.Index))
        if difDias <= paramDiasResultado:
            print("Encontrei diferença menor ou igual a " + str(paramDiasResultado) + ", removendo...")
            flagRemover = True
            break
    if flagRemover:
        appendTabelaAuxiliar(listaDispensacoesIncorreto, row, "Já possui resultado com até " + str(paramDiasResultado) + " dias de diferença")
        listaTotalDispensacoes.drop(row.Index, inplace=True)
        
with pd.ExcelWriter('Dispensacoes x Agravos.xlsx') as writerDisp:
    listaTotalDispensacoes["DATA DISP."] = listaTotalDispensacoes["DATA DISP."].dt.strftime('%d/%m/%Y')
    listaTotalDispensacoes.to_excel(writerDisp, index=False)
    
with pd.ExcelWriter('Incorreto Dispensacoes x Agravos.xlsx') as writerIncorreto:
    listaDispensacoesIncorreto = pd.DataFrame(listaDispensacoesIncorreto)
    listaDispensacoesIncorreto["DATA DISP."] = pd.to_datetime(listaDispensacoesIncorreto["DATA DISP."])
    listaDispensacoesIncorreto["DATA DISP."] = listaTotalDispensacoes["DATA DISP."].dt.strftime('%d/%m/%Y')
    listaDispensacoesIncorreto.to_excel(writerIncorreto, index=False)