# -*- coding: utf-8 -*-
"""
Created on Tue Jul 13 08:32:38 2021

@author: Info
"""

import pandas as pd
import re
import sys
from datetime import datetime

def appendTabelaAuxiliar(tabela, row, situacao, motivo):
    aux = {'ID PACIENTE': row[paramColunaIdPacienteExame], 'DT RESULTADO': row[paramColunaDataExame], 'SITUACAO': situacao}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

paramColunaIdAgravoAssessor = 3
paramColunaDataNotifAssessor = 11
paramColunaSituacaoAgravoAssessor = 14

paramColunaIdPacienteExame = 4
paramColunaDataExame = 6

paramDifDias = 3
paramDataAtual = datetime.today()

listaExames = pd.read_excel("examesUPA.xlsx")
listaExames["DT RESULTADO"] = pd.to_datetime(listaExames["DT RESULTADO"], dayfirst=True)
listaExames.insert(8, "SITUAÇÃO", None)
listaExames.insert(9, "TEM SUSPEITA", None)

listaExamesIncorreto = []

listaAssessor = pd.read_excel("lista total assessor.xlsx")

regexPositivo = re.compile('.*POS.*', re.IGNORECASE)
regexNegativo = re.compile('.*NEG.*', re.IGNORECASE)

# contPositivo = 0
# contNegativo = 0

for row in listaExames.itertuples():
    situacaoRow = ""
    if regexPositivo.match(row.RESULTADO):
        # contPositivo += 1
        # print("Achei um positivo! ContPositivo = " + str(contPositivo))
        situacaoRow = "CONFIRMADO"
    elif regexNegativo.match(row.RESULTADO):
        # contNegativo += 1
        # print("Achei um negativo! ContNegativo = " + str(contNegativo))
        situacaoRow = "NEGATIVO"
    else:
        print("Encontrei um resultado que não é positivo nem negativo, linha " + row.Index)
        sys.exit()
    
    listaExames.at[row.Index, "SITUAÇÃO"] = situacaoRow
    filtroSituacao = (listaAssessor["Situação"] == situacaoRow) | (listaAssessor["Situação"] == "SUSPEITA")
    agravosAssessor = listaAssessor.where((listaAssessor["Cód. Paciente"] == row[paramColunaIdPacienteExame]) & filtroSituacao).dropna(how='all')
    if agravosAssessor.empty:
        #print("Ih vazio")
        listaExames.at[row.Index, "TEM SUSPEITA"] = "NÃO"
        continue
    flagSuspeita = False
    flagRemover = False
    #print("Achei um não vazio")
    #sys.exit()
    for rowAgravo in agravosAssessor.itertuples():
        if rowAgravo[paramColunaSituacaoAgravoAssessor] == "SUSPEITA":
            dataNotifAssessor = rowAgravo[paramColunaDataNotifAssessor]
            dataExame = row[paramColunaDataExame]
            difDias = abs((dataNotifAssessor - dataExame).days)
            if difDias <= paramDifDias:
                flagSuspeita = True
            continue
        else:
            dataNotifAssessor = rowAgravo[paramColunaDataNotifAssessor]
            dataExame = row[paramColunaDataExame]
            difDias = abs((dataNotifAssessor - dataExame).days)
            if difDias <= paramDifDias:
                flagRemover = True
                break
    if flagRemover:
        appendTabelaAuxiliar(listaExamesIncorreto, row, situacaoRow, "Já tem agravo " + situacaoRow + " dentro de " + str(paramDifDias) + " dias")
        listaExames.drop(row.Index, inplace=True)
    elif flagSuspeita:
        listaExames.at[row.Index, "TEM SUSPEITA"] = "SIM"
    else:
        listaExames.at[row.Index, "TEM SUSPEITA"] = "NÃO"
        
with pd.ExcelWriter('Exames UPA ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:
    listaExames["DT RESULTADO"] = pd.to_datetime(listaExames["DT RESULTADO"])
    listaExames = listaExames[["ID. PAC.", "DT RESULTADO", "SITUAÇÃO", "TEM SUSPEITA"]]
    examesPositivo = listaExames.where(listaExames["SITUAÇÃO"] == "CONFIRMADO").dropna(how='all')
    examesNegativo = listaExames.where(listaExames["SITUAÇÃO"] == "NEGATIVO").dropna(how='all')
    examesPositivo.to_excel(writer, "Positivos", index=False)
    examesNegativo.to_excel(writer, "Negativos", index=False)

with pd.ExcelWriter('Incorretos Exames UPA ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerIncorreto:
    listaExamesIncorreto = pd.DataFrame(listaExamesIncorreto)
    if not listaExamesIncorreto.empty:
        listaExamesIncorreto['DT RESULTADO'] = pd.to_datetime(listaExamesIncorreto['DT RESULTADO'])
        listaExamesIncorreto.to_excel(writerIncorreto, index=False)