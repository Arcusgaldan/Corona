# -*- coding: utf-8 -*-
"""
Created on Wed Jun 23 14:58:29 2021

@author: Info
"""

mapaCnes = {}
mapaCnes['2035731'] = 'UNIDADE BASICA DE SAUDE DO IBITU'
mapaCnes['2048736'] = 'UNIDADE BASICA DE SAUDE DR APOLONIO MORAES E SOUZA'
mapaCnes['2048744'] = 'UNIDADE BASICA DR JOSE PARASSU CARVALHO'
mapaCnes['2053314'] = 'UNIDADE BASICA DE SAUDE DR LOTFALLAH MIZIARA'
mapaCnes['2056062'] = 'UNIDADE BASICA ARCHIMEDES MACHADO DE BARRETOS'
mapaCnes['2061473'] = 'UNIDADE BASICA DE SAUDE DR PAULO PRATA'
mapaCnes['2064081'] = 'UBS DR MILTON BARONI DE BARRETOS'
mapaCnes['2064103'] = 'UBS DR ALLY ALAHMAR DE BARRETOS'
mapaCnes['2093642'] = 'UNIDADE BASICA DE SAUDE DR WILSON HAYEK SAIHG'
mapaCnes['2093650'] = 'UNIDADE BASICA DE SAUDE DR BARTOLOMEU MARAGLIANO V'
mapaCnes['2784580'] = 'UNIDADE BASICA DE SAUDE DR SERGIO PIMENTA'
mapaCnes['2784599'] = 'UNIDADE BASICA DE SAUDE FRANCOLINO GALVAO DE SOUZA'
mapaCnes['7035861'] = 'UPA 24H ZAID ABRAO GERAIGE'
mapaCnes['7122217'] = 'UNIDADE ESF DR LUIZ SPINA'
mapaCnes['7565577'] = 'AMBULATORIO SAUDE DO IDOSO'
mapaCnes['7585020'] = 'USF DO BAIRRO NOVA BARRETOS'


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
paramColunaDataNotifAssessor = 11
paramColunaDataResultadoAssessor = 13
paramColunaSituacaoAgravoAssessor = 14
paramColunaCodAgravoAssessor = 1

paramDiasResultado = 3
paramDiasAtendimento = 0

listaTotalAssessor = pd.read_excel("lista total assessor.xlsx")
# filtroPositivo = listaTotalAssessor["Situação"] == "CONFIRMADO"
# filtroNegativo = listaTotalAssessor["Situação"] == "NEGATIVO"
# listaTotalAssessor = listaTotalAssessor.where(filtroPositivo | filtroNegativo).dropna(how='all')

listaTotalDispensacoes = pd.read_excel("lista total dispensacoes.xlsx", usecols="A:D,J,K")
listaTotalDispensacoes = listaTotalDispensacoes.where(listaTotalDispensacoes["ID DISP."] != "ID DISP.").dropna(how='all')
listaTotalDispensacoes["DATA DISP."] = pd.to_datetime(listaTotalDispensacoes["DATA DISP."], dayfirst=True)
listaTotalDispensacoes.insert(len(listaTotalDispensacoes.columns), "Situação", None)
listaTotalDispensacoes.insert(len(listaTotalDispensacoes.columns), "Foi atendido", None)

listaTotalAgravos = pd.read_excel("lista total agravos.xlsx")

listaDispensacoesIncorreto = []

for row in listaTotalDispensacoes.itertuples():
    agravosTotal = listaTotalAgravos.where((listaTotalAgravos["Cód. Paciente"] == row[paramColunaPacienteDispensacao]) & (listaTotalAgravos["Unidade"] == mapaCnes[str(row[paramColunaCodUnidadeDispensacao])])).dropna(how='all')
    flagAtendido = False
    if(not agravosTotal.empty):               
        for rowTotal in agravosTotal.itertuples():
            print("Achei o agravo " + str(rowTotal[paramColunaCodAgravoAssessor]))
            dataNotif = rowTotal[paramColunaDataNotifAssessor]
            dataTeste = row[paramColunaDataDispensacao]
            difDias = abs((dataNotif - dataTeste).days)
            if difDias <= paramDiasAtendimento:
                print("E é dentro do período!! Foi atendido = SIM")
                listaTotalDispensacoes.at[row.Index, "Foi atendido"] = "SIM"
                flagAtendido = True
                break
            else:
                print("Mas não é dentro do período, deu " + str(difDias) + " dias de diferença, param é " + str(paramDiasAtendimento))
    if not flagAtendido:
        listaTotalDispensacoes.at[row.Index, "Foi atendido"] = "NÃO"        
    
    agravos = listaTotalAssessor.where(listaTotalAssessor["Cód. Paciente"] == row[paramColunaPacienteDispensacao]).dropna(how='all')
    if agravos.empty:
        listaTotalDispensacoes.at[row.Index, "Situação"] = "PACIENTE SEM SUSPEITA"
        continue
    flagRemover = False
    flagSuspeita = False
    for rowAgravo in agravos.itertuples():
        if rowAgravo[paramColunaSituacaoAgravoAssessor] == "SUSPEITA":
            #print("Achei uma suspeita")
            dataNotif = rowAgravo[paramColunaDataNotifAssessor]
            dataTeste = row[paramColunaDataDispensacao]
            difDias = abs((dataNotif - dataTeste).days)
            if difDias <= paramDiasResultado:
                #print("E está dentro dos dias do parâmetro! Tem suspeita!")
                flagSuspeita = True
                continue
            else:
                #print("Porém não estava dentro dos dias do parâmetro")
                continue
        elif rowAgravo[paramColunaSituacaoAgravoAssessor] == "CONFIRMADO" or rowAgravo[paramColunaSituacaoAgravoAssessor] == "NEGATIVO":
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
    else:
        if flagSuspeita:
            listaTotalDispensacoes.at[row.Index, "Situação"] = "PACIENTE COM SUSPEITA EM ABERTO"
        else:
            listaTotalDispensacoes.at[row.Index, "Situação"] = "PACIENTE SEM SUSPEITA"
        
with pd.ExcelWriter('Dispensacoes x Agravos.xlsx', engine='openpyxl') as writerDisp:
    listaTotalDispensacoes["DATA DISP."] = listaTotalDispensacoes["DATA DISP."].dt.strftime('%d/%m/%Y')
    listaTotalDispensacoes.to_excel(writerDisp, index=False)
    
with pd.ExcelWriter('Incorreto Dispensacoes x Agravos.xlsx', engine='openpyxl') as writerIncorreto:
    listaDispensacoesIncorreto = pd.DataFrame(listaDispensacoesIncorreto)
    listaDispensacoesIncorreto["DATA DISP."] = pd.to_datetime(listaDispensacoesIncorreto["DATA DISP."], dayfirst=True)
    listaDispensacoesIncorreto["DATA DISP."] = listaDispensacoesIncorreto["DATA DISP."].dt.strftime('%d/%m/%Y')
    listaDispensacoesIncorreto.to_excel(writerIncorreto, index=False)