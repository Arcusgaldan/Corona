# -*- coding: utf-8 -*-
"""
Created on Mon Aug 30 17:04:21 2021

@author: Info
"""

import pandas as pd
#from datetime import timedelta
from datetime import datetime
import os
from timeit import default_timer as timer

def appendTabelaAuxiliar(templateIncorreto, tabelaIncorreta, row, opcionais=None):
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
    aux = {}
    for key in templateIncorreto:
        aux[key] = row[templateIncorreto[key]]
    #aux = {"Notificacao": row[paramColunaNotifAssessor], "ID Pct": row[paramColunaIdAssessor], "Dt. Notif": row[paramColunaDataNotifAssessor]} #Criação de novo Dict a ser inserido na lista de incorretos
    # if motivo is not None:
    #     aux['Motivo'] = motivo #Adiciona o motivo ao Dict
    # if idDisp is not None:
    #     aux['ID DISP.'] = idDisp
    if opcionais is not None:
        for key in opcionais:
            aux[key] = opcionais[key]
    tabelaIncorreta.append(aux) #Insere o Dict na lista
    
def loading(qtdNotif, countLinhas, porcAtual, somaVeloc, startProcessamento, end=False):
    if end:
        velocMedia = somaVeloc / porcAtual
        print("Finalizado: " + str(countLinhas) + " linhas processadas em " + segundoEmHora(timer() - startProcessamento) + "\nVelocidade média: {:.2f}".format(velocMedia) + " linhas por segundo")
        return
    porcentil = round(qtdNotif/100)
    countLinhas += 1
    if countLinhas >= (porcAtual + 1) * porcentil:
        porcAtual = round((countLinhas/qtdNotif) * 100)
        os.system('cls')
        porcTimer = timer()
        tempoAtual = porcTimer - startProcessamento
        velocEstimada = countLinhas / tempoAtual
        somaVeloc += velocEstimada
        velocMedia = somaVeloc / porcAtual
        tempoEstimadoTotal = (qtdNotif - countLinhas) / velocEstimada
        stringTempoEstimado = segundoEmHora(tempoEstimadoTotal)
        print("Concluido: " + str(porcAtual) + "% aos " + "{:.0f}".format(tempoAtual) + " segundos (desde o início do laço).\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.\nVelocidade estimada: " + "{:.2f}".format(velocEstimada) + " linhas por segundo.\nVelocidade media: " + "{:.2f}".format(velocMedia) + " linhas por segundo.\nTempo estimado: " + stringTempoEstimado + " para conlusão\n\n")
    return countLinhas, porcAtual, somaVeloc

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

paramDataAtual = datetime.today()
paramDataInicioTestes = "01/04/2021"
paramIntervaloDispAgravo = 7

paramColunaNotifAssessor = 1
paramColunaIdAssessor = 2
paramColunaDataNotifAssessor = 11

paramColunaIdDispensacao = 1
paramColunaDataDispensacao = 2

templateIncorreto = {"Notificacao": paramColunaNotifAssessor, "ID Pct": paramColunaIdAssessor, "Dt. Notif": paramColunaDataNotifAssessor}

listaAssessor = pd.read_excel("lista total assessor.xlsx").sort_values(by=["Cód. Paciente", "Data da Notificação"])
listaAssessor = listaAssessor.where((listaAssessor['Situação'] == "CONFIRMADO") & (listaAssessor['Data da Notificação'] >= paramDataInicioTestes)).dropna(how='all')

listaDisp = pd.read_excel("lista total dispensacoes.xlsx")
listaDisp = listaDisp.where(listaDisp['ID DISP.'] != "ID DISP.").dropna(how='all')
listaDisp["DATA DISP."] = pd.to_datetime(listaDisp["DATA DISP."])

listaAssessorIncorreta = []

countLinhas = 0
porcAtual = 0
somaVeloc = 0
qtdNotif = len(listaAssessor)
startProcessamento = timer()

print("Iniciando processamento...")
for row in listaAssessor.itertuples():
    dispPaciente = listaDisp.where(listaDisp['ID. PACIENTE'] == row[paramColunaIdAssessor]).dropna(how='all')
    if not dispPaciente.empty:
        for rowDisp in dispPaciente.itertuples():
            dataAgravo = row[paramColunaDataNotifAssessor]
            dataDisp = rowDisp[paramColunaDataDispensacao]
            if abs((dataAgravo - dataDisp).days) <= paramIntervaloDispAgravo:
                appendTabelaAuxiliar(templateIncorreto, listaAssessorIncorreta, row, {'Motivo': 'Dispensação com menos de ' + str(paramIntervaloDispAgravo) + ' dias de diferença', 'ID. DISP': rowDisp[paramColunaIdDispensacao], 'DATA DISP.': rowDisp[paramColunaDataDispensacao]})
                listaAssessor.drop(row.Index, inplace=True)
                break
    countLinhas, porcAtual, somaVeloc = loading(qtdNotif, countLinhas, porcAtual, somaVeloc, startProcessamento)
    
loading(qtdNotif, countLinhas, porcAtual, somaVeloc, startProcessamento, end=True)

with pd.ExcelWriter('Positivos sem teste ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:    
    listaAssessor.to_excel(writer, "Sem teste", index=False)
    listaAssessorIncorreta = pd.DataFrame(listaAssessorIncorreta)
    listaAssessorIncorreta.to_excel(writer, "Incorretos", index=False)