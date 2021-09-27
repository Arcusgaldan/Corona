# -*- coding: utf-8 -*-
"""
Created on Fri Aug 27 11:44:42 2021

@author: Info
"""

import pandas as pd
#from datetime import timedelta
from datetime import datetime
import os
from timeit import default_timer as timer

def appendTabelaAuxiliar(tabelaIncorreta, row, motivo):
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
    aux = {"Notificacao": row[paramColunaNotifAssessor], "ID Pct": row[paramColunaIdAssessor], "Dt. Notif": row[paramColunaDataNotifAssessor]} #Criação de novo Dict a ser inserido na lista de incorretos
    if motivo is not None:
        aux['Motivo'] = motivo #Adiciona o motivo ao Dict
    tabelaIncorreta.append(aux) #Insere o Dict na lista
    
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

def loading(qtdNotif, countLinhas, porcAtual, somaVeloc, startProcessamento, end=False):
    if end:
        velocMedia = somaVeloc / porcAtual
        print("Finalizado: " + str(countLinhas) + " de " + str(qtdNotif) + " linhas processadas em " + segundoEmHora(timer() - startProcessamento) + "\nVelocidade Média: " + str(velocMedia) + " linhas por segundo")
        return
    porcentil = round(qtdNotif/100)
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
        print("Concluido: " + str(porcAtual) + "% aos " + "{:.0f}".format(tempoAtual) + " segundos (desde o início do laço).\n" + str(countLinhas) + " linhas lidas de " + str(qtdNotif) + " total.\nVelocidade estimada: " + "{:.2f}".format(velocEstimada) + " linhas por segundo.\nVelocidade media: " + "{:.2f}".format(velocMedia) + " linhas por segundo.\nTempo estimado: " + stringTempoEstimado + " para conlusão\n\n")
    return countLinhas, porcAtual, somaVeloc

paramColunaNotifAssessor = 1
paramColunaIdAssessor = 2
paramColunaDataNotifAssessor = 11

paramQtdDias = 14
paramDataAtual = datetime.now()

while True:
    while True:
        try:
            print("Insira a data inicial (dd/mm/aaaa): ", end="")
            dataInicial = input()
            dataInicial = datetime.strptime(dataInicial, "%d/%m/%Y")
        except Exception:
            if Exception == KeyboardInterrupt:
                raise KeyboardInterrupt
            print("Falha ao gravar a data inicial. Por favor tente novamente")
            continue
        break
    
    while True:
        try:
            print("Insira a data final (dd/mm/aaaa): ", end="")
            dataFinal = input()
            dataFinal = datetime.strptime(dataFinal, "%d/%m/%Y")
        except Exception:
            if Exception == KeyboardInterrupt:
                raise KeyboardInterrupt
            print("Falha ao gravar a data inicial. Por favor tente novamente")
            continue
        break
    
    if dataInicial > dataFinal:
        print("Data inicial não pode ser maior que data final")
        continue
    break

print("Aguarde...")

tabelaAssessor = pd.read_excel("lista total assessor.xlsx").sort_values(by=["Cód. Paciente", "Data da Notificação"], ignore_index=True) #Lista total dos agravos de março até hoje. Organizado por ID do paciente, index de 0 a n-1
tabelaAssessor = tabelaAssessor.where(tabelaAssessor["Município de residência"] == "BARRETOS - SP").dropna(how='all')
tabelaAssessorDias = tabelaAssessor.where((tabelaAssessor["Data da Notificação"] >= dataInicial) & (tabelaAssessor["Data da Notificação"] <= dataFinal)).dropna(how='all')

tabelaNotificadosIncorretos = []

countLinhas = 0
porcAtual = 0
somaVeloc = 0
qtdNotif = len(tabelaAssessorDias)
startProcessamento = timer()
lastPct = None

print("Iniciando processamento...")
for row in tabelaAssessorDias.itertuples():
    if row[paramColunaIdAssessor] == lastPct:
        continue
    print("Entrei no laço, notificação " + str(row[paramColunaNotifAssessor]) + "\nContagem: " + str(countLinhas))
    notificacoesPct = tabelaAssessor.where(tabelaAssessor["Cód. Paciente"] == row[paramColunaIdAssessor]).dropna(how='all')
    if not (notificacoesPct.empty or len(notificacoesPct) == 1):
        pivot = None
        for rowNotif in notificacoesPct.itertuples():
            if pivot == None:
                pivot = rowNotif
                #print("Não havia pivot, agora pivot = " + str(pivot[paramColunaNotifAssessor]))
            if rowNotif[paramColunaNotifAssessor] == pivot[paramColunaNotifAssessor]:
                continue
            if (rowNotif[paramColunaDataNotifAssessor] - pivot[paramColunaDataNotifAssessor]).days <= 14:
                appendTabelaAuxiliar(tabelaNotificadosIncorretos, row, "Possui notificação com 14 dias de diferença")
                tabelaAssessor.drop(rowNotif.Index, inplace=True)
                if rowNotif[paramColunaDataNotifAssessor] >= dataInicial and rowNotif[paramColunaDataNotifAssessor] <= dataFinal:
                    tabelaAssessorDias.drop(rowNotif.Index, inplace=True)
                notificacoesPct.drop(rowNotif.Index, inplace=True)
                #print("Removi a notificação " + str(rowNotif[paramColunaNotifAssessor]) + " pois tinha menos de 14 dias de diferença do meu pivot.\nData Pivot: " + str(pivot[paramColunaDataNotifAssessor]) + "\nData do Excluido: " + str(rowNotif[paramColunaDataNotifAssessor]))
            else:                
                #print("Mudando pivot para " + str(rowNotif[paramColunaNotifAssessor]) + " pois a diferença é maior que 14 dias.\nData antigo pivot: " + str(pivot[paramColunaDataNotifAssessor]) + "\nData novo pivot: " + str(rowNotif[paramColunaDataNotifAssessor]))
                pivot = rowNotif
    
    lastPct = row[paramColunaIdAssessor]
    countLinhas, porcAtual, somaVeloc = loading(qtdNotif, countLinhas, porcAtual, somaVeloc, startProcessamento)
        
with pd.ExcelWriter('Notificados ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:    
    tabelaAssessorDias.to_excel(writer, "Notificados", index=False)
    tabelaNotificadosIncorretos = pd.DataFrame(tabelaNotificadosIncorretos)
    tabelaNotificadosIncorretos.to_excel(writer, "Incorretos", index=False)
            