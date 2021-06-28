# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 13:12:28 2020

@author: Info
"""


import pandas as pd
from datetime import timedelta
from datetime import datetime

tabelaTotal = pd.read_excel("lista total.xls").sort_values(by=["Cód. Paciente", "Data da Notificação"], ignore_index=True) #Lista total dos agravos de março até hoje. Organizado por ID do paciente, index de 0 a n-1
tabelaPositivos = tabelaTotal.where(tabelaTotal["Situação"] == "CONFIRMADO").dropna(how='all').sort_values(by=["Cód. Paciente", "Data da Notificação"])
tabelaPositivosIncorretos = [] #Lista de dicts que vai segurar meus suspeitos incorretos
ultimoId = None
for i in tabelaPositivos.index: #Passando por cada positivo
    if i not in tabelaPositivos.index:
        continue
    if tabelaPositivos['Cód. Paciente'][i] == ultimoId:
        continue
    ultimoId = tabelaPositivos['Cód. Paciente'][i]
    notificacoesPaciente = tabelaPositivos.where(tabelaPositivos["Cód. Paciente"] == tabelaPositivos['Cód. Paciente'][i])[["Cód. Paciente", "Paciente", "Situação", "Data da Notificação"]].dropna(how='all') #Pega apenas a situação e data de notificação de cada registro que o suspeito possui na lista total
    if(len(notificacoesPaciente.index) == 1): #Se só houver uma notificação do paciente, continuar a iteração
        continue
    referencia = notificacoesPaciente.iloc[[0]]    
    indexVerificar = 1
    verificar = notificacoesPaciente.iloc[[indexVerificar]]
    for j in notificacoesPaciente.index:
        #print("Diferença: " + verificar["Data da Notificação"] - timedelta(days=153))
        if verificar.iloc[0]["Data da Notificação"] - timedelta(days=153) < referencia.iloc[0]["Data da Notificação"]: #A notificação verificada tem menos de 5 meses
            print("Encontrei uma notificação com menos de 5 meses, index = " + str(verificar.index[0]))
            notificacoesPaciente.drop(index=int(verificar.index[0]), inplace=True)
            tabelaPositivos.drop(index=int(verificar.index[0]), inplace=True)
            tabelaPositivosIncorretos.append({"ID": verificar.iloc[0]["Cód. Paciente"], "Data da Notificação": verificar.iloc[0]["Data da Notificação"]})
            if(len(notificacoesPaciente.index) == indexVerificar): #Se não houver mais registros deste paciente, sair do loop
                break
            else:
                verificar = notificacoesPaciente.iloc[[indexVerificar]]
        else: #Notificação tem mais de 5 meses
            print("Encontrei uma notificação com mais de 5 meses, index = " + str(verificar.index[0]))
            #print("Achei uma notificação com mais de 5 meses, paciente ID " + verificar["Cód. Paciente"])
            referencia = verificar
            indexVerificar += 1
            if(len(notificacoesPaciente.index) == indexVerificar): #Se não houver mais registros deste paciente, sair do loop
                break
            else:
                verificar = notificacoesPaciente.iloc[[indexVerificar]]

tabelaPositivosIncorretos = pd.DataFrame(tabelaPositivosIncorretos)
tabelaPositivos.to_excel('lista_positivos.xls')
tabelaPositivosIncorretos.to_excel('lista_positivos_incorretos.xls')
        

