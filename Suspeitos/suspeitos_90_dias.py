# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 13:12:28 2020

@author: Info
"""


import pandas as pd
from datetime import timedelta
from datetime import datetime
import numpy as np

tabelaTotal = pd.read_excel("lista total.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by="Cód. Paciente", ignore_index=True) #Lista total dos agravos de março até hoje. Organizado por ID do paciente, index de 0 a n-1
#tabelaSuspeitos = pd.read_excel("suspeitos junho.xls").sort_values(by=["Cód. Paciente", "Data da Notificação"]).drop_duplicates("Cód. Paciente", ignore_index=True) #Lista de agravos suspeitos de junho até hoje. Organizado por ID do paciente e Data de Notificação. Removendo as duplicatas por ID, mantendo a da data mais antiga. Index do 0 a n-1
teste = datetime.strptime('01-06-2020', '%d-%m-%Y')
tabelaSuspeitos = tabelaTotal.where(tabelaTotal["Data da Notificação"] >= datetime.strptime('01-06-2020', '%d-%m-%Y'))
tabelaSuspeitos = tabelaSuspeitos.where(tabelaSuspeitos["Situação"] == "SUSPEITA").dropna(how='all').sort_values(by=["Cód. Paciente", "Data da Notificação"]).drop_duplicates("Cód. Paciente", ignore_index=True)
tabelaSuspeitosIncorretos = [] #Lista de dicts que vai segurar meus suspeitos incorretos
for i in tabelaSuspeitos.index: #Passando por cada suspeito
    situacoes = tabelaTotal.where(tabelaTotal["Cód. Paciente"] == tabelaSuspeitos['Cód. Paciente'][i])[["Situação", "Data da Notificação"]].dropna() #Pega apenas a situação e data de notificação de cada registro que o suspeito possui na lista total
    flag = False
    #print("Teste de situação: " + situacoes["Situação"].values)
    if "CONFIRMADO" in situacoes["Situação"].values: #Se houver um confirmado na lista de situações
        for j in situacoes.index: #Passa por cada situação
            flag = False
            if situacoes["Situação"][j] != "CONFIRMADO": #Se a situação não for Confirmada, ignora
                continue
            if situacoes["Data da Notificação"][j] >= (tabelaSuspeitos["Data da Notificação"][i] - timedelta(days=90)):
                aux = tabelaSuspeitos.loc[i] #Pega o suspeito
                aux["Motivo"] = "Confirmado" #Adiciona o motivo do "erro" como "Confirmado"
                tabelaSuspeitosIncorretos.append(aux) #Coloca o suspeito na lista de Suspeitos Incorretos (pois é confirmado)
                tabelaSuspeitos.drop(i, inplace=True) #Tira o suspeito da lista de suspeitos
                flag = True
                break
        if flag:
            continue
    if "NEGATIVO" in situacoes["Situação"].values: #Se houver um negativo nas situações
        for j in situacoes.index: #Passa por cada situação
            if situacoes["Situação"][j] != "NEGATIVO": #Se a situação não for Negativo, ignora
                continue
            if situacoes["Data da Notificação"][j] >= (tabelaSuspeitos["Data da Notificação"][i] - timedelta(days=2)):
                #print("Achei um negativo dentro do período\nData da suspeita (suspeito "+ tabelaSuspeitos["Paciente"][i] +", Index: "+ str(i) +"): "+ str(tabelaSuspeitos["Data da Notificação"][i]) + "\nData do Negativo: "+ str(situacoes["Data da Notificação"][j]))
                aux = tabelaSuspeitos.loc[i] #Pega o suspeito
                aux["Motivo"] = "Negativo" #Adiciona o motivo do "erro" como "Negativo"
                tabelaSuspeitosIncorretos.append(aux) #Coloca o suspeito na lista de Suspeitos Incorretos (pois é negativo)
                tabelaSuspeitos.drop(i, inplace=True) #Tira o suspeito da lista de suspeitos
                #print("Removi o suspeito que é negativo")
                break


tabelaSuspeitosIncorretos = pd.DataFrame(tabelaSuspeitosIncorretos)
tabelaSuspeitos.to_excel('lista_realmente_suspeitos.xlsx')
tabelaSuspeitosIncorretos.to_excel('suspeitos_incorretos.xlsx')
tabelaTotal = tabelaTotal.where(tabelaTotal["Situação"] != "SUSPEITA").dropna(how='all')
tabelaTotal = pd.concat([tabelaTotal, tabelaSuspeitos])
tabelaTotal.to_excel('total_e_realmente_suspeitos.xlsx', index=False)
        

