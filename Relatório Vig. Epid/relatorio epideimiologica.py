# -*- coding: utf-8 -*-
"""
Created on Sat Jul  3 19:26:40 2021

@author: Thales
"""

import pandas as pd
from datetime import timedelta
from datetime import datetime

listaUnidades = [
    {'CNES':2035731, 'Nome':'UNIDADE BASICA DE SAUDE DO IBITU'},
    {'CNES':2048736, 'Nome':'UNIDADE BASICA DE SAUDE DR APOLONIO MORAES E SOUZA'},
    {'CNES':2048744, 'Nome':'UNIDADE BASICA DR JOSE PARASSU CARVALHO'},
    {'CNES':2053314, 'Nome':'UNIDADE BASICA DE SAUDE DR LOTFALLAH MIZIARA'},
    {'CNES':2056062, 'Nome':'UNIDADE BASICA ARCHIMEDES MACHADO DE BARRETOS'},
    {'CNES':2061473, 'Nome':'UNIDADE BASICA DE SAUDE DR PAULO PRATA'},
    {'CNES':2064081, 'Nome':'UBS DR MILTON BARONI DE BARRETOS'},
    {'CNES':2064103, 'Nome':'UBS DR ALLY ALAHMAR DE BARRETOS'},
    {'CNES':2093642, 'Nome':'UNIDADE BASICA DE SAUDE DR WILSON HAYEK SAIHG'},
    {'CNES':2093650, 'Nome':'UNIDADE BASICA DE SAUDE DR BARTOLOMEU MARAGLIANO V'},
    {'CNES':2784580, 'Nome':'UNIDADE BASICA DE SAUDE DR SERGIO PIMENTA'},
    {'CNES':2784599, 'Nome':'UNIDADE BASICA DE SAUDE FRANCOLINO GALVAO DE SOUZA'},
    {'CNES':7035861, 'Nome':'UPA 24H ZAID ABRAO GERAIGE'},
    {'CNES':7122217, 'Nome':'UNIDADE ESF DR LUIZ SPINA'},
    {'CNES':7565577, 'Nome':'AMBULATORIO SAUDE DO IDOSO'},
    {'CNES':7585020, 'Nome':'USF DO BAIRRO NOVA BARRETOS'},
    {'CNES':1456678, 'Nome':'TESTE DO ZERO'}
]

paramDataAtual = datetime.now()
paramTamanhoSemana = 7

def qtdAgravos(situacao, unidade, semana): #Retorna a quantidade de suspeitas da unidade, total ou da semana
    filtro = (listaTotalAssessor["Situação"] == situacao) & (listaTotalAssessor["Unidade"] == unidade["Nome"])
    if semana:
        filtro = filtro & (listaTotalAssessor["Data da Digitação"] <= paramDataAtual) & (listaTotalAssessor["Data da Digitação"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    agravosUnidade = listaTotalAssessor.where(filtro).dropna(how='all')
    return len(agravosUnidade)

def semInicioSintomas(unidade, semana): #Retorna a quantidade de notificações sem data de inicio de sintomas da unidade, total ou da semana
    filtro = ((listaTotalAssessor["Situação"] == "SUSPEITA") | (listaTotalAssessor["Situação"] == "CONFIRMADO")) & (listaTotalAssessor["Unidade"] == unidade["Nome"]) & (pd.isnull(listaTotalAssessor["Data dos Primeiros Sintomas"]))
    if semana:
        filtro = filtro & (listaTotalAssessor["Data da Digitação"] <= paramDataAtual) & (listaTotalAssessor["Data da Digitação"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    semSintomasUnidade = listaTotalAssessor.where(filtro).dropna(how='all')
    return len(semSintomasUnidade)

def difDigitNotif(unidade, semana): #Retorna a quantidade de notificações com diferença entre Data de Digitação e Data de Notificação, total ou da semana
    filtro = (listaTotalAssessor["Situação"] != "DESCARTADO") & (listaTotalAssessor["Unidade"] == unidade["Nome"]) & (listaTotalAssessor["Data da Digitação"] != listaTotalAssessor["Data da Notificação"])
    if semana:
        filtro = filtro & (listaTotalAssessor["Data da Digitação"] <= paramDataAtual) & (listaTotalAssessor["Data da Digitação"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    diferenca = listaTotalAssessor.where(filtro).dropna(how='all')
    return len(diferenca)

listaTotalAssessor = pd.read_excel("lista total assessor.xlsx")

for unidade in listaUnidades:
    unidade["Suspeitas Total"] = qtdAgravos("SUSPEITA", unidade, False)   
    unidade["Suspeitas Semana"] = qtdAgravos("SUSPEITA", unidade, True)
    unidade["Positivos Total"] = qtdAgravos("CONFIRMADO", unidade, False)   
    unidade["Positivos Semana"] = qtdAgravos("CONFIRMADO", unidade, True)
    unidade["Negativos Total"] = qtdAgravos("NEGATIVO", unidade, False)   
    unidade["Negativos Semana"] = qtdAgravos("NEGATIVO", unidade, True)
    unidade["Sem Sintomas Total"] = semInicioSintomas(unidade, False)
    unidade["Sem Sintomas Semana"] = semInicioSintomas(unidade, True)
    unidade["Diferença Digitação x Notificação Total"] = difDigitNotif(unidade, False)
    unidade["Diferença Digitação x Notificação Semana"] = difDigitNotif(unidade, True)