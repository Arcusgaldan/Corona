# -*- coding: utf-8 -*-
"""
Created on Sat Jul  3 19:26:40 2021

@author: Thales
"""

import pandas as pd
from datetime import timedelta
from datetime import datetime
from datetime import date

from openpyxl import Workbook

listaUnidades = [
    {'CNES':2035731, 'Nome':'UNIDADE BASICA DE SAUDE DO IBITU', 'Apelido': 'Ibitu'},
    {'CNES':2048736, 'Nome':'UNIDADE BASICA DE SAUDE DR APOLONIO MORAES E SOUZA', 'Apelido': 'Los Angeles'},
    {'CNES':2048744, 'Nome':'UNIDADE BASICA DR JOSE PARASSU CARVALHO', 'Apelido': 'Ibirapuera'},
    {'CNES':2053314, 'Nome':'UNIDADE BASICA DE SAUDE DR LOTFALLAH MIZIARA', 'Apelido': 'Barretos 2'},
    {'CNES':2056062, 'Nome':'UNIDADE BASICA ARCHIMEDES MACHADO DE BARRETOS', 'Apelido': 'Pimenta'},
    {'CNES':2061473, 'Nome':'UNIDADE BASICA DE SAUDE DR PAULO PRATA', 'Apelido': 'Cristiano de Carvalho'},
    {'CNES':2064103, 'Nome':'UBS DR ALLY ALAHMAR DE BARRETOS', 'Apelido': 'CSU'},
    {'CNES':2093642, 'Nome':'UNIDADE BASICA DE SAUDE DR WILSON HAYEK SAIHG', 'Apelido': 'Cecapinha'},
    {'CNES':2093650, 'Nome':'UNIDADE BASICA DE SAUDE DR BARTOLOMEU MARAGLIANO V', 'Apelido': 'Derby'},
    {'CNES':2784580, 'Nome':'UNIDADE BASICA DE SAUDE DR SERGIO PIMENTA', 'Apelido': 'Marilia'},
    {'CNES':2784599, 'Nome':'UNIDADE BASICA DE SAUDE FRANCOLINO GALVAO DE SOUZA', 'Apelido': 'São Francisco'},
    {'CNES':7035861, 'Nome':'UPA 24H ZAID ABRAO GERAIGE', 'Apelido': 'UPA'},
    {'CNES':7122217, 'Nome':'UNIDADE ESF DR LUIZ SPINA', 'Apelido': 'Spina'},
    {'CNES':7565577, 'Nome':'AMBULATORIO SAUDE DO IDOSO', 'Apelido': 'Saúde do Idoso'},
    {'CNES':7585020, 'Nome':'USF DO BAIRRO NOVA BARRETOS', 'Apelido': 'Nova Barretos'},
    {'CNES':2784572, 'Nome':'AMBULATORIO DE ESPECIALIDADES ARE I', 'Apelido': 'Postão'},
    {'CNES':2064081, 'Nome':'UBS DR MILTON BARONI DE BARRETOS', 'Apelido': 'América'}
]

codigosVacina = [17969, 17970, 17971, 17972]


paramDataAtual = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
paramDataTruncAtual = date.today()
paramTamanhoSemana = 7
paramDataAtras = paramDataAtual - timedelta(days=7)

filtroTeste1 = None
filtroTeste2 = None
filtroCodsTeste = None

paramCodigoSaidaVacina = 15

def qtdAgravos(situacao, unidade, semana): #Retorna a quantidade de suspeitas da unidade, total ou da semana
    filtro = (listaTotalAssessor["Situação"] == situacao) & (listaTotalAssessor["Unidade"] == unidade["Nome"])
    if semana:
        filtro = filtro & (listaTotalAssessor["Data da Digitação"] <= paramDataAtual) & (listaTotalAssessor["Data da Digitação"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    agravosUnidade = listaTotalAssessor.where(filtro).dropna(how='all')
    return len(agravosUnidade)

def semInicioSintomas(unidade, semana): #Retorna a quantidade de notificações (suspeitas e confirmadas) sem data de inicio de sintomas da unidade, total ou da semana
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

def dispSemAgravo(unidade, semana):
    filtro = listaDispSemAgravo["COD. UNID."] == unidade["CNES"]
    if semana:
        filtro = filtro & (listaDispSemAgravo["DATA DISP."] <= paramDataAtual) & (listaDispSemAgravo["DATA DISP."] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    semAgravo = listaDispSemAgravo.where(filtro).dropna(how='all')
    # if unidade["CNES"] == 7035861:
    #     with pd.ExcelWriter('debug.xlsx') as writer:
    #         semAgravo.to_excel(writer, "UPA sem agravo da semana")
    #         listaDispSemAgravo.to_excel(writer, "Total")
    return len(semAgravo)

def dispensacoes(unidade, semana):
    filtro = listaTotalDispensacoes["COD. UNID."] == unidade["CNES"]
    if semana:
        filtro = filtro & (listaTotalDispensacoes["DATA DISP."] <= paramDataAtual) & (listaTotalDispensacoes["DATA DISP."] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    totalDispensacoes = listaTotalDispensacoes.where(filtro).dropna(how='all')
    return totalDispensacoes["QTD DISP."].sum()

def vacivida(unidade, semana):
    filtro = listaVacivida["CNES_ESTABELECIMENTO"] == str(unidade["CNES"])
    if semana:
        filtro = filtro & (listaVacivida["DATA_APLICACAO_VACINA"] <= paramDataAtual) & (listaVacivida["DATA_APLICACAO_VACINA"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))
    retornoVacivida = listaVacivida.where(filtro).dropna(how='all')
    return len(retornoVacivida)

def saidaVacina(unidade, tipoSaida, semana):
    global filtroTeste1
    global filtroTeste2
    global filtroCodsTeste
    filtro = (listaTotalSaidas["CÓD UNID."] == unidade["CNES"]) & (listaTotalSaidas["TIPO MOVIMENTAÇÃO"] == "SAÍDAS DIVERSAS") & (listaTotalSaidas["TIPO"] == tipoSaida)
    filtroTeste1 = filtro
    filtroCods = None
    for cod in codigosVacina:
        if filtroCods is None:
            filtroCods = (listaTotalSaidas["CÓD MAT."] == cod)
        else:
            filtroCods = filtroCods | (listaTotalSaidas["CÓD MAT."] == cod)
    filtroCodsTeste = filtroCods
    filtro = filtro & filtroCods
    filtroTeste2 = filtro
    if semana:
        filtro = filtro & (listaTotalSaidas["DATA MOV"] <= paramDataAtual) & (listaTotalSaidas["DATA MOV"] >= paramDataAtual - timedelta(days=paramTamanhoSemana))    
    saidas = listaTotalSaidas.where(filtro).dropna(how='all')
    return saidas["QTD"].sum()
    
listaTotalAssessor = pd.read_excel("lista total assessor.xlsx")

listaDispSemAgravo = pd.read_excel("disp sem agravo.xlsx").drop_duplicates("ID DISP.")
listaDispSemAgravo["DATA DISP."] = pd.to_datetime(listaDispSemAgravo["DATA DISP."], dayfirst=True)

listaTotalDispensacoes = pd.read_excel("lista total dispensacoes.xlsx")
listaTotalDispensacoes = listaTotalDispensacoes.where(listaTotalDispensacoes["DATA DISP."] != "DATA DISP.").dropna(how='all')
listaTotalDispensacoes["DATA DISP."] = pd.to_datetime(listaTotalDispensacoes["DATA DISP."], dayfirst=True)

listaVacivida = pd.read_excel("vacivida.xlsx")
listaVacivida = listaVacivida.where(listaVacivida["DATA_APLICACAO_VACINA"] != "DATA_APLICACAO_VACINA")
listaVacivida["DATA_APLICACAO_VACINA"] = pd.to_datetime(listaVacivida["DATA_APLICACAO_VACINA"], dayfirst=True)

print("Comecei a ler a lista de saidas")
listaTotalSaidas = pd.read_excel("lista total saidas.xlsx")
listaTotalSaidas["DATA MOV"] = pd.to_datetime(listaTotalSaidas["DATA MOV"], dayfirst=True)
print("Terminei de ler a lista de saidas")

#print("Dia de hoje: " + str(paramDataAtual) + "\n" + str(paramTamanhoSemana) + " dias atrás: " + str(paramDataAtual - timedelta(days=paramTamanhoSemana)))

for unidade in listaUnidades:
    unidade["Suspeitas Total"] = qtdAgravos("SUSPEITA", unidade, False)
    unidade["Suspeitas Semana"] = qtdAgravos("SUSPEITA", unidade, True)
    unidade["Positivos Total"] = qtdAgravos("CONFIRMADO", unidade, False)
    unidade["Positivos Semana"] = qtdAgravos("CONFIRMADO", unidade, True)
    unidade["Negativos Total"] = qtdAgravos("NEGATIVO", unidade, False)
    unidade["Negativos Semana"] = qtdAgravos("NEGATIVO", unidade, True)
    unidade["Descartados Total"] = qtdAgravos("DESCARTADO", unidade, False)
    unidade["Descartados Semana"] = qtdAgravos("DESCARTADO", unidade, True)
    unidade["Sem Sintomas Total"] = semInicioSintomas(unidade, False)
    unidade["Sem Sintomas Semana"] = semInicioSintomas(unidade, True)
    unidade["Diferença Digitação x Notificação Total"] = difDigitNotif(unidade, False)
    unidade["Diferença Digitação x Notificação Semana"] = difDigitNotif(unidade, True)
    unidade["Dispensações sem Agravo Total"] = dispSemAgravo(unidade, False)
    unidade["Dispensações sem Agravo Semana"] = dispSemAgravo(unidade, True)
    unidade["Testes Dispensados Total"] = dispensacoes(unidade, False)
    unidade["Testes Dispensados Semana"] = dispensacoes(unidade, True)
    unidade["Vacivida Total"] = vacivida(unidade, False)
    unidade["Vacivida Semana"] = vacivida(unidade, True)
    unidade["Saída de Vacina (Aplicação) Total"] = saidaVacina(unidade, paramCodigoSaidaVacina, False)
    unidade["Saída de Vacina (Aplicação) Semana"] = saidaVacina(unidade, paramCodigoSaidaVacina, True)
    unidade["Aplicação Vacina Assessor - Vacivida (Semana)"] = unidade["Saída de Vacina (Aplicação) Semana"] -  unidade["Vacivida Semana"]
    
    wb = Workbook()
    ws = wb.active
    ws['A1'] = str(unidade["CNES"]) + " " + unidade["Nome"] + " - " + unidade["Apelido"]
    ws.merge_cells('A1:F1')
    ws['A3'] = "Dispensações sem Agravo (Total)"
    ws['B3'] = unidade["Dispensações sem Agravo Total"]
    ws['A4'] = "Dispensações sem Agravo (Semana)"
    ws['B4'] = unidade["Dispensações sem Agravo Semana"]
    
    ws['A6'] = "Notificações sem Data de Inicio de Sintomas (Total)"
    ws['B6'] = unidade["Sem Sintomas Total"]
    ws['A7'] = "Notificações sem Data de Inicio de Sintomas (Semana)"
    ws['B7'] = unidade["Sem Sintomas Semana"]
    
    ws['A9'] = "Notificações com diferença entre Data de Notificação e Data de Digitação (Total)"
    ws['B9'] = unidade["Diferença Digitação x Notificação Total"]
    ws['A10'] = "Notificações com diferença entre Data de Notificação e Data de Digitação (Semana)"
    ws['B10'] = unidade["Diferença Digitação x Notificação Semana"]
    
    ws['A12'] = "Notificações (Total)"
    ws['B12'] = "Suspeitas"
    ws['C12'] = "Positivas"
    ws['D12'] = "Negativas"
    ws['E12'] = "Descartadas"
    ws['F12'] = "Total"    
    ws['B13'] = unidade["Suspeitas Total"]
    ws['C13'] = unidade["Positivos Total"]
    ws['D13'] = unidade["Negativos Total"]
    ws['E13'] = unidade["Descartados Total"]
    somaNotifTotal = unidade["Suspeitas Total"] + unidade["Positivos Total"] + unidade["Negativos Total"] + unidade["Descartados Total"]    
    if somaNotifTotal > 0:
        porcSuspeitasTotal = (unidade["Suspeitas Total"] / somaNotifTotal) * 100
        porcPositivasTotal = (unidade["Positivos Total"] / somaNotifTotal) * 100
        porcNegativasTotal = (unidade["Negativos Total"] / somaNotifTotal) * 100
        porcDescartadasTotal = (unidade["Descartados Total"] / somaNotifTotal) * 100
        ws['B14'] = "{:.2f}".format(porcSuspeitasTotal) + "%"
        ws['C14'] = "{:.2f}".format(porcPositivasTotal) + "%"
        ws['D14'] = "{:.2f}".format(porcNegativasTotal) + "%"
        ws['E14'] = "{:.2f}".format(porcDescartadasTotal) + "%"
    ws['F13'] = somaNotifTotal
    
    ws['A16'] = "Notificações (Semana)"
    ws['B16'] = "Suspeitas"
    ws['C16'] = "Positivas"
    ws['D16'] = "Negativas"
    ws['E16'] = "Descartadas"
    ws['F16'] = "Total"
    ws['B17'] = unidade["Suspeitas Semana"]
    ws['C17'] = unidade["Positivos Semana"]
    ws['D17'] = unidade["Negativos Semana"]
    ws['E17'] = unidade["Descartados Semana"]    
    somaNotifSemana = unidade["Suspeitas Semana"] + unidade["Positivos Semana"] + unidade["Negativos Semana"] + unidade["Descartados Semana"]        
    if somaNotifSemana > 0:
        porcSuspeitasSemana = (unidade["Suspeitas Semana"] / somaNotifSemana) * 100    
        porcPositivasSemana = (unidade["Positivos Semana"] / somaNotifSemana) * 100
        porcNegativasSemana = (unidade["Negativos Semana"] / somaNotifSemana) * 100
        porcDescartadasSemana = (unidade["Descartados Semana"] / somaNotifSemana) * 100    
        ws['B18'] = "{:.2f}".format(porcSuspeitasSemana) + "%"
        ws['C18'] = "{:.2f}".format(porcPositivasSemana) + "%"
        ws['D18'] = "{:.2f}".format(porcNegativasSemana) + "%"    
        ws['E18'] = "{:.2f}".format(porcDescartadasSemana) + "%"
    ws['F17'] = somaNotifSemana
    
    
    ws['A20'] = "Aplicações Vacivida (Total)"
    ws['B20'] = unidade["Vacivida Total"]
    ws['A21'] = "Aplicações Assessor (Total)"
    ws['B21'] = unidade["Saída de Vacina (Aplicação) Total"]
    
    ws['A23'] = "Aplicações Vacivida (Semana)"
    ws['B23'] = unidade["Vacivida Semana"]
    ws['A24'] = "Aplicações Assessor (Semana)"
    ws['B24'] = unidade["Saída de Vacina (Aplicação) Semana"]
    ws['A25'] = "Diferença (Assessor - Vacivida) (Semana)"
    ws['B25'] = unidade["Aplicação Vacina Assessor - Vacivida (Semana)"]
    
    ws['A27'] = "Testes rápidos dispensados (Total)"
    ws['B27'] = unidade["Testes Dispensados Total"]
    ws['A28'] = "Testes rápidos dispensados (Semana)"
    ws['B28'] = unidade["Testes Dispensados Semana"]
    
    wb.save('Unidades/' + unidade['Apelido'] + " - " + paramDataAtras.strftime("%d.%m") + " a " + paramDataAtual.strftime("%d.%m") + ".xlsx")