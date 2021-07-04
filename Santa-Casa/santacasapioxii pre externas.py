# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 13:12:28 2020

@author: Info
"""


import pandas as pd
from datetime import timedelta
from datetime import datetime
import numpy as np
from timeit import default_timer as timer

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

def acharId(Cpf, Cns, Nome, Dn, tabela):
    #print("TESTE - DN param " + str(Dn) + "\nPrimeira data da tabela: " + str(tabela['Data Nasc.'][0]))
    if(type(Cpf) is float):
        print("ACHEI UM CPF FLOAT: " + str(Cpf) + "\nNome: " + str(Nome))
    if(Cpf is not None):
        filtro1 = tabela['CPF'] == Cpf.strip()
    else:
        filtro1 = False
        
    if(Cns is not None):
        filtro2 = tabela['CNS'] == Cns.strip()
    else:
        filtro2 = False
        
    if(Nome is not None and Dn is not None):
        filtro3 = tabela['Paciente'] == Nome.strip()
        filtro4 = tabela['Data Nasc.'] == Dn
    else:
        filtro3 = False
        filtro4 = False
    teste = tabela.where(filtro1 | filtro2 | (filtro3 & filtro4)).dropna(how='all')
    return teste

def mascaraDataEsus(tabelaEsus):
    if tabelaEsus.empty:
        return tabelaEsus
    colunasData = ['Data da NotificaCAo', 'Data de Nascimento', 'Data do inIcio dos sintomas']
    for i in colunasData:
        tabelaEsus[i] = tabelaEsus[i].dt.strftime('%d/%m/%Y')
    return tabelaEsus

def limpaUtf8(tabela): #Função que limpa os caracteres bugados do UTF-8 do arquivo de exportação do eSUS
    tabela.replace(to_replace={"Ã§": "C", "Ã‡": "C", "Ãµ": "O", "Ã³": "O", "Ã´": "O", "Ã\”": "O", "Ã­": "I", "Ãº": "U", "Ãš": "U", "ÃŠ": "E", "Ãª": "E", "Ã©": "E", "Ã‰": "E", "ÃG": "IG", "Ãƒ": "A", "Ã£": "A", "Ã¡": "A", "Ã¢": "A"}, inplace=True, regex=True)
    tabela.replace(to_replace={"Ã": "A"}, inplace=True, regex=True)
    tabela.rename(columns={"Data da NotificaÃ§Ã£o": "Data da NotificaCAo", "Resultado (PCR/RÃ¡pidos)": "Resultado (PCR/RApidos)", "MunicÃ­pio de ResidÃªncia": "MunicIpio de ResidEncia", "Data do inÃ­cio dos sintomas": "Data do inIcio dos sintomas", "ClassificaÃ§Ã£o Final": "ClassificaCAo Final"}, inplace=True)
    return tabela

def limpaAcentos(tabela): #Função que limpa os caracteres especiais da tabela do Assessor
    tabela.replace(to_replace={"Á": "A", "á": "A", "Â": "A", "â": "A", "Ã": "A", "ã": "A",
                               "É": "E", "é": "E", "Ê": "E", "ê": "E",
                               "Í": "I", "í": "I", "Î": "I", "î": "I",
                               "Ó": "O", "ó": "O", "Ô": "O", "ô": "O", "Õ": "O", "õ": "O",
                               "Ú": "U", "ú": "U", "Û": "U", "û": "U",
                               "Ç": "C", "ç": "C"}, inplace=True, regex=True)
    tabela.rename(columns={"Cód. Paciente": "COd. Paciente", "Data da Notificação": "Data da NotificaCAo", "Situação": "SituaCAo"}, inplace=True)    
    return tabela

def appendTabelaAuxiliar(tabela, row, motivo):
    aux = {"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Data Not.": row[paramColunaDataNotifEsus], "Notificante CNES": row[paramColunaCnesEsus]}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

paramDiasAtras = 16 #Parâmetro que define quantos dias atrás ele considera na planilha de suspeitos e monitoramento (ex: notificações de até X dias atrás serão analisadas, antes disso serão ignoradas)
paramDiasNegativo = 5 #Parâmetro que define quantos dias atrás uma notificacao de negativo deve ser considerada "a mesma" notificacao. Acima desse parâmetro deve ser considerada uma nova notificacao (agora também usada para descartados)
paramColunaNomeEsus = 30 ##Parâmetro que define qual é o index da coluna que possui o nome completo do paciente no eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaDataNascEsus = 25 ##Parâmetro que define qual é o index da coluna que possui a data de nascimento do paciente no eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaDataNotifEsus = 52 #Parâmetro que define qual é o index da coluna que possui a data de notificacao do eSUS (parâmetro necessário pois o itertuples não permite indexacao por nome de coluna com espaço)
paramColunaCnesEsus = 65
paramColunaDataNotifAssessor = 11 #Parâmetro igual ao paramColunaDataNotifEsus porém para o Assessor
paramDataAtual = datetime.today() #Parâmetro que define qual é a data atual para o script fazer a comparacao dos dias para trás (padrão: datetime.today() = data atual do sistema)

print("Data de hoje: " + paramDataAtual.strftime("%d/%m/%Y"))
print("Data de " + str(paramDiasAtras) + " dias atrás: " + (paramDataAtual - timedelta(days=paramDiasAtras)).strftime("%d/%m/%Y"))

start = timer()

tabelaTotalEsus = pd.read_excel("lista total esus.xlsx", dtype={'CPF': np.unicode_, 'CNS': np.unicode_, 'Nome Completo': np.unicode_}).sort_values(by="Data da NotificaÃ§Ã£o", ignore_index=True) #Lista total das notificações do eSus
tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS'] = formataCpfCns(tabelaTotalEsus['CPF'], tabelaTotalEsus['CNS']) #Formata as colunas de CNS e CPF conforme a regra de negócio
tabelaTotalEsus['Nome Completo'] = tabelaTotalEsus['Nome Completo'].str.upper() #Transforma a coluna de nome em Caixa Alta
tabelaTotalEsus = limpaUtf8(tabelaTotalEsus) #Limpa os caracteres que vem "bugados" do eSUS por conta de acentos e 'ç'
tabelaTotalEsus = tabelaTotalEsus.where(tabelaTotalEsus['MunicIpio de ResidEncia'] == "Barretos").dropna(how='all') #Filtra apenas pelo município de residência Barretos

endReadEsus = timer()
print("Terminei de ler a tabela do eSUS aos " + str(endReadEsus - start) + " segundos.")

filtroTotalPio1 = tabelaTotalEsus['Notificante CNES'] == 2090236
filtroTotalPio2 = tabelaTotalEsus['Notificante CNES'] == 9662561

filtroTotalOutras1 = tabelaTotalEsus['Notificante CNES'] == 9344209
filtroTotalOutras2 = tabelaTotalEsus['Notificante CNES'] == 2052970
filtroTotalOutras3 = tabelaTotalEsus['Notificante CNES'] == 9567771
filtroTotalOutras4 = tabelaTotalEsus['Notificante CNES'] == 9371508
filtroTotalOutras5 = tabelaTotalEsus['Notificante CNES'] == 3549208
filtroTotalOutras6 = tabelaTotalEsus['Notificante CNES'] == 7909837
filtroTotalOutras7 = tabelaTotalEsus['Notificante CNES'] == 3142531
filtroTotalOutras8 = tabelaTotalEsus['Notificante CNES'] == 5124700
filtroTotalOutras9 = tabelaTotalEsus['Notificante CNES'] == 6316344


tabelaTotalStc = tabelaTotalEsus.where(tabelaTotalEsus['Notificante CNES'] == 2092611).dropna(how='all') #Popula a tabela da Santa Casa com aquilo que foi notificado pelo CNES da Santa Casa
tabelaTotalPio = tabelaTotalEsus.where(filtroTotalPio1 | filtroTotalPio2).dropna(how='all') #Popula a tabela do Pio XII com aquilo que foi notificado pelo CNES do Pio XII
tabelaTotalOutras = tabelaTotalEsus.where(filtroTotalOutras1 | filtroTotalOutras2 | filtroTotalOutras3 | filtroTotalOutras4 | filtroTotalOutras5 | filtroTotalOutras6 | filtroTotalOutras7 | filtroTotalOutras8 | filtroTotalOutras9).dropna(how='all')

endSeparaEsus = timer()
print("Terminei de separar a tabela do eSUS por unidades aos " + str(endSeparaEsus - start) + " segundos.")

#Filtros para resultados positivos da Santa Casa
filtroPosStc1 = tabelaTotalStc["Resultado Totais"] == "Reagente"
filtroPosStc2 = tabelaTotalStc["Resultado IgA"] == "Reagente"
filtroPosStc3 = tabelaTotalStc["Resultado (PCR/RApidos)"] == "Positivo"
filtroPosStc4 = tabelaTotalStc["Resultado IgM"] == "Reagente"
filtroPosStc5 = tabelaTotalStc["Resultado IgG"] == "Reagente"

#Filtros para resultados positivos do Pio XII
filtroPosPio1 = tabelaTotalPio["Resultado Totais"] == "Reagente"
filtroPosPio2 = tabelaTotalPio["Resultado IgA"] == "Reagente"
filtroPosPio3 = tabelaTotalPio["Resultado (PCR/RApidos)"] == "Positivo"
filtroPosPio4 = tabelaTotalPio["Resultado IgM"] == "Reagente"
filtroPosPio5 = tabelaTotalPio["Resultado IgG"] == "Reagente"

#Filtros para resultados positivos das Outras Unidades
filtroPosOut1 = tabelaTotalOutras["Resultado Totais"] == "Reagente"
filtroPosOut2 = tabelaTotalOutras["Resultado IgA"] == "Reagente"
filtroPosOut3 = tabelaTotalOutras["Resultado (PCR/RApidos)"] == "Positivo"
filtroPosOut4 = tabelaTotalOutras["Resultado IgM"] == "Reagente"
filtroPosOut5 = tabelaTotalOutras["Resultado IgG"] == "Reagente"

#Filtros para resultados negativos da Santa Casa
filtroNegStc1 = tabelaTotalStc["Resultado Totais"] != "Reagente"
filtroNegStc2 = tabelaTotalStc["Resultado IgA"] != "Reagente"
filtroNegStc3 = tabelaTotalStc["Resultado (PCR/RApidos)"] != "Positivo"
filtroNegStc4 = tabelaTotalStc["Resultado IgM"] != "Reagente"
filtroNegStc5 = tabelaTotalStc["Resultado IgG"] != "Reagente"
filtroConcluidoStc = tabelaTotalStc["Estado do Teste"] == "ConcluIdo"

#Filtros para resultados negativos do Pio XII
filtroNegPio1 = tabelaTotalPio["Resultado Totais"] != "Reagente"
filtroNegPio2 = tabelaTotalPio["Resultado IgA"] != "Reagente"
filtroNegPio3 = tabelaTotalPio["Resultado (PCR/RApidos)"] != "Positivo"
filtroNegPio4 = tabelaTotalPio["Resultado IgM"] != "Reagente"
filtroNegPio5 = tabelaTotalPio["Resultado IgG"] != "Reagente"
filtroConcluidoPio = tabelaTotalPio["Estado do Teste"] == "ConcluIdo"

#Filtros para resultados negativos das Outras Unidades
filtroNegOut1 = tabelaTotalOutras["Resultado Totais"] != "Reagente"
filtroNegOut2 = tabelaTotalOutras["Resultado IgA"] != "Reagente"
filtroNegOut3 = tabelaTotalOutras["Resultado (PCR/RApidos)"] != "Positivo"
filtroNegOut4 = tabelaTotalOutras["Resultado IgM"] != "Reagente"
filtroNegOut5 = tabelaTotalOutras["Resultado IgG"] != "Reagente"
filtroConcluidoOut = tabelaTotalOutras["Estado do Teste"] == "ConcluIdo"

#Filtros para resultados vazios e não descartados da Santa Casa
filtroVazioStc1 = pd.isnull(tabelaTotalStc["Resultado Totais"])
filtroVazioStc2 = pd.isnull(tabelaTotalStc["Resultado IgA"])
filtroVazioStc3 = pd.isnull(tabelaTotalStc["Resultado (PCR/RApidos)"])
filtroVazioStc4 = pd.isnull(tabelaTotalStc["Resultado IgM"])
filtroVazioStc5 = pd.isnull(tabelaTotalStc["Resultado IgG"])
filtroNaoDescartadoStc = tabelaTotalStc["ClassificaCAo Final"] != "Descartado"

#Filtros para resultados vazios e não descartados do Pio XII
filtroVazioPio1 = pd.isnull(tabelaTotalPio["Resultado Totais"])
filtroVazioPio2 = pd.isnull(tabelaTotalPio["Resultado IgA"])
filtroVazioPio3 = pd.isnull(tabelaTotalPio["Resultado (PCR/RApidos)"])
filtroVazioPio4 = pd.isnull(tabelaTotalPio["Resultado IgM"])
filtroVazioPio5 = pd.isnull(tabelaTotalPio["Resultado IgG"])
filtroNaoDescartadoPio = tabelaTotalPio["ClassificaCAo Final"] != "Descartado"

#Filtros para resultados vazios e não descartados das Outras Unidades
filtroVazioOut1 = pd.isnull(tabelaTotalOutras["Resultado Totais"])
filtroVazioOut2 = pd.isnull(tabelaTotalOutras["Resultado IgA"])
filtroVazioOut3 = pd.isnull(tabelaTotalOutras["Resultado (PCR/RApidos)"])
filtroVazioOut4 = pd.isnull(tabelaTotalOutras["Resultado IgM"])
filtroVazioOut5 = pd.isnull(tabelaTotalOutras["Resultado IgG"])
filtroNaoDescartadoOut = tabelaTotalOutras["ClassificaCAo Final"] != "Descartado"

#Aplica os filtros positivos definidos acima
tabelaPositivosStc = tabelaTotalStc.where(filtroPosStc1 | filtroPosStc2 | filtroPosStc3 | filtroPosStc4 | filtroPosStc5).dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaPositivosPio = tabelaTotalPio.where(filtroPosPio1 | filtroPosPio2 | filtroPosPio3 | filtroPosPio4 | filtroPosPio5).dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaPositivosOut = tabelaTotalOutras.where(filtroPosOut1 | filtroPosOut2 | filtroPosOut3 | filtroPosOut4 | filtroPosOut5).dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')

endPositivos = timer()
print("Terminei de separar os positivos aos " + str(endPositivos - start) + " segundos.")

#Aplica os filtros negativos definidos acima
tabelaNegativosStc = tabelaTotalStc.where(filtroConcluidoStc & filtroNegStc1 & filtroNegStc2 & filtroNegStc3 & filtroNegStc4 & filtroNegStc5).dropna(how='all')
#tabelaNegativosStc = tabelaNegativosStc.where(tabelaNegativosStc["Data da NotificaCAo"] >= (paramDataAtual - timedelta(days=paramDiasAtras))).dropna(how='all') #Pega os negativos apenas de 15 dias atrás pra frente
tabelaNegativosPio = tabelaTotalPio.where(filtroConcluidoPio & filtroNegPio1 & filtroNegPio2 & filtroNegPio3 & filtroNegPio4 & filtroNegPio5).dropna(how='all')
#tabelaNegativosPio = tabelaNegativosPio.where(tabelaNegativosPio["Data da NotificaCAo"] >= (paramDataAtual - timedelta(days=paramDiasAtras))).dropna(how='all') #Pega os negativos apenas de 15 dias atrás pra frente
tabelaNegativosOut = tabelaTotalOutras.where(filtroConcluidoOut & filtroNegOut1 & filtroNegOut2 & filtroNegOut3 & filtroNegOut4 & filtroNegOut5).dropna(how='all')

endNegativos = timer()
print("Terminei de separar os negativos aos " + str(endNegativos - start) + " segundos.")

#Aplica os filtros de vazios e não descartados
tabelaSemResultadoStc = tabelaTotalStc.where(filtroVazioStc1 & filtroVazioStc2 & filtroVazioStc3 & filtroVazioStc4 & filtroVazioStc5 & filtroNaoDescartadoStc).dropna(how='all')
tabelaSemResultadoPio = tabelaTotalPio.where(filtroVazioPio1 & filtroVazioPio2 & filtroVazioPio3 & filtroVazioPio4 & filtroVazioPio5 & filtroNaoDescartadoPio).dropna(how='all')
tabelaSemResultadoOut = tabelaTotalOutras.where(filtroVazioOut1 & filtroVazioOut2 & filtroVazioOut3 & filtroVazioOut4 & filtroVazioOut5 & filtroNaoDescartadoOut).dropna(how='all')

#Puxa apenas os suspeitos baseado em resultado vazio, não descartado e estado 'Coletado'
tabelaSuspeitosStc = tabelaSemResultadoStc.where(tabelaTotalStc["Estado do Teste"] == "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaSuspeitosStc = tabelaSuspeitosStc.where(tabelaSuspeitosStc["Data da NotificaCAo"] >= (paramDataAtual - timedelta(days=paramDiasAtras))).dropna(how='all')
tabelaSuspeitosPio = tabelaSemResultadoPio.where(tabelaTotalPio["Estado do Teste"] == "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaSuspeitosPio = tabelaSuspeitosPio.where(tabelaSuspeitosPio["Data da NotificaCAo"] >= (paramDataAtual - timedelta(days=paramDiasAtras))).dropna(how='all')
tabelaSuspeitosOut = tabelaSemResultadoOut.where(tabelaTotalOutras["Estado do Teste"] == "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaSuspeitosOut = tabelaSuspeitosOut.where(tabelaSuspeitosOut["Data da NotificaCAo"] >= (paramDataAtual - timedelta(days=paramDiasAtras))).dropna(how='all')

endSuspeitos = timer()
print("Terminei de separar os suspeitos aos " + str(endSuspeitos - start) + " segundos.")

#Puxa apenas os monitoramentos baseado em resultado vazio, não descartado e estado diferente de 'Coletado'
tabelaMonitoramentoStc = tabelaSemResultadoStc.where(tabelaSemResultadoStc["Estado do Teste"] != "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaMonitoramentoStc = tabelaMonitoramentoStc.where(tabelaMonitoramentoStc["Data da NotificaCAo"] >= paramDataAtual - timedelta(days=paramDiasAtras)).dropna(how='all')
tabelaMonitoramentoPio = tabelaSemResultadoPio.where(tabelaSemResultadoPio["Estado do Teste"] != "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaMonitoramentoPio = tabelaMonitoramentoPio.where(tabelaMonitoramentoPio["Data da NotificaCAo"] >= paramDataAtual - timedelta(days=paramDiasAtras)).dropna(how='all')
tabelaMonitoramentoOut = tabelaSemResultadoOut.where(tabelaSemResultadoOut["Estado do Teste"] != "Coletado").dropna(how='all').drop_duplicates("CPF", ignore_index=True, keep='last')
tabelaMonitoramentoOut = tabelaMonitoramentoOut.where(tabelaMonitoramentoOut["Data da NotificaCAo"] >= paramDataAtual - timedelta(days=paramDiasAtras)).dropna(how='all')

endMonitoramento = timer()
print("Terminei de separar os monitoramentos aos " + str(endMonitoramento - start) + " segundos.")

#Puxa a tabela do Assessor
tabelaTotalAssessor = pd.read_excel("lista total assessor.xls", dtype={'CPF': np.unicode_, 'CNS': np.unicode_}).sort_values(by=["Cód. Paciente", "Data da Notificação"]) #Lê a tabela do Assessor e classifica por ID do paciente e Data da Notificação
tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS'] = formataCpfCns(tabelaTotalAssessor['CPF'], tabelaTotalAssessor['CNS']) #Formata as colunas de CPF e CNS baseado na regra de negócio
tabelaTotalAssessor = limpaAcentos(tabelaTotalAssessor) #Limpa acentos e 'ç' da tabela do Assessor

endReadAssessor = timer()
print("Terminei de ler a tabela do Assessor aos " + str(endReadAssessor - start) + " segundos.")

tabelaMonitoramentoStcFalso = []
tabelaSuspeitosStcFalso = []
tabelaPositivosStcFalso = []
tabelaNegativosStcFalso = []

tabelaMonitoramentoPioFalso = []
tabelaSuspeitosPioFalso = []
tabelaPositivosPioFalso = []
tabelaNegativosPioFalso = []

tabelaMonitoramentoOutFalso = []
tabelaSuspeitosOutFalso = []
tabelaPositivosOutFalso = []
tabelaNegativosOutFalso = []

for row in tabelaMonitoramentoStc.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaMonitoramentoStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaMonitoramentoStcFalso, row, "CONFIRMADO")
        tabelaMonitoramentoStc.drop(row.Index, inplace=True)
        continue
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #tabelaMonitoramentoStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "SUSPEITO EM ABERTO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaMonitoramentoStcFalso, row, "SUSPEITO EM ABERTO")
        tabelaMonitoramentoStc.drop(row.Index, inplace=True)
        continue
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaMonitoramentoStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaMonitoramentoStcFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaMonitoramentoStc.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaMonitoramentoStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaMonitoramentoStcFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaMonitoramentoStc.drop(row.Index, inplace=True)
            continue

for row in tabelaSuspeitosStc.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaSuspeitosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaSuspeitosStcFalso, row, "CONFIRMADO")
        tabelaSuspeitosStc.drop(row.Index, inplace=True)
        continue
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #tabelaSuspeitosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "SUSPEITO EM ABERTO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaSuspeitosStcFalso, row, "SUSPEITO EM ABERTO")
        tabelaSuspeitosStc.drop(row.Index, inplace=True)
        continue
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaSuspeitosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaSuspeitosStcFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaSuspeitosStc.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaSuspeitosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaSuspeitosStcFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaSuspeitosStc.drop(row.Index, inplace=True)
            continue
        
for row in tabelaPositivosStc.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaPositivosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaPositivosStcFalso, row, "CONFIRMADO")
        tabelaPositivosStc.drop(row.Index, inplace=True)
        continue
    
for row in tabelaNegativosStc.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaNegativosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaNegativosStcFalso, row, "CONFIRMADO")
        tabelaNegativosStc.drop(row.Index, inplace=True)
        continue    
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaNegativosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaNegativosStcFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaNegativosStc.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaNegativosStcFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaNegativosStcFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaNegativosStc.drop(row.Index, inplace=True)
            continue
        
endStc = timer()
print("Terminei de processar a Santa Casa aos " + str(endStc - start) + " segundos.")
        
for row in tabelaMonitoramentoPio.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaMonitoramentoPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaMonitoramentoPioFalso, row, "CONFIRMADO")
        tabelaMonitoramentoPio.drop(row.Index, inplace=True)
        continue
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #tabelaMonitoramentoPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "SUSPEITO EM ABERTO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaMonitoramentoPioFalso, row, "SUSPEITO EM ABERTO")
        tabelaMonitoramentoPio.drop(row.Index, inplace=True)
        continue
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaMonitoramentoPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaMonitoramentoPioFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaMonitoramentoPio.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaMonitoramentoPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaMonitoramentoPioFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaMonitoramentoPio.drop(row.Index, inplace=True)
            continue

for row in tabelaSuspeitosPio.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaSuspeitosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaSuspeitosPioFalso, row, "CONFIRMADO")
        tabelaSuspeitosPio.drop(row.Index, inplace=True)
        continue
    if "SUSPEITA" in notifAssessor["SituaCAo"].values:
        #tabelaSuspeitosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "SUSPEITO EM ABERTO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaSuspeitosPioFalso, row, "SUSPEITO EM ABERTO")
        tabelaSuspeitosPio.drop(row.Index, inplace=True)
        continue
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaSuspeitosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaSuspeitosPioFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaSuspeitosPio.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaSuspeitosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaSuspeitosPioFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaSuspeitosPio.drop(row.Index, inplace=True)
            continue
        
for row in tabelaPositivosPio.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaPositivosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaPositivosPioFalso, row, "CONFIRMADO")
        tabelaPositivosPio.drop(row.Index, inplace=True)
        continue
    
for row in tabelaNegativosPio.itertuples():
    notifAssessor = acharId(row.CPF, row.CNS, row[paramColunaNomeEsus], row[paramColunaDataNascEsus], tabelaTotalAssessor)
    # print("TESTE: " + str(row[paramColunaDataNascEsus]))
    if "CONFIRMADO" in notifAssessor["SituaCAo"].values:
        #print("Achei um confirmado!")
        #tabelaNegativosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "CONFIRMADO", "Data Not.": row[paramColunaDataNotifEsus]})
        appendTabelaAuxiliar(tabelaNegativosPioFalso, row, "CONFIRMADO")
        tabelaNegativosPio.drop(row.Index, inplace=True)
        continue    
    if "NEGATIVO" in notifAssessor["SituaCAo"].values:
        negativosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "NEGATIVO").dropna(how='all')
        for rowNegativo in negativosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver negativo no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowNegativo[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaNegativosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "NEGATIVO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus]})
            appendTabelaAuxiliar(tabelaNegativosPioFalso, row, "NEGATIVO DENTRO DO PERIODO")
            tabelaNegativosPio.drop(row.Index, inplace=True)
            continue
    if "DESCARTADO" in notifAssessor["SituaCAo"].values:
        descartadosAssessor = notifAssessor.where(notifAssessor["SituaCAo"] == "DESCARTADO").dropna(how='all')
        for rowDescartado in descartadosAssessor.itertuples():
            flagIsWrong = False #Flag que indica se a notificacao do eSUS é incorreta por já haver descartado no período indicado pelo parâmetro. Flag usada para sair corretamente dos laços.
            dataNotifAssessor = rowDescartado[paramColunaDataNotifAssessor]
            dataNotifEsus = row[paramColunaDataNotifEsus]
            if dataNotifAssessor >= dataNotifEsus - timedelta(days=paramDiasNegativo):
                flagIsWrong = True
                break
        if flagIsWrong:
            #tabelaNegativosPioFalso.append({"Nome": row[paramColunaNomeEsus], "CPF": row.CPF, "Motivo": "DESCARTADO DENTRO DO PERIODO", "Data Not.": row[paramColunaDataNotifEsus], "Notificante CNES": row[paramColunaCnesEsus]})
            appendTabelaAuxiliar(tabelaNegativosPioFalso, row, "DESCARTADO DENTRO DO PERIODO")
            tabelaNegativosPio.drop(row.Index, inplace=True)
            continue
        
endPio = timer()
print("Terminei de processar o Pio XII e HANS aos " + str(endPio - start) + " segundos.")
        
#for row in tabelaPositivosOut.itertuples():
    
        
tabelaMonitoramentoHans = tabelaMonitoramentoPio.where(tabelaMonitoramentoPio['Notificante CNES'] == 9662561).dropna(how='all')
tabelaSuspeitosHans = tabelaSuspeitosPio.where(tabelaSuspeitosPio['Notificante CNES'] == 9662561).dropna(how='all')
tabelaPositivosHans = tabelaPositivosPio.where(tabelaPositivosPio['Notificante CNES'] == 9662561).dropna(how='all')
tabelaNegativosHans = tabelaNegativosPio.where(tabelaNegativosPio['Notificante CNES'] == 9662561).dropna(how='all')

tabelaMonitoramentoPioFalso = pd.DataFrame(tabelaMonitoramentoPioFalso)
tabelaSuspeitosPioFalso = pd.DataFrame(tabelaSuspeitosPioFalso)
tabelaPositivosPioFalso = pd.DataFrame(tabelaPositivosPioFalso)
tabelaNegativosPioFalso = pd.DataFrame(tabelaNegativosPioFalso)

tabelaMonitoramentoHansFalso = pd.DataFrame()
tabelaSuspeitosHansFalso = pd.DataFrame()
tabelaPositivosHansFalso = pd.DataFrame()
tabelaNegativosHansFalso = pd.DataFrame()

if not tabelaMonitoramentoPioFalso.empty:
    tabelaMonitoramentoHansFalso = tabelaMonitoramentoPioFalso.where(tabelaMonitoramentoPioFalso['Notificante CNES'] == 9662561).dropna(how='all')    
if not tabelaSuspeitosPioFalso.empty:
    tabelaSuspeitosHansFalso = tabelaSuspeitosPioFalso.where(tabelaSuspeitosPioFalso['Notificante CNES'] == 9662561).dropna(how='all')
if not tabelaPositivosPioFalso.empty:
    tabelaPositivosHansFalso = tabelaPositivosPioFalso.where(tabelaPositivosPioFalso['Notificante CNES'] == 9662561).dropna(how='all')
if not tabelaNegativosPioFalso.empty:    
    tabelaNegativosHansFalso = tabelaNegativosPioFalso.where(tabelaNegativosPioFalso['Notificante CNES'] == 9662561).dropna(how='all')

tabelaMonitoramentoPio = tabelaMonitoramentoPio.where(tabelaMonitoramentoPio['Notificante CNES'] == 2090236).dropna(how='all')
tabelaSuspeitosPio = tabelaSuspeitosPio.where(tabelaSuspeitosPio['Notificante CNES'] == 2090236).dropna(how='all')
tabelaPositivosPio = tabelaPositivosPio.where(tabelaPositivosPio['Notificante CNES'] == 2090236).dropna(how='all')
tabelaNegativosPio = tabelaNegativosPio.where(tabelaNegativosPio['Notificante CNES'] == 2090236).dropna(how='all')

if not tabelaMonitoramentoPioFalso.empty:
    tabelaMonitoramentoPioFalso = tabelaMonitoramentoPioFalso.where(tabelaMonitoramentoPioFalso['Notificante CNES'] == 2090236).dropna(how='all')
if not tabelaSuspeitosPioFalso.empty:
    tabelaSuspeitosPioFalso = tabelaSuspeitosPioFalso.where(tabelaSuspeitosPioFalso['Notificante CNES'] == 2090236).dropna(how='all')
if not tabelaPositivosPioFalso.empty:    
    tabelaPositivosPioFalso = tabelaPositivosPioFalso.where(tabelaPositivosPioFalso['Notificante CNES'] == 2090236).dropna(how='all')
if not tabelaNegativosPioFalso.empty:
    tabelaNegativosPioFalso = tabelaNegativosPioFalso.where(tabelaNegativosPioFalso['Notificante CNES'] == 2090236).dropna(how='all')
    
endHans = timer()
print("Terminei de separar HANS da Santa Casa aos " + str(endHans - start) + " segundos.")

with pd.ExcelWriter('Santa Casa ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerStcCorreto:
    tabelaMonitoramentoStc = mascaraDataEsus(tabelaMonitoramentoStc)    
    tabelaMonitoramentoStc.to_excel(writerStcCorreto, "Monitoramento")    
    
    tabelaSuspeitosStc = mascaraDataEsus(tabelaSuspeitosStc)    
    tabelaSuspeitosStc.to_excel(writerStcCorreto, "Suspeitos")    
    
    tabelaPositivosStc = mascaraDataEsus(tabelaPositivosStc)    
    tabelaPositivosStc.to_excel(writerStcCorreto, "Positivos")    
    
    tabelaNegativosStc = mascaraDataEsus(tabelaNegativosStc)    
    tabelaNegativosStc.to_excel(writerStcCorreto, "Negativos")
    
with pd.ExcelWriter('PIO XII ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerPioCorreto:
    tabelaMonitoramentoPio = mascaraDataEsus(tabelaMonitoramentoPio)    
    tabelaMonitoramentoPio.to_excel(writerPioCorreto, "Monitoramento")    
    
    tabelaSuspeitosPio = mascaraDataEsus(tabelaSuspeitosPio)    
    tabelaSuspeitosPio.to_excel(writerPioCorreto, "Suspeitos")    
    
    tabelaPositivosPio = mascaraDataEsus(tabelaPositivosPio)    
    tabelaPositivosPio.to_excel(writerPioCorreto, "Positivos")    
    
    tabelaNegativosPio = mascaraDataEsus(tabelaNegativosPio)    
    tabelaNegativosPio.to_excel(writerPioCorreto, "Negativos")
    
with pd.ExcelWriter('HANS ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerHansCorreto:
    tabelaMonitoramentoHans = mascaraDataEsus(tabelaMonitoramentoHans)    
    tabelaMonitoramentoHans.to_excel(writerHansCorreto, "Monitoramento")    
    
    tabelaSuspeitosHans = mascaraDataEsus(tabelaSuspeitosHans)    
    tabelaSuspeitosHans.to_excel(writerHansCorreto, "Suspeitos")    
    
    tabelaPositivosHans = mascaraDataEsus(tabelaPositivosHans)    
    tabelaPositivosHans.to_excel(writerHansCorreto, "Positivos")
    
    tabelaNegativosHans = mascaraDataEsus(tabelaNegativosHans)    
    tabelaNegativosHans.to_excel(writerHansCorreto, "Negativos") 
    
with pd.ExcelWriter('INCORRETO ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerIncorreto:
    tabelaMonitoramentoStcFalso = pd.DataFrame(tabelaMonitoramentoStcFalso)
    tabelaMonitoramentoStcFalso.to_excel(writerIncorreto, "Monitoramento StaCasa")    
    tabelaMonitoramentoPioFalso.to_excel(writerIncorreto, "Monitoramento Pio XII")    
    tabelaMonitoramentoHansFalso.to_excel(writerIncorreto, "Monitoramento Hans")
    
    tabelaSuspeitosStcFalso = pd.DataFrame(tabelaSuspeitosStcFalso)
    tabelaSuspeitosStcFalso.to_excel(writerIncorreto, "Suspeitos StaCasa")    
    tabelaSuspeitosPioFalso.to_excel(writerIncorreto, "Suspeitos Pio XII")    
    tabelaSuspeitosHansFalso.to_excel(writerIncorreto, "Suspeitos Hans")
    
    tabelaPositivosStcFalso = pd.DataFrame(tabelaPositivosStcFalso)
    tabelaPositivosStcFalso.to_excel(writerIncorreto, "Positivos StaCasa")    
    tabelaPositivosPioFalso.to_excel(writerIncorreto, "Positivos Pio XII")    
    tabelaPositivosHansFalso.to_excel(writerIncorreto, "Positivos Hans")
    
    tabelaNegativosStcFalso = pd.DataFrame(tabelaNegativosStcFalso)
    tabelaNegativosStcFalso.to_excel(writerIncorreto, "Negativos StaCasa")    
    tabelaNegativosPioFalso.to_excel(writerIncorreto, "Negativos Pio XII")    
    tabelaNegativosHansFalso.to_excel(writerIncorreto, "Negativos Hans")
    
endWrite = timer()
print("Terminei de escrever as planilhas aos " + str(endWrite - start) + " segundos.")