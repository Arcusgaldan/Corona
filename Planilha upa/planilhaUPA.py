# -*- coding: utf-8 -*-
"""
Created on Mon Jun 21 13:52:09 2021

@author: Info
"""

import pandas as pd
from datetime import timedelta

def appendTabelaAuxiliar(tabela, row, motivo):
    aux = {"PACIENTE": row.PACIENTE, "ID": row.ID, "COLETA": row.COLETA, "RESULTADO": row.RESULTADO}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)

tabelaUPA = pd.read_excel("planilhaUPA.xlsx")
tabelaAssessor = pd.read_excel("lista total assessor.xlsx")
tabelaUPAIncorreto = []

paramColunaDataNotifAssessor = 11
paramDiasNegativo = 10

for row in tabelaUPA.itertuples():
    agravos = tabelaAssessor.where(tabelaAssessor["Cód. Paciente"] == row.ID).dropna(how='all')
    if(row.RESULTADO == "POSITIVO"):
        if(agravos.empty or "CONFIRMADO" not in agravos["Situação"].values):
            continue
        else:
            appendTabelaAuxiliar(tabelaUPAIncorreto, row, "Já existe positivo no Assessor")
            tabelaUPA.drop(row.Index, inplace=True)
    elif(row.RESULTADO == "NEGATIVO"):
        agravosSuspeitos = agravos.where(agravos["Situação"] == "SUSPEITA")
        if agravosSuspeitos.empty:
            appendTabelaAuxiliar(tabelaUPAIncorreto, row, "Não tem suspeita nenhuma")
            tabelaUPA.drop(row.Index, inplace=True)
        else:
            flagTrazer = False
            for rowSuspeito in agravosSuspeitos.itertuples():
                dataNotifAssessor = rowSuspeito[paramColunaDataNotifAssessor]
                dataTeste = row.COLETA
                if dataTeste >= dataNotifAssessor and dataNotifAssessor >= dataTeste - timedelta(days=paramDiasNegativo):
                    flagTrazer = True
                    break
            if not flagTrazer:
                appendTabelaAuxiliar(tabelaUPAIncorreto, row, "Não tem suspeita até 10 dias antes do teste")
                tabelaUPA.drop(row.Index, inplace=True)
                
with pd.ExcelWriter('Resultado Analise Planilha UPA.xlsx') as writerUPA:    
    tabelaUPA.to_excel(writerUPA, index=False)
    
with pd.ExcelWriter('Incorreto Analise Planilha UPA.xlsx') as writerIncorreto:
    tabelaUPAIncorreto = pd.DataFrame(tabelaUPAIncorreto)
    tabelaUPAIncorreto.to_excel(writerIncorreto, index=False)    
    


            