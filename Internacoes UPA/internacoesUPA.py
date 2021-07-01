# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 13:48:49 2021

@author: Info
"""
import pandas as pd
from datetime import datetime

def appendTabelaAuxiliar(tabela, row, motivo):
    aux = {"COD.": row[0], "NOME": row[1], "NASCIMENTO": row[2]}
    if motivo is not None:
        aux['Motivo'] = motivo
    tabela.append(aux)
    
paramDataAtual = datetime.today()

listaTotalUPA = pd.read_excel("planilhaUPA.xlsx").drop_duplicates(subset=['COD.', 'DATA']).sort_values(by=['COD.', 'DATA'])
listaIDs = listaTotalUPA[['COD.', 'NOME', 'NASCIMENTO']]
listaIDs = listaIDs.drop_duplicates('COD.')
listaIDs.insert(3, "Data Inicio", None)
listaIDs.insert(4, "Data Final", None)
listaIDs.insert(5, "Procedimentos", None)
listaIDs.insert(6, "Duração", None)
listaIDs.insert(7, "Dias por procedimento", None)

listaIDsIncorretos = []

for row in listaIDs.itertuples():
    listaAtendimentos = listaTotalUPA.where(listaTotalUPA['COD.'] == row[1]).dropna(how='all').reset_index()    
    if listaAtendimentos['COD.'].size == 1:
        appendTabelaAuxiliar(listaIDsIncorretos, row, None)
        listaIDs.drop(row.Index, inplace=True)
        continue
    listaIDs.at[row.Index, 'Data Inicio'] = listaAtendimentos.at[0, 'DATA']
    listaIDs.at[row.Index, 'Data Final'] = listaAtendimentos.at[listaAtendimentos['COD.'].size - 1, 'DATA']
    listaIDs.at[row.Index, 'Duração'] = (listaIDs.at[row.Index, 'Data Final'] - listaIDs.at[row.Index, 'Data Inicio']).days + 1
    listaIDs.at[row.Index, 'Procedimentos'] = listaAtendimentos['COD.'].size
    listaIDs.at[row.Index, 'Dias por procedimento'] = listaIDs.at[row.Index, 'Duração'] / listaIDs.at[row.Index, 'Procedimentos']
    
with pd.ExcelWriter('Internacoes UPA ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writer:
    listaIDs['Data Inicio'] = pd.to_datetime(listaIDs['Data Inicio'])
    listaIDs['Data Final'] = pd.to_datetime(listaIDs['Data Final'])
    listaIDs['NASCIMENTO'] = pd.to_datetime(listaIDs['NASCIMENTO'])
    listaIDs['Data Inicio'] = listaIDs['Data Inicio'].dt.strftime('%d/%m/%Y')
    listaIDs['Data Final'] = listaIDs['Data Final'].dt.strftime('%d/%m/%Y')
    listaIDs['NASCIMENTO'] = listaIDs['NASCIMENTO'].dt.strftime('%d/%m/%Y')
    listaIDs.to_excel(writer, index=False)
    
with pd.ExcelWriter('Internacoes UPA INCORRETO ' + paramDataAtual.strftime("%d.%m") + '.xlsx') as writerIncorreto:
    listaIDsIncorretos = pd.DataFrame(listaIDsIncorretos)
    listaIDsIncorretos.to_excel(writerIncorreto, index=False)