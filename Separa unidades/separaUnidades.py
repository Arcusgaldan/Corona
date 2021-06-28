# -*- coding: utf-8 -*-
"""
Created on Thu Jun 24 11:53:47 2021

@author: Info
"""
import pandas as pd

def filtroOutros(coluna):
    filtro = pd.Series()
    for key in cnesUnidades:
        if filtro.empty:
            filtro = coluna != int(key)
            #print("Aplicando o primeiro filtro, key = " + key)
        else:
            #print("Aplicando o resto do filtro, key = " + key)
            filtro = filtro & (coluna != int(key))
            #teste = filtro.where(filtro == False).dropna(how='all')
            #print("Falsos = " + str(teste.size))
    return filtro
        

cnesUnidades = {"2035731": "Ibitu",
"2048736": "Los Angeles",
"2048744": "Ibirapuera",
"2053314": "Barretos II",
"2056062": "Pimenta",
"2061473": "Cristiano",
"2064081": "America",
"2064103": "CSU",
"2093642": "Cecapinha",
"2093650": "Derby",
"2784580": "Marilia",
"2784599": "Sao Francisco",
"7035861": "UPA",
"7122217": "Spina",
"7565577": "Idoso",
"7585020": "Nova Barretos",
"2784572": "ARE I",
"2043211": "CMR"
}

planilha = pd.read_excel("planilha.xlsx")
while True:
    print("Insira o número da coluna com o CNES da unidade (1 é a primeira, 2 é a segunda etc.)")
    try:
        coluna = int(input())
    except:
        print("Valor inválido")
        continue
    break

while True:
    print("1 - Arquivos separados\n2 - Mesmo arquivo, abas separadas\n\nSelecione: ")
    opcao = input()
    if opcao != "1" and opcao != "2":
        print("Valor inválido")
        continue
    break
        

if opcao == "1":
    for key in cnesUnidades:
        planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == int(key)).dropna(how='all')
        if not planilhaUnidade.empty:
            print("Achei uma planilha não vazia, unidade = " + cnesUnidades[key])
            with pd.ExcelWriter("Unidades/" + cnesUnidades[key] + '.xlsx') as writer:
                planilhaUnidade.to_excel(writer, cnesUnidades[key], index=False)
        else:
            print("Unidade veio vazia: " + cnesUnidades[key])
    filtro = filtroOutros(planilha.iloc[:, coluna - 1])
    planilhaUnidade = planilha.where(filtro).dropna(how='all')
    if not planilhaUnidade.empty:
        print("Existem unidades fora da lista de CNES")
        with pd.ExcelWriter('Unidades/Outras.xlsx') as writer:
            planilhaUnidade.to_excel(writer, "Outras", index=False)
    else:
        print("Sem outras unidades")
        
else:
    with pd.ExcelWriter('Planilha Organizada.xlsx') as writer:
        for key in cnesUnidades:
            planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == int(key)).dropna(how='all')
            if not planilhaUnidade.empty:
                print("Achei uma planilha não vazia, unidade = " + cnesUnidades[key])
                planilhaUnidade.to_excel(writer, cnesUnidades[key], index=False)
            else:
                print("Unidade veio vazia: " + cnesUnidades[key])
        filtro = filtroOutros(planilha.iloc[:, coluna - 1])
        planilhaUnidade = planilha.where(filtro).dropna(how='all')
        if not planilhaUnidade.empty:
            print("Existem unidades fora da lista de CNES")        
            planilhaUnidade.to_excel(writer, "Outras", index=False)
        else:
            print("Sem outras unidades")
        