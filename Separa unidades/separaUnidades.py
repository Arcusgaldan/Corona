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
            filtro = coluna != key
            #print("Aplicando o primeiro filtro, key = " + key)
        else:
            #print("Aplicando o resto do filtro, key = " + key)
            filtro = filtro & (coluna != key)
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
'9662561': 'HANS',
'2092611': 'SANTA CASA',
'2090236': 'PIOXII',
'6316344': 'CENTRO DIAGNOSTICO MEDICO ASSOCIADO',
'2074176': 'CEDIB',
'3142531': 'HOSPITAL SAO JORGE',
'7909837': 'INSTITUTO RESPIRATORIO BARRETOS',
'5124700': 'LABORAMAS ',
'2052970': 'LABORATORIO ESTIMA',
'3549208': 'LABORATORIO SUZUKI',
'9567771': 'SAO FRANCISCO (Avenida 43)',
'9344209': 'SAO FRANCISCO (Avenida 7)',
'9371508': 'SAUDE MED',
"2043211": "CMR",
'419907': 'DROGA RAIA (Av.43)',
'120': 'Telessaude',
'yaah_cardoso@hotmail.com': 'DROGA RAIA (Frade Monte)',
'robertamap@yahoo.com.br': 'DROGA RAIA (Centro)',
'maicon.douglas19@hotmail.com': 'DROGARIA SÃO PAULO (City Barretos)',
'marcela@ltaseguranca.com.br': 'LTA ',
'UNIDADE BASICA DE SAUDE DO IBITU': 'Ibitu',
'UNIDADE BASICA DE SAUDE DR APOLONIO MORAES E SOUZA': 'Los Angeles',
'UNIDADE BASICA DR JOSE PARASSU CARVALHO': 'Ibirapuera',
'UNIDADE BASICA DE SAUDE DR LOTFALLAH MIZIARA': 'Barretos II',
'UNIDADE BASICA ARCHIMEDES MACHADO DE BARRETOS': 'Pimenta',
'UNIDADE BASICA DE SAUDE DR PAULO PRATA': 'Cristiano',
'UBS DR MILTON BARONI DE BARRETOS': 'America',
'UBS DR ALLY ALAHMAR DE BARRETOS': 'CSU',
'FUNDACAO PIO XII - BARRETOS': 'PIOXII',
'SANTA CASA DE BARRETOS': 'SANTA CASA',
'UNIDADE BASICA DE SAUDE DR WILSON HAYEK SAIHG': 'Cecapinha',
'UNIDADE BASICA DE SAUDE DR BARTOLOMEU MARAGLIANO V': 'Derby',
'AMBULATORIO DE ESPECIALIDADES ARE I': 'ARE I',
'UNIDADE BASICA DE SAUDE DR SERGIO PIMENTA': 'Marilia',
'UNIDADE BASICA DE SAUDE FRANCOLINO GALVAO DE SOUZA': 'Sao Francisco',
'UPA 24H ZAID ABRAO GERAIGE': 'UPA',
'UNIDADE ESF DR LUIZ SPINA': 'Spina',
'AMBULATORIO SAUDE DO IDOSO': 'Idoso',
'USF DO BAIRRO NOVA BARRETOS': 'Nova Barretos',
'HANS - HOSPITAL DE AMOR NOSSA SENHORA': 'HANS',
'SANTA CASA (TELESSAUDE)': 'Telessaude'
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

planilha.iloc[:, coluna - 1] = planilha.iloc[:, coluna - 1].apply(str)
        

if opcao == "1":
    for key in cnesUnidades:
        #planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == int(key)).dropna(how='all')
        planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == key).dropna(how='all')
        if not planilhaUnidade.empty:
            print("Achei uma planilha não vazia, unidade = " + cnesUnidades[key])
            with pd.ExcelWriter("Unidades/" + cnesUnidades[key] + '.xlsx') as writer:
                planilhaUnidade.to_excel(writer, index=False)
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
            #planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == int(key)).dropna(how='all')
            planilhaUnidade = planilha.where(planilha.iloc[:, coluna - 1] == key).dropna(how='all')
            if not planilhaUnidade.empty:
                print("Achei uma planilha não vazia, unidade = " + cnesUnidades[key])
                planilhaUnidade.to_excel(writer, index=False)
            else:
                print("Unidade veio vazia: " + cnesUnidades[key])
        filtro = filtroOutros(planilha.iloc[:, coluna - 1])
        planilhaUnidade = planilha.where(filtro).dropna(how='all')
        if not planilhaUnidade.empty:
            print("Existem unidades fora da lista de CNES")        
            planilhaUnidade.to_excel(writer, "Outras", index=False)
        else:
            print("Sem outras unidades")
        