# -*- coding: utf-8 -*-
"""
Created on Mon Jul  6 11:53:48 2020

@author: Info
"""


import pandas as pd

aux = pd.read_excel("confirmados.xls")
tabela = aux[["Cód. Paciente", "Paciente", "Agravo (CID)", "Data da Notificação"]].sort_values("Data da Notificação")
dup = tabela.drop_duplicates("Cód. Paciente", ignore_index=True)
dup.to_excel("confirmados_final.xls")
        
    