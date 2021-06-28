# -*- coding: utf-8 -*-
"""
Created on Mon Jul 27 13:41:29 2020

@author: Info
"""

import pandas as pd
from datetime import datetime

tabelaTotal = pd.read_excel("lista total.xls").sort_values(by="Cód. Paciente", ignore_index=True)
suspeitos = tabelaTotal.where(tabelaTotal["Data da Notificação"] >= datetime.strptime('01-06-2020', '%d-%m-%Y'))
suspeitos = suspeitos.where(suspeitos["Situação"] == "SUSPEITA").dropna(how='all').sort_values(by=["Cód. Paciente", "Data da Notificação"]).drop_duplicates("Cód. Paciente", ignore_index=True)
suspeitos.to_excel('testeSUSPEITOS.xls')