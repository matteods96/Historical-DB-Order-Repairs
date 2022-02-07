# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 18:14:37 2020

@author: bonfpao
"""

import pandas as pd
import openpyxl 


#importo files di input

 gate=pd.read_excel('Gate_reports_hist_from_2012_to_2017 .xlsx')
gate.to_pickle('gate.pickle')
gate=pd.read_pickle('gate.pickle') #questo è il file gate contenente il job number che serve a fare il merge col file degli ordini mediante l'order number




