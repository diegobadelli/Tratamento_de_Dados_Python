#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from time import sleep
import matplotlib.pyplot as plt
from scipy.stats import norm
import numpy as np
from datetime import date


# In[ ]:


def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()


# In[ ]:


Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
planilha = askopenfilename(initialdir = "/", title = "Selecione a planilha") # show an "Open" dialog box and return the path to the selected file
printProgressBar(10, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


#Leitura dos dados
dados = pd.DataFrame(pd.read_excel(planilha))
printProgressBar(15, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


dados


# In[ ]:


dados = dados.drop(dados.index[[0, 1, 2, 3]])
printProgressBar(20, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


#Renomeando colunas
dados.columns = ['CÓDIGO', 'TIPO ECD', 'LOTE', 'LANCE', 'EQUIPAMENTO', 'CAIXA', 'DATA','OBSERVAÇÃO','Máquina AZ','Código IS AZ','Haste de Vergalhão AZ','Fornecedor Vergalhão AZ','Máquina AZ/BR','Código IS AZ/BR','Haste de Vergalhão AZ/BR','Fornecedor Vergalhão AZ/BR','Máquina AZC','Código IS AZC','Haste de Vergalhão AZC','Fornecedor Vergalhão AZC','Máquina LA','Código IS LA','Haste de Vergalhão LA','Fornecedor Vergalhão LA','Máquina LA/BR','Código IS LA/BR','Haste de Vergalhão LA/BR','Fornecedor Vergalhão LA/BR','Máquina BR','Código IS BR','Haste de Vergalhão BR','Fornecedor Vergalhão BR','Máquina VD','Código IS VD','Haste de Vergalhão VD','Fornecedor Vergalhão VD','Máquina VD/BR','Código IS VD/BR','Haste de Vergalhão VD/BR','Fornecedor Vergalhão VD/BR','Máquina VDC','Código IS VDC','Haste de Vergalhão VDC','Fornecedor Vergalhão VDC','Máquina MR','Código IS MR','Haste de Vergalhão MR','Fornecedor Vergalhão MR','Máquina MR/BR','Código IS MR/BR','Haste de Vergalhão MR/BR','Fornecedor Vergalhão MR/BR','Máquina MRC','Código IS MRC','Haste de Vergalhão MRC','Fornecedor Vergalhão MRC','Máquina BR1','Código IS BR1','Haste de Vergalhão BR1','Fornecedor Vergalhão BR1','Máquina BR2','Código IS BR2','Haste de Vergalhão BR2','Fornecedor Vergalhão BR2','Máquina BR3','Código IS BR3','Haste de Vergalhão BR3','Fornecedor Vergalhão BR3','Máquina BR4','Código IS BR4','Haste de Vergalhão BR4','Fornecedor Vergalhão BR4','Código BI AZ/AZ/BR','Máquina AZ/AZ/BR','Código BI LA/LA/BR','Máquina LA/LA/BR','Código BI VD/VD/BR','Máquina VD/VD/BR','Código BI MR/MR/BR','Máquina MR/MR/BR','Código BI AZ/BR/AZ','Máquina AZ/BR/AZ','Código BI LA/BR/LA','Máquina LA/BR/LA','Código BI VD/BR/VD','Máquina VD/BR/VD','Código BI MR/BR/MR','Máquina MR/BR/MR','Código BI AZ/AZC','Máquina AZ/AZC','Código BI LA/BR','Máquina LA/BR','Código BI VD/VDC','Máquina VD/VDC','Código BI MR/MRC','Máquina MR/MRC','Código BI AZC/AZ','Máquina AZC/AZ','Código BI BR/LA','Máquina BR/LA','Código BI VDC/VD','Máquina VDC/VD','Código BI MRC/MR','Máquina MRC/MR','Código BI AZ/BR1','Máquina AZ/BR1','Código BI LR/BR2','Máquina LR/BR2','Código BI VD/BR3','Máquina VD/BR3','Código BI MR/BR4','Máquina MR/BR4','Código BI BR1/AZ','Máquina BR1/AZ','Código BI BR2/LA','Máquina BR2/LA','Código BI BR3/VD','Máquina BR3/VD','Código BI BR4/MR','Máquina BR4/MR','Código BI BR/AZ','Máquina BR/AZ','Código BI BR/VD','Máquina BR/VD','Código BI BR/MR','Máquina BR/MR','Máquina de Reunião','Código RE','Máquina de Capa','Máquina de Corte','Freq_Att 1', 'Att 1', 'Freq_Att 2', 'Att 2', 'Freq_Att 3', 'Att 3', 'Freq_Att 4', 'Att 4', 'Freq_RL 1', 'RL 1', 'Freq_RL 2', 'RL 2', 'Freq_RL 3', 'RL 3', 'Freq_RL 4', 'RL 4', 'Freq_NEXT 1x2', 'NEXT 1x2', 'Freq_NEXT 1x3', 'NEXT 1x3', 'Freq_NEXT 1x4', 'NEXT 1x4', 'Freq_NEXT 2x1', 'NEXT 2x1', 'Freq_NEXT 2x3', 'NEXT 2x3', 'Freq_NEXT 2x4', 'NEXT 2x4', 'Freq_NEXT 3x1','NEXT 3x1', 'Freq_NEXT 3x2', 'NEXT 3x2', 'Freq_NEXT 3x4', 'NEXT 3x4', 'Freq_NEXT 4x1', 'NEXT 4x1', 'Freq_NEXT 4x2', 'NEXT 4x2', 'Freq_NEXT 4x3', 'NEXT 4x3', 'Freq_FEXT 1x2','FEXT 1x2', 'Freq_FEXT 1x3', 'FEXT 1x3', 'Freq_FEXT 1x4', 'FEXT 1x4', 'Freq_FEXT 2x1', 'FEXT 2x1', 'Freq_FEXT 2x3', 'FEXT 2x3', 'Freq_FEXT 2x4', 'FEXT 2x4', 'Freq_FEXT 3x1','FEXT 3x1', 'Freq_FEXT 3x2', 'FEXT 3x2', 'Freq_FEXT 3x4', 'FEXT 3x4', 'Freq_FEXT 4x1', 'FEXT 4x1', 'Freq_FEXT 4x2', 'FEXT 4x2', 'Freq_FEXT 4x3', 'FEXT 4x3', 'Freq_ACRF 1x2',
'ACRF 1x2', 'Freq_ACRF 1x3', 'ACRF 1x3', 'Freq_ACRF 1x4', 'ACRF 1x4', 'Freq_ACRF 2x1', 'ACRF 2x1', 'Freq_ACRF 2x3', 'ACRF 2x3', 'Freq_ACRF 2x4', 'ACRF 2x4', 'Freq_ACRF 3x1','ACRF 3x1', 'Freq_ACRF 3x2', 'ACRF 3x2', 'Freq_ACRF 3x4', 'ACRF 3x4', 'Freq_ACRF 4x1', 'ACRF 4x1', 'Freq_ACRF 4x2', 'ACRF 4x2', 'Freq_ACRF 4x3', 'ACRF 4x3', 'Freq_ELFEXT 1x2',
'ELFEXT 1x2', 'Freq_ELFEXT 1x3', 'ELFEXT 1x3', 'Freq_ELFEXT 1x4', 'ELFEXT 1x4', 'Freq_ELFEXT 2x1', 'ELFEXT 2x1', 'Freq_ELFEXT 2x3', 'ELFEXT 2x3', 'Freq_ELFEXT 2x4', 'ELFEXT 2x4','Freq_ELFEXT 3x1', 'ELFEXT 3x1', 'Freq_ELFEXT 3x2', 'ELFEXT 3x2', 'Freq_ELFEXT 3x4', 'ELFEXT 3x4', 'Freq_ELFEXT 4x1', 'ELFEXT 4x1', 'Freq_ELFEXT 4x2', 'ELFEXT 4x2', 'Freq_ELFEXT 4x3',
'ELFEXT 4x3', 'Freq_ACR 1', 'ACR 1', 'Freq_APR 1', 'APR 1', 'Freq_APR 2', 'APR 2', 'Freq_APR 3', 'APR 3', 'Freq_APR 4', 'APR 4', 'Freq_DS 1', 'DS 1', 'Freq_DS 1x2', 'DS 1x2', 'Freq_DS 1x3', 'DS 1x3', 'Freq_DS 1x4', 'DS 1x4', 'Freq_DS 2x3', 'DS 2x3', 'Freq_DS 2x4', 'DS 2x4', 'Freq_DS 3x4', 'DS 3x4', 'Freq_PSNEXT 1', 'PSNEXT 1', 'Freq_PSNEXT 2', 'PSNEXT 2',
'Freq_PSNEXT 3', 'PSNEXT 3', 'Freq_PSNEXT 4', 'PSNEXT 4', 'Freq_IMP_LIN 1', 'IMP_LIN 1', 'Freq_IMP_LIN 2', 'IMP_LIN 2', 'Freq_IMP_LIN 3', 'IMP_LIN 3', 'Freq_IMP_LIM 4', 'IMP_LIM 4','Freq_AT._FASE 1', 'AT._FASE 1', 'Freq_AT._FASE 2', 'AT._FASE 2', 'Freq_AT._FASE 3', 'AT._FASE 3', 'Freq_AT._FASE 4', 'AT._FASE 4', 'Freq_PSELFEXT 1', 'PSELFEXT 1', 'Freq_PSELFEXT 2',
'PSELFEXT 2', 'Freq_PSELFEXT 3', 'PSELFEXT 3', 'Freq_PSELFEXT 4', 'PSELFEXT 4', 'Freq_PSACRF 1', 'PSACRF 1', 'Freq_PSACRF 2', 'PSACRF 2', 'Freq_PSACRF 3', 'PSACRF 3', 'Freq_PSACRF 4','PSACRF 4', 'Freq_SRL 1', 'SRL 1', 'Freq_SRL 2', 'SRL 2', 'Freq_SRL 3', 'SRL 3', 'Freq_SRL 4', 'SRL 4', 'Freq_NEXT FAR 1x2', 'NEXT FAR 1x2', 'Freq_NEXT FAR 1x3', 'NEXT FAR 1x3',
'Freq_NEXT FAR 1x4', 'NEXT FAR 1x4', 'Freq_NEXT FAR 2x3', 'NEXT FAR 2x3', 'Freq_NEXT FAR 2x4', 'NEXT FAR 2x4', 'Freq_NEXT FAR 3x4', 'NEXT FAR 3x4', 'Freq_IMP ZO 1', 'IMP ZO 1','Freq_IMP ZO 2',  'IMP ZO 2', 'Freq_IMP ZO 3', 'IMP ZO 3', 'Freq_IMP ZO 4', 'IMP ZO 4', 'Freq_FRL 1', 'FRL 1', 'Freq_FRL 2', 'FRL 2', 'Freq_FRL 3', 'FRL 3', 'Freq_FRL 4', 'FRL 4',
'DR 1x2', 'DR 1', 'DR 2x1', 'DR 2x3', 'DR 2', 'DR 3x4', 'DR 3', 'DR 4x1', 'DR 4', 'CAP_MT 1x2', 'CAP_MT 1', 'CAP_MT 2x1', 'CAP_MT 2x3', 'CAP_MT 2', 'CAP_MT 3x4', 'CAP_MT 3', 'CAP_MT 4x1','CAP_MT 4', 'DC_PP 1x2', 'DC_PP 1', 'DC_PP 2x1', 'DC_PP 2x3', 'DC_PP 3x4', 'DC_PP 4x1', 'DC_PT 1x2', 'DC_PT 1', 'DC_PT 2x1', 'DC_PT 2x3', 'DC_PT 2', 'DC_PT 3x4', 'DC_PT 3', 'DC_PT 4x1', 'DC_PT 4', 'RES._C._A 1x2', 'RES._C._A 1', 'RES._C._A 2x1', 'RES._C._A 2x3', 'RES._C._A 2', 'RES._C._A 3x4', 'RES._C._A 3', 'RES._C._A 4x1',
'RES._C._A 4', 'RES._C._B 1x2', 'RES._C._B 1', 'RES._C._B 2x1','RES._C._B 2x3', 'RES._C._B 2', 'RES._C._B 3x4', 'RES._C._B 3', 'RES._C._B 4x1', 'RES._C._B 4']
printProgressBar(40, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


##Retirando Frequências
dados_sem_freq = dados.loc[:,~dados.columns.str.startswith('Freq')]
printProgressBar(70, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


##Retirando valores não utilizados
lista = ['2x1','3x1','3x2','4x1','4x2','4x3']
dados_sem_vals = dados_sem_freq
for i in range(len(lista)):
    dados_sem_vals = dados_sem_vals.loc[:,~dados_sem_vals.columns.str.endswith(lista[i])]
printProgressBar(80, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


#Retirando dados nulos ou não finitos
dados_limpos = dados_sem_vals.dropna(axis=1, thresh=10)
printProgressBar(90, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)


# In[ ]:


#Salvando uma cópia da planilha tratada
plan_final = asksaveasfilename(initialdir = "/", title = "Escolha o caminho para salvar a planilha tratada", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
dados_limpos.to_excel(plan_final + ".xlsx", index = False, sheet_name="Resultado")
printProgressBar(100, 100, prefix = 'Progresso:', suffix = 'Completo', length = 50)

