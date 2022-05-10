#!/usr/bin/env python
# coding: utf-8

# # Relatório Estatístico
# #### Autor: Diego do Carmo Badelli

# In[ ]:


import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import matplotlib.pyplot as plt
from scipy.stats import norm
import numpy as np
from datetime import date


# In[ ]:


d1 = str(date.today().strftime("%d/%m/%Y"))
print("Data: {}".format(d1))


# In[ ]:


Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
planilha = askopenfilename(initialdir = "/", title = "Selecione a planilha") # show an "Open" dialog box and return the path to the selected file


# In[ ]:


#Leitura dos dados
dados = pd.DataFrame(pd.read_excel(planilha))


# ## Boxplots e Histogramas

# In[ ]:


def Parametro(nome, atributo):
    nome = dados.loc[:, dados.columns.str.startswith(atributo)]
    dados_limpos = nome.dropna()
    return dados_limpos


# In[ ]:


IL = Parametro ('IL', "Att")
RL = Parametro ('RL', 'RL')
NEXT = Parametro ('NEXT', 'NEXT')
ELFEXT = Parametro ('ELFEXT', 'ACRF')


# In[ ]:


boxplot_Il = IL.boxplot()


# In[ ]:


boxplot_RL = RL.boxplot()


# In[ ]:


boxplot_NEXT = NEXT.boxplot()


# In[ ]:


boxplot_ELFEXT = ELFEXT.boxplot()


# In[ ]:


# Generate some data for this demonstration
def Histonorm(data):
    dados = data
    for i in data:
        # Fit a normal distribution to the data:
        mu, std = norm.fit(i)

        # Plot the histogram.
        plt.hist(i, bins=25, density=True, alpha=0.6, color='g')

        # Plot the PDF.
        xmin, xmax = plt.xlim()
        x = np.linspace(xmin, xmax, 100)
        p = norm.pdf(x, mu, std)
        plt.plot(x, p, 'k', linewidth=2)
        title = "%s /n Média = %.2f,  Desvio Padrão = %.2f" % (str(i), mu, std)
        plt.title(title)
        plt.grid(True)
        plt.show()


# In[ ]:


Histonorm([IL['Att 1'], IL['Att 2'], IL['Att 3'], IL['Att 4']])


# In[ ]:


Histonorm([RL['RL 1'],RL['RL 2'],RL['RL 3'],RL['RL 4']])


# In[ ]:


# Generate some data for this demonstration
Histonorm([NEXT['NEXT 1x2'], NEXT['NEXT 1x3'], NEXT['NEXT 1x4'], NEXT['NEXT 2x3'], NEXT['NEXT 2x4'], NEXT['NEXT 3x4']])


# In[ ]:


# Generate some data for this demonstration
Histonorm([ELFEXT['ACRF 1x2'], ELFEXT['ACRF 1x3'], ELFEXT['ACRF 1x4'], ELFEXT['ACRF 2x3'], ELFEXT['ACRF 2x4'], ELFEXT['ACRF 3x4']])

