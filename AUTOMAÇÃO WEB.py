#!/usr/bin/env python
# coding: utf-8

# In[13]:


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time 
import yagmail
from tkinter import *
import tkinter.filedialog
from tkinter import messagebox


# In[ ]:





# In[23]:


driver = webdriver.Chrome()
produtos_df = pd.read_excel('Produtos.xlsx')
display(produtos_df)
precos = {}
iphones = []
tv = []
lista_aux = []
lojas = {}

for c in produtos_df:
    if c not in lista_aux:
        lista_aux.append(c)
        if c == 'Amazon':
            produto = {}
            for cont, i in enumerate(produtos_df[c]):
                driver.get(produtos_df['Amazon'][cont])
                time.sleep(5)
                elemento = driver.find_elements_by_class_name('a-color-price')
                valor = str(elemento[0].text)
                valor = valor.replace('R$ ', '').replace(',', '-').replace('.', '').replace('-', '.').replace(':', '').replace('(', '')
                if valor.isnumeric():
                    valor = float(valor)
                else:
                    valor = 0
                produto[produtos_df.loc[cont, 'Link Produto']] = valor
            lojas['Amazon'] = produto
        if c == 'Lojas Americanas':
            produto = {}
            for cont, i in enumerate(produtos_df[c]):
                driver.get(produtos_df['Lojas Americanas'][cont])
                time.sleep(5)
                try:
                    try:
                        elemento = driver.find_element_by_class_name('priceSales')
                        valor = str(elemento.text)
                    except:
                        elemento = driver.find_element_by_class_name('fbVrjL')
                        valor = str(elemento.text)
                except:
                    janela = Tk()
                    messagebox.showinfo('Automação Web', 'Programa parado por um Captcha, resolva-o, para continuar')
                    janela.destroy()
                    while len(driver.find_elements_by_class_name('logo')) == 0:
                        time.sleep(1)
                    time.sleep(7)
                    try:
                        elemento = driver.find_element_by_class_name('priceSales')
                        valor = str(elemento.text)
                    except:
                        elemento = driver.find_element_by_class_name('fbVrjL')
                        valor = str(elemento.text)
                valor = str(elemento.text)
                valor = valor.replace('R$ ', '').replace(',', '-').replace('.', '').replace('-', '.').replace(':', '').replace('(', '')
                if valor.isnumeric():
                    valor = float(valor)
                else:
                    valor = 0
                produto[produtos_df.loc[cont, 'Link Produto']] = valor
            lojas['Lojas Americanas'] = produto
        if c == 'Magazine Luiza':
            produto = {}
            for cont, i in enumerate(produtos_df[c]):
                driver.get(produtos_df['Magazine Luiza'][cont])
                time.sleep(5)
                try:
                    try:
                        elemento = driver.find_element_by_class_name('price-template__text')
                        valor = str(elemento.text)
                    except:
                        elemento = driver.find_element_by_class_name('unavailable__product-title')
                        valor = str(elemento.text)
                except:
                    janela = Tk()
                    messagebox.showinfo('Automação Web', 'Programa parado por um Captcha, resolva-o, para continuar')
                    janela.destroy()
                    while len(driver.find_elements_by_class_name('header-lu-image')) == 0:
                        time.sleep(1)
                    time.sleep(7)
                    try:
                        elemento = driver.find_element_by_class_name('price-template__text')
                        valor = str(elemento.text)
                    except:
                        elemento = driver.find_element_by_class_name('unavailable__product-title')
                        valor = str(elemento.text)
                valor = str(elemento.text)
                valor = valor.replace('R$ ', '').replace(',', '-').replace('.', '').replace('-', '.').replace(':', '').replace('(', '')
                if valor.isnumeric():
                    valor = float(valor)
                else:
                    valor = 0
                produto[produtos_df.loc[cont, 'Link Produto']] = valor
            lojas['Magazine Luiza'] = produto


# In[15]:


print(lojas)


# In[24]:


menor_iphone = menor_tv = cont = 0
menor = {}
menores = {}
menores_lojas = {}
for l, c in lojas.items():
    print(c)
    if cont == 0:
        menor_iphone = c['iPhone 11 Apple 64GB Preto']
    if c['iPhone 11 Apple 64GB Preto'] < menor_iphone:
        menor_iphone = c['iPhone 11 Apple 64GB Preto']
   
    if cont == 0:
        menor_tv = c["Smart TV LED 50'' LG Ultra HD 4K Thinq AI"]
    if c["Smart TV LED 50'' LG Ultra HD 4K Thinq AI"] < menor_tv:
        menor_tv = c["Smart TV LED 50'' LG Ultra HD 4K Thinq AI"]

for l, c in lojas.items():
    if c['iPhone 11 Apple 64GB Preto'] == menor_iphone:
        menores[f'{l}-iphone'] =  menor_iphone
    if c["Smart TV LED 50'' LG Ultra HD 4K Thinq AI"] == menor_tv:
        menores[f'{l}-tv'] =  menor_tv

        
temp = []
melhor_preco = dict()
for key, val in menores.items():
    if val not in temp:
        temp.append(val)
        melhor_preco[key] = val
temp = list(reversed(temp))
print(temp)
print(melhor_preco)
print(menor_iphone)
print(menor_tv)


# In[25]:


desconto = {}
for i, c in enumerate(produtos_df['Preço Original']):
    c = int(c)
    desconto[produtos_df.loc[i, 'Link Produto']] = c - (c * 0.2)
print(desconto)


# In[26]:


for i, c in enumerate(produtos_df['Preço Atual']):
    if temp[i] != 0:
        produtos_df.loc[i, 'Preço Atual'] = temp[i]
    else:
        produtos_df.loc[i, 'Preço Atual'] = produtos_df.loc[i, 'Preço Original'] 
    for x, v in melhor_preco.items():
        if v == temp[i]:
            if 'Magazine' in x:
                produtos_df.loc[i, 'Local'] = f'{x[:-7]}'
            else:
                produtos_df.loc[i, 'Local'] = f'{x[:-3]}'
        if c < 100000000000:
            print(c)
            usuario = yagmail.SMTP(user='malvesbruno0+cliente@gmail.com', password='BMA9191F')
            usuario.send(to='malvesbruno@gmail.com', subject=f'Item {produtos_df.loc[i, "Link Produto"]} em promoção',                         contents=f'''<html>

<p>O produto <strong>{produtos_df.loc[i, 'Link Produto']}</strong> teve um abaixo de preço, confira: </p>


<table>
  <tr>
      <th>Produto</th>
      <th>Valor Original</th>
      <th>Preço Atual</th>
      <th>Local</th>
  </tr>
  <tr>
    <td style="text-align: center">{produtos_df.loc[i, 'Link Produto']}</td>
    <td style="text-align: center">R${produtos_df.loc[i, 'Preço Original']:,.2f}</td>
    <td style="text-align: center">R${produtos_df.loc[i, 'Preço Atual']}</td>
    <td style="text-align: center">{produtos_df.loc[i, 'Local']}</td>
  </tr>  
</table>

<p>Segue o link do site</p>
</html>
''')
display(produtos_df)


# In[ ]:




