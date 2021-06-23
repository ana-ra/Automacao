#!/usr/bin/env python
# coding: utf-8

# In[29]:


import pyautogui


# In[30]:


import time


# In[31]:


import pyperclip


# In[52]:


import pandas as pd


# In[53]:


pyautogui.PAUSE = 1
#vamos abrir o navegador
pyautogui.press("winleft")
pyautogui.PAUSE = 1
pyautogui.write("chrome")
pyautogui.PAUSE = 1
pyautogui.press("enter")
pyautogui.PAUSE = 1
pyautogui.alert("Vai começar, aperte OK e não mexa em nada!")
pyautogui.hotkey('ctrl','t')
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")
time.sleep(10)

#Vamos baixar a base de dados atualizada 
pyautogui.click(451,336, clicks =2)
time.sleep(3)
pyautogui.click(440,336, clicks = 1)
time.sleep(3)
pyautogui.click(1667, 196, clicks = 1)
time.sleep(3)
pyautogui.click(1439, 697, clicks = 1)
time.sleep(4)
pyautogui.press("enter")
time.sleep(10)
print(pyautogui.position())


# In[55]:


df = pd.read_excel(r'C:\Users\vitor\Downloads\Vendas - Dez.xlsx')
display(df)


# In[58]:


faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()
print(faturamento)
print(qtde_produtos)


# In[70]:


#Abrindo a aba do gmail
pyautogui.hotkey('ctrl', 't')
pyautogui.write("mail.google.com")
pyautogui.press('enter')
time.sleep(4)
pyautogui.click(261, 191, clicks = 1)
print(pyautogui.position())
#escrevendo o email 
time.sleep(4)
pyautogui.write('silvanarodriguesalvess@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl",'v')
pyautogui.press("tab")
texto = f"""
    Prezados, bom dia 
 O faturamento de ontem foi de: R$ {faturamento:,.2f}
 A quantidade de produtos foi de:{qtde_produtos:,}
 
 Abs 
 Atenciosamente Silvama"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", 'v')
pyautogui.hotkey('ctrl','enter')


# In[ ]:




