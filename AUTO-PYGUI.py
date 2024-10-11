#!/usr/bin/env python
# coding: utf-8

# In[11]:


import pandas as pd
import pyautogui as pag
from PIL import Image
import subprocess
import cv2
import time
import pyperclip
import win32com.client as win32


# In[9]:


pag.FAILSAFE = True

subprocess.Popen([r'C:\Program Files\Fakturama2\Fakturama.exe'])
def direita(posicao):
    return posicao[0] + posicao[2], posicao[1] + posicao[3]/2

def escrever_texto(texto):
    pyperclip.copy(texto)
    pag.hotkey('ctrl','v')

def encontrar_imagem(imagem):
    while True:
        try:
            posicao = pag.locateOnScreen(imagem, grayscale=True, confidence=0.9)
            if posicao:
                return posicao
                break
            else:
                print('não encontrado')
                time.sleep(1)
        except pag.ImageNotFoundException:  # Caso seja uma exceção personalizada no seu ambiente
            time.sleep(1)
            
pic = encontrar_imagem('faktu.png') 

tabela = pd.read_excel('Produtos.xlsx')           
       
            
for linha in tabela.index:
    
    
    
    id = tabela.loc[linha, 'ID']
    nome = tabela.loc[linha, 'Nome']
    categoria = tabela.loc[linha, 'Categoria']
    gtin = tabela.loc[linha, 'GTIN']
    supplier = tabela.loc[linha, 'Supplier']
    descricao = tabela.loc[linha, 'Descrição']
    imagem = tabela.loc[linha, 'Imagem']
    preco = tabela.loc[linha, 'Preço']
    custo = tabela.loc[linha, 'Custo']
    estoque = tabela.loc[linha, 'Estoque']
    

    pic = encontrar_imagem('botaonew.png')
    pag.click(pag.center(pic))

    pic = encontrar_imagem('newproduto.png')
    pag.click(pag.center(pic))

    pic = encontrar_imagem('itemclicar.png')
    pag.click(direita(pic))
    escrever_texto(str(id))
    pag.press('tab')

    escrever_texto(str(nome))
    pag.press('tab')

    escrever_texto(str(categoria))
    pag.press('tab')

    escrever_texto(str(gtin))
    pag.press('tab')

    escrever_texto(str(supplier))
    pag.press('tab')

    escrever_texto(str(descricao))
    pag.press('tab')

    preco_texto = f'{preco:.2f}'.replace('.',',')
    escrever_texto(str(preco_texto))
    pag.press('tab')

    custo_texto = f'{custo:.2f}'.replace('.',',')
    escrever_texto(custo_texto)

    #stock

    pic = encontrar_imagem('botaostock.png')
    pag.click(direita(pic))
    stock_texto = f'{estoque:.2f}'.replace('.',',')
    escrever_texto(stock_texto)
    
    pic = encontrar_imagem('selectPic.png')
    pag.click(pag.center(pic))
    
    pic = encontrar_imagem('arquivo_nome.png')
    escrever_texto(rf'C:\Users\Lucas B Maciel\Downloads\PYTHON BIBLIOTECA\AUTOMACAO PYAUTOGUI\ERP - FAKTA\Imagens Produtos\{str(imagem)}')
    pag.press('enter')

    pic = encontrar_imagem('botaosave.png')
    pag.click(pag.center(pic))
    
tabela[linha, 'Status'] = 'Registrado'
    
tabela.to_excel('Tabela atualizada.xlsx', index=False)
    


# In[18]:


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)


mail.Subject = 'Relatório de cadastro de produtos'
mail.Body = """
Lukeira
    
Segue abaixo relatório com Status atualizado
    
Abs
"""
attachment = r'C:\Users\Lucas B Maciel\Downloads\PYTHON BIBLIOTECA\AUTOMACAO PYAUTOGUI\ERP - FAKTA\Tabela atualizada.xlsx'
mail.Attachments.Add(attachment)
mail.To = 'lucasboni_business@outlook.com'  
mail.Send()


# In[7]:


display(tabela)
tabela[linha, 'Status'] = 'registrado'
display(tabela)


# In[ ]:




