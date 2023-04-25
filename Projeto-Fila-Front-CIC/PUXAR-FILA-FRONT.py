import pyautogui
import pyperclip 
import time 
import pandas as pd
import win32com.client as win32
from PIL import ImageGrab
from PIL import Image
import pytesseract
import cv2
import re

## AVISO: HÁ MODIFICAÇÕES A SEREM FEITAS DEPOIS NAS LINHAS 290 E 311 !!!!!!!!!!!!!!!
## AVISO: A CADA HORA MUDAR O NOME DO ARQUIVO  DA LINHA 335 !!!!!!!!!!!!!!!!!!!!!!!!

##### ETAPA 1/3: PUXAR QUANTOS DOCS EM ABERTO HÁ EM CADA PASTA DO FRONT 

time.sleep(20)

SapGuiAuto = win32.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

session.findById("wnd[0]/tbar[0]/okcd").text = "YVFAX"

#PEGUEI O SCRIPT ABRIR-FAX

session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode ("F00002")
session.findById("wnd[1]/usr/lbl[0,2]").setFocus()
session.findById("wnd[1]/usr/lbl[0,2]").caretPosition = 14
session.findById("wnd[1]").sendVKey(2)

#PEGUEI O SCRIPT ABRIR-PEDIDOS

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6").select()
session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_BES").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").setFocus()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Capturar uma área específica da tela

left = 1196
top = 676
right = 1241
bottom = 695

time.sleep(1)

imagem = ImageGrab.grab((left, top, right, bottom))

time.sleep(1)

imagem.save("imagem.png")

novaimagem = cv2.imread("imagem.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
Pedidos_em_Aberto = pytesseract.image_to_string(imagem) 

Pedidos_em_Aberto_numeros = re.sub(r'[^0-9]', '', Pedidos_em_Aberto)

print("Pedidos em aberto", Pedidos_em_Aberto_numeros)

time.sleep(1)

#****************************************************************************************************************

# COTAÇÕES EM ABERTO

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_ANG").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

# Capturar uma área específica da tela

left = 1201
top = 672
right = 1239
bottom = 693

time.sleep(1)

imagem1 = ImageGrab.grab((left, top, right, bottom))

time.sleep(1)

imagem1.save("imagem1.png")

novaimagem1 = cv2.imread("imagem1.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
Cotacoes_em_Aberto = pytesseract.image_to_string(imagem1) 

# Remover letras usando regex
Cotacoes_em_Aberto_numeros = re.sub(r'[^0-9]', '', Cotacoes_em_Aberto)

# Imprimir resultado

print("Cotações em aberto", Cotacoes_em_Aberto_numeros)

time.sleep(1)

#****************************************************************************************************************

# OUTROS EM ABERTO

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_SON").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

left = 1222
top = 677
right = 1242
bottom = 694

time.sleep(1)

imagem2 = ImageGrab.grab((left, top, right, bottom))

time.sleep(1)

imagem2.save("imagem2.png")

novaimagem2 = cv2.imread("imagem2.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
Outros_em_Aberto = pytesseract.image_to_string(imagem2) 

# Remover letras usando regex
Outros_em_Aberto_numeros = re.sub(r'[^0-9]', '', Outros_em_Aberto)

# Imprimir resultado

print("Outros em aberto", Outros_em_Aberto_numeros)

time.sleep(1)

#****************************************************************************************************************

# CS EM ABERTO

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2224"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

left = 1206
top = 672
right = 1241
bottom = 694


time.sleep(1)

imagem3 = ImageGrab.grab((left, top, right, bottom))

time.sleep(1)

imagem3.save("imagem3.png")

novaimagem3 = cv2.imread("imagem3.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
CS_em_Aberto = pytesseract.image_to_string(imagem3) 

# Remover letras usando regex
CS_em_Aberto_numeros = re.sub(r'[^0-9]', '', CS_em_Aberto)

# Imprimir resultado

print("CS em aberto", CS_em_Aberto_numeros)

time.sleep(1)

#****************************************************************************************************************

# CORREÇÃO DPC EM ABERTO

#session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
#session.findById("wnd[1]/tbar[0]/btn[7]").press()
#session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_PROV").select()
#session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
#session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
#session.findById("wnd[1]/tbar[0]/btn[0]").press()

#left = 1216
#top = 674
#right = 1238
#bottom = 691

#time.sleep(1)

#imagem4 = ImageGrab.grab((left, top, right, bottom))

#time.sleep(1)

#imagem4.save("imagem4.png")

#novaimagem4 = cv2.imread("imagem4.png")

#caminho = r"C:\Program Files\Tesseract-OCR"
#pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
#CORRECAODPC_em_Aberto = pytesseract.image_to_string(imagem4) 

# Remover letras usando regex
#CORRECAODPC_em_Aberto_numeros = re.sub(r'[^0-9]', '', CORRECAODPC_em_Aberto)

# Imprimir resultado

#print("CORRECAODPC em aberto", CORRECAODPC_em_Aberto_numeros)

#time.sleep(1)

#****************************************************************************************************************

# DPC EM ABERTO

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "DPC"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 3
session.findById("wnd[1]/tbar[0]/btn[0]").press()

left = 1223
top = 676
right = 1237
bottom = 693

time.sleep(1)

imagem5 = ImageGrab.grab((left, top, right, bottom))

time.sleep(1)

imagem5.save("imagem5.png")

novaimagem5 = cv2.imread("imagem5.png")

caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'
DPC_em_Aberto = pytesseract.image_to_string(imagem5) 

# Remover letras usando regex
DPC_em_Aberto_numeros = re.sub(r'[^0-9]', '', DPC_em_Aberto)

# Imprimir resultado

print("DPC em aberto", DPC_em_Aberto_numeros)

time.sleep(2)


#****************************************************************************************************************
#****************************************************************************************************************
#****************************************************************************************************************

## FIM DA ETAPA 1

#****************************************************************************************************************
#****************************************************************************************************************
#****************************************************************************************************************

##### ETAPA 2/3: TRANSFERIR PARA O EXCEL OS DOCS EM ABERTO EM CADA PASTA

e = win32.Dispatch("Excel.Application")
e.Visible = 1
wb = e.Workbooks.Add()

pedidos = Pedidos_em_Aberto_numeros
cotacao = Cotacoes_em_Aberto_numeros
outros = Outros_em_Aberto_numeros
cs = 10 # CS_em_Aberto_numeros
correcaodpc = 20 # Substituir depois para: CORRECAODPC_em_Aberto_numeros ####################################################################
dpc = 30 # DPC_em_Aberto_numeros

strings = [Pedidos_em_Aberto_numeros, Cotacoes_em_Aberto_numeros, Outros_em_Aberto_numeros]

def extrair_numeros(string):
    numeros = re.findall(r'\d+', string)
    numeros_inteiros = [int(numero) for numero in numeros]
    return numeros_inteiros

# Variável para armazenar a soma dos números
soma = 0

# Iterar sobre as strings e somar os números
for string in strings:
    numeros = extrair_numeros(string)
    soma += sum(numeros)

wb.Sheets(1).Cells(8,2).Value = (soma)

wb.Sheets(1).Cells(1, 1).Value = "Fila Front"
wb.Sheets(1).Cells(1, 2).Value = "10h"  #tenho que alterar manualmente essa hora a cada hora (10h, 11h, 12h, etc)############################

wb.Sheets(1).Cells(2,1).Value = "Pedidos"
wb.Sheets(1).Cells(2,2).Value = pedidos

wb.Sheets(1).Cells(3,1).Value = "Cotação"
wb.Sheets(1).Cells(3,2).Value = cotacao

wb.Sheets(1).Cells(4,1).Value = "Outros"
wb.Sheets(1).Cells(4,2).Value = outros

wb.Sheets(1).Cells(5,1).Value = "CS"
wb.Sheets(1).Cells(5,2).Value = cs

wb.Sheets(1).Cells(6,1).Value = "Correção DPC"
wb.Sheets(1).Cells(6,2).Value = correcaodpc

wb.Sheets(1).Cells(7,1).Value = "DPC"
wb.Sheets(1).Cells(7,2).Value = dpc

wb.Sheets(1).Cells(8,1).Value = "TOTAL"

# salva o Workbook em um arquivo
wb.SaveAs(r'C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\excel-para-dashboard.xlsx')

# fecha o Workbook
wb.Close()

# encerra o Excel
e.Quit()

#****************************************************************************************************************
#****************************************************************************************************************
#****************************************************************************************************************

## FIM DA ETAPA 2

#****************************************************************************************************************
#****************************************************************************************************************
#****************************************************************************************************************

##### ETAPA 3/3: CONSTRUIR UM DASHBOARD COM OS DADOS DA ETAPA 2