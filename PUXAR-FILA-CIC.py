import pyautogui
import pandas as pd
import time 
import datetime
from datetime import date
import win32com.client as win32
from tkinter import Tk
import getpass
import csv
import pyperclip
import clipboard
import re
from colorama import init, Fore, Style
import os.path
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys  
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.common.exceptions import TimeoutException

# Inicializar colorama
init()

root = Tk() # O codigo está usando TK para criar uma aplicação em segundo plano para armazenar a informação que é obtida dentro do control C. Se fosse pyautogui controlc e controlv, esse TK não seria necessário 
root.withdraw()  

# ABRIR O SAP

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

##################################################################################################################################
############################################### ENTRADA DE DOCUMENTOS DO DIA #####################################################
##################################################################################################################################

# obter data atual
data_atual = date.today()

# formatar a data como string e com "." ao invés de "/" (04.05.2023 ao invés de 04/05/2023)
data_formatada = data_atual.strftime("%d.%m.%Y")

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6").select()
session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/ctxtO_DATUM-LOW").text = data_formatada
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").setFocus
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

pyautogui.moveTo(1227, 684)
pyautogui.tripleClick(1227, 684)
pyautogui.hotkey('ctrl','c')

valor1 = str(root.clipboard_get())
#valor1 = '1-10 de 134'


entrada_de_documentos = int(str(valor1[valor1.find('de'):]).replace('de ', ''))
#entrada_de_documentos = 134

print('Entrada de Documentos:', entrada_de_documentos)

##################################################################################################################################
##################################################### PEDIDOS ####################################################################
##################################################################################################################################

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

clipboard.copy('0')

if not pyautogui.moveTo(1227, 684):  
    pyautogui.moveTo(1227, 684)
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor2 = str(root.clipboard_get())

    pedidos_em_aberto = int(str(valor2[valor2.find('de'):]).replace('de ', ''))

    if pedidos_em_aberto <= 180:
            print('\033[32mPedidos em aberto:', pedidos_em_aberto, '\033[m')
    elif pedidos_em_aberto > 180:
            print('\033[31mPedidos em aberto:', pedidos_em_aberto, '\033[m')
    else:
        0        

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')  

##################################################################################################################################
##################################################### COTAÇÃO ####################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_ANG").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()


try:        
    pyautogui.moveTo(1227, 684)
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor3 = str(root.clipboard_get())

    cotacoes_em_aberto = int(str(valor3[valor3.find('de'):]).replace('de ', ''))

    if cotacoes_em_aberto <= 80:
        print('\033[32mCotações em aberto:', cotacoes_em_aberto, '\033[m')
    else:
        print('\033[31mCotações em aberto:', cotacoes_em_aberto, '\033[m')

except:
    cotacoes_em_aberto = 0
    print('\033[32mCotações em aberto:', cotacoes_em_aberto, '\033[m')
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

##################################################################################################################################
##################################################### OUTROS #####################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_SON").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
clipboard.copy('0')
if not pyautogui.moveTo(1227, 684):  
    pyautogui.moveTo(1227, 684)
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor4 = str(root.clipboard_get())

    outros_em_aberto = int(str(valor4[valor4.find('de'):]).replace('de ', ''))

    if outros_em_aberto <= 10:
        print('\033[32mOutros em aberto:', outros_em_aberto, '\033[m')
    elif outros_em_aberto > 10:
        print('\033[31mOutros em aberto:', outros_em_aberto, '\033[m')
    else:
        0

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')  

##################################################################################################################################
########################################################## CS ####################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2224"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

clipboard.copy('0')

if not pyautogui.moveTo(1227, 684):
    pyautogui.moveTo(1227, 684)
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor5 = str(root.clipboard_get())

    cs_em_aberto = int(str(valor5[valor5.find('de'):]).replace('de ', ''))     

    if cs_em_aberto <= 5:
        print('\033[32mCS em aberto:', cs_em_aberto, '\033[m')
    elif cs_em_aberto>5:
        print('\033[31mCS em aberto:', cs_em_aberto, '\033[m')
    else:
        0    

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')  

##################################################################################################################################
########################################################## DEALERS ####################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2227"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

clipboard.copy('0')

if not pyautogui.moveTo(1227, 684):  # Verifica se o cursor do mouse não pode ser movido para a posição especificada
    pyautogui.moveTo(1227, 684)      # Move o cursor do mouse para a posição especificada    
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c') # Simula pressionar simultaneamente as teclas 'ctrl' e 'c', copiando para a área de transferência

    valor6 = str(root.clipboard_get())  # Obtém o valor da área de transferência e o converte para string

    # Encontra o texto 'de' no valor copiado e converte os caracteres seguintes em um número inteiro
    dealers_em_aberto = int(str(valor6[valor6.find('de'):]).replace('de ', '')) 

    if dealers_em_aberto <= 10:
        print('\033[32mDealers em aberto:', dealers_em_aberto, '\033[m') # Printa Dealers na cor verde no terminal
    elif dealers_em_aberto > 10:
        print('\033[31mDealers em aberto:', dealers_em_aberto, '\033[m') # Printa Dealers na cor vermelha no terminal
    else:    
        0

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')

##################################################################################################################################
################################################## CORREÇÃO DPC ##################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_PROV").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

clipboard.copy('0')

if not pyautogui.moveTo(1227, 684):  
    pyautogui.moveTo(1227, 684)    
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c') 

valor7 = str(root.clipboard_get())

correcao_dpc_em_aberto = int(str(valor7[valor7.find('de'):]).replace('de ', ''))

if correcao_dpc_em_aberto <= 10:
    print('\033[32mCorreção DPC em aberto:', correcao_dpc_em_aberto, '\033[m')
elif correcao_dpc_em_aberto > 10: 
    print('\033[31mCorreção DPC em aberto:', correcao_dpc_em_aberto, '\033[m')  
else:
    0

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')  
    
##################################################################################################################################
########################################################## DPC ###################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "dpc"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").setFocus
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 3
session.findById("wnd[1]").sendVKey (0)

clipboard.copy('0') 

if not pyautogui.moveTo(1227, 684): 
    pyautogui.moveTo(1227, 684)  
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor8 = str(root.clipboard_get())

    dpc_em_aberto = int(str(valor8[valor8.find('de'):]).replace('de ', ''))

    if dpc_em_aberto >=1:
        print('DPC em aberto:', dpc_em_aberto) 
    else:
        0    

# Tentativa de fechar o pop-up com a mensagem "No matches with this selection"
pyautogui.press('esc')
time.sleep(1)  # Espere 1 segundo para garantir que o pop-up tenha sido fechado
pyautogui.press('esc')  

##################################################################################################################################
###################################################### SOMA FILA CIC #############################################################
##################################################################################################################################

fila_cic = pedidos_em_aberto + cotacoes_em_aberto + outros_em_aberto + correcao_dpc_em_aberto

if fila_cic  <= 280:

    print('\033[32mFila CIC:', fila_cic, '\033[m') # 32 é o código da cor verde

else:
    print('\033[31mFila CIC:', fila_cic, '\033[m') # 31 é o código da cor vermelha

time.sleep(1)

##################################################################################################################################
######################################################### POS VENDA ##############################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()

session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2221"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

###################################
### VERIFICA SE HÁ ERRO NA TELA ###
###################################    

if not pyautogui.locateCenterOnScreen(fr'C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\no_matches.png'):

# Se a imagem não for encontrada, o bloco abaixo é executado

    if not pyautogui.locateCenterOnScreen(r"C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\no_matches.png"):
    
        pyautogui.moveTo(1227, 684)
        pyautogui.tripleClick(1227, 684)
        pyautogui.hotkey('ctrl', 'c')

        valor9 = str(root.clipboard_get())

        pos_venda_em_aberto = int(str(valor9[valor9.find('de'):]).replace('de ', ''))
       

        if  pos_venda_em_aberto <= 45:
            print('\033[32mPós-venda em aberto:',  pos_venda_em_aberto, '\033[m')
        else:
            print('\033[31mPós-venda em aberto:',  pos_venda_em_aberto, '\033[m')

else:
# Se a imagem for encontrada, o bloco abaixo é executado
    pos_venda_em_aberto = 0
    print('\033[32mPós-venda em aberto:',  pos_venda_em_aberto, '\033[m')

##################################################################################################################################
######################################################### ANÁLISE ################################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2203"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

###################################
### VERIFICA SE HÁ ERRO NA TELA ###
###################################    

if not pyautogui.locateCenterOnScreen(fr'C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\no_matches.png'):

# Se a imagem não for encontrada, o bloco abaixo é executado

    if not pyautogui.locateCenterOnScreen(r"C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\no_matches.png"):
        
        pyautogui.moveTo(1227, 684)
        pyautogui.tripleClick(1227, 684)
        pyautogui.hotkey('ctrl', 'c')

        valor10 = str(root.clipboard_get())

        analise_em_aberto = int(str(valor10[valor10.find('de'):]).replace('de ', ''))

        print('Análise em aberto:',  analise_em_aberto)
        
    else:
# Se a imagem for encontrada, o bloco abaixo é executado
        analise_em_aberto = 0
        print('\033[32mAnálise em aberto:',  analise_em_aberto, '\033[m')

##################################################################################################################################
###################################################### REFATURAMENTOS ############################################################
##################################################################################################################################

session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
session.findById("wnd[1]/tbar[0]/btn[7]").press()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2226"
session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press()

clipboard.copy('0')

try:
    pyautogui.moveTo(1227, 684)
    pyautogui.tripleClick(1227, 684)
    pyautogui.hotkey('ctrl', 'c')

    valor11 = str(root.clipboard_get())

    refaturamento_em_aberto = int(str(valor11[valor11.find('de'):]).replace('de ', ''))

    print('Refaturamento em aberto:',  refaturamento_em_aberto)
        
except:
    refaturamento_em_aberto = 0
    print('\033[32mRefaturamento em aberto:',  refaturamento_em_aberto, '\033[m')
    
##################################################################################################################################
################################################# FINAL DE PUXAR A FILA  #########################################################
#################################################      DADOS VIA SAP     #########################################################
##################################################################################################################################

##################################################################################################################################
################################################### E-MAILS NÃO LIDOS ############################################################

username = getpass.getuser()

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

##################################################### MARKET PLACE ################################################################

#root_folder = Diretório Principal
root_folder = outlook.folders['BR Festo MBX Marketplace']

# Imprimir o root_folder em amarelo
print(Fore.LIGHTYELLOW_EX + str(root_folder) + Style.RESET_ALL)

entrada = root_folder.Folders.Item("Caixa de Entrada")

sub = entrada.Folders.Item("2024")

##################################################################################################################################
##################################################### PEDIDOS PORTAL #############################################################
##################################################################################################################################

gui = sub.Folders.Item("Pedidos - Guilherme")
dai = sub.Folders.Item("Pedidos - Daiana")
let = sub.Folders.Item("Pedidos - Letícia")

####################################
######## PEDIDOS GUILHERME #########
####################################

# Lista para armazenar os pedidos não lidos
lista_pedidos_guilherme = [] 

# Iterar sobre as subpastas de "Pedidos - Guilherme"
for sub_folder in gui.Folders:
    # Contar e-mails não lidos em cada subpasta
    unread_count = sub_folder.UnReadItemCount
    # Adicionar a contagem à lista
    lista_pedidos_guilherme.append(unread_count)

# Somar o total de pedidos não lidos
pedidos_em_aberto_guilherme = sum(lista_pedidos_guilherme)
print('Pedidos Guilherme em aberto:', pedidos_em_aberto_guilherme)

####################################
########## PEDIDOS DAIANA ##########
####################################

# Lista para armazenar os pedidos não lidos
lista_pedidos_daiana = [] 

# Iterar sobre as subpastas de "Pedidos - Daiana"
for sub_folder in dai.Folders:
    # Contar e-mails não lidos em cada subpasta
    unread_count = sub_folder.UnReadItemCount
    # Adicionar a contagem à lista
    lista_pedidos_daiana.append(unread_count)

# Somar o total de pedidos não lidos
pedidos_em_aberto_daiana = sum(lista_pedidos_daiana)
print('Pedidos Daiana em aberto:', pedidos_em_aberto_daiana)

####################################
########## PEDIDOS LETÍCIA #########
####################################

# Lista para armazenar os pedidos não lidos
lista_pedidos_leticia = [] 

# Iterar sobre as subpastas de "Pedidos - Letícia"
for sub_folder in let.Folders:
    # Contar e-mails não lidos em cada subpasta
    unread_count = sub_folder.UnReadItemCount
    # Adicionar a contagem à lista
    lista_pedidos_leticia.append(unread_count)

# Somar o total de pedidos não lidos
pedidos_em_aberto_leticia = sum(lista_pedidos_leticia)
print('Pedidos Letícia em aberto:', pedidos_em_aberto_leticia)

#####################################
## SOMAR PEDIDOS ABERTOS NO PORTAL ##
#####################################

pedidos_em_aberto_portal = (pedidos_em_aberto_guilherme + pedidos_em_aberto_daiana + pedidos_em_aberto_leticia)

if pedidos_em_aberto_portal <= 80:
    print('\033[32mPedidos em aberto no Portal:', pedidos_em_aberto_portal, '\033[m') # 32mPedidos é o código da cor verde

else:
    print('\033[31mPedidos em aberto no Portal:', pedidos_em_aberto_portal, '\033[m') # 31mPedidos é o código da cor vermelha

##################################################################################################################################
##################################################### PORTAL CADASTRO ############################################################
##################################################################################################################################

cad = sub.Folders.Item("PORTAL - CADASTRO")

cadastro_em_aberto_portal = cad.UnReadItemCount
print('\033[32mCadastros em aberto no Portal:', cadastro_em_aberto_portal, '\033[m')

##################################################################################################################################
##################################################### COTAÇÕES PORTAL ############################################################
##################################################################################################################################

# Obtém a lista de e-mails não lidos dentro da subpasta "COTAÇÕES"
cot1portal = sub.Folders.Item("COTAÇÕES").UnReadItemCount
print('COTAÇÕES EM ABERTO:', cot1portal)

cot2portal = sub.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Abner").UnReadItemCount
print('Cotações - Abner:', cot2portal)

cotacoes_abner_folder = sub.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Abner")
retorno_a_folder = cotacoes_abner_folder.Folders.Item("RETORNO - A")
cot3portal = retorno_a_folder.UnReadItemCount
print('Retorno - A:', cot3portal)

cot4portal = sub.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Thainara").UnReadItemCount
print('Cotações - Thainara:', cot4portal)

cotacoes_thainara_folder = sub.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Thainara")
retorno_t_folder = cotacoes_thainara_folder.Folders.Item("RETORNO - T")
cot5portal = retorno_t_folder.UnReadItemCount
print('Retorno - T:', cot5portal)

cotacoes_em_aberto_portal=cot1portal+cot2portal+cot3portal+cot4portal+cot5portal

if cotacoes_em_aberto_portal <= 30:
    print('\033[32mCotações em aberto no Portal:', cotacoes_em_aberto_portal, '\033[m') # 32m é o código da cor verde

else:
    print('\033[31mCotações em aberto no Portal:', cotacoes_em_aberto_portal, '\033[m') # 31m é o código da cor vermelha

##################################################################################################################################
##################################################################################################################################
######################################################### CARTEIRA ###############################################################
    
root_folder = outlook.folders['BR Festo MBX Carteira']

print(Fore.LIGHTYELLOW_EX + str(root_folder) + Style.RESET_ALL)

entrada = root_folder.Folders.Item("Caixa de Entrada")

carteira = entrada.UnReadItemCount
carteira += entrada.Folders.Item("0.1 - DIVERSOS").UnReadItemCount
carteira += entrada.Folders.Item("1 - Assuntos Logísticos").UnReadItemCount
carteira += entrada.Folders.Item("2 - Antecipações").UnReadItemCount
carteira += entrada.Folders.Item("3 - Cancelamento").UnReadItemCount
carteira += entrada.Folders.Item("4 - Bloqueio de período").UnReadItemCount
carteira += entrada.Folders.Item("5 - FOLLOW UP").UnReadItemCount
carteira += entrada.Folders.Item("7-Top Priority").UnReadItemCount

emails_nao_lidos_carteira = carteira

# Na primeira linha estamos inicializando a variável cart com o número de mensagens não lidas na pasta "Caixa de Entrada".
# Nas demais linhas estamos adicionando o número de mensagens não lidas nas demais pastas como a pasta "0.1 - DIVERSOS" 
# ao valor já existente na varável cart.

if emails_nao_lidos_carteira <= 40:
    print('\033[32mE-mails não lidos Carteira:', emails_nao_lidos_carteira, '\033[m') # 32m é o código da cor verde

else:
    print('\033[31mE-mails não lidos Carteira:', emails_nao_lidos_carteira, '\033[m') # 31m é o código da cor vermelha

##################################################################################################################################
##################################################################################################################################
######################################################### CADASTRO ###############################################################
    
root_folder = outlook.folders['BR Festo MBX Cadastro']

print(Fore.LIGHTYELLOW_EX + str(root_folder) + Style.RESET_ALL)

entrada = root_folder.Folders.Item("Caixa de Entrada")

cadastro = entrada.UnReadItemCount
cadastro+= entrada.Folders.Item("ATUALIZAÇÃO CADASTRO").UnReadItemCount
cadastro+= entrada.Folders.Item("BLOQUEIO DE PERIODO").UnReadItemCount
cadastro+= entrada.Folders.Item("CAD. DISTRIBUIDOR").UnReadItemCount
cadastro+= entrada.Folders.Item("CAD. NOVO - FORMS").UnReadItemCount
cadastro+= entrada.Folders.Item("CAD. NOVO - FOX").UnReadItemCount
cadastro+= entrada.Folders.Item("CAD. NOVO - OUTROS").UnReadItemCount
cadastro+= entrada.Folders.Item("CERTIF. QUALIDADE").UnReadItemCount
cadastro+= entrada.Folders.Item("CONV I220 P/ K220").UnReadItemCount
#cadastro+= entrada.Folders.Item("DE/PARA").UnReadItemCount
cadastro+= entrada.Folders.Item("EXTENSÃO 20 - DIDACTIC").UnReadItemCount
cadastro+= entrada.Folders.Item("FERIAS COLETIVAS").UnReadItemCount
cadastro+= entrada.Folders.Item("FERIAS CONSUTOR").UnReadItemCount
cadastro+= entrada.Folders.Item("OUTROS ASSUNTOS").UnReadItemCount
cadastro+= entrada.Folders.Item("POP UP").UnReadItemCount
cadastro+= entrada.Folders.Item("PORTAL").UnReadItemCount
cadastro+= entrada.Folders.Item("PORTAL TRANSPARENCIA").UnReadItemCount
cadastro+= entrada.Folders.Item("TRANSF. CARTEIRA").UnReadItemCount
cadastro+= entrada.Folders.Item("TRANSPORTADORA").UnReadItemCount
cadastro+= entrada.Folders.Item("XML").UnReadItemCount
   
emails_nao_lidos_cadastro = cadastro

if emails_nao_lidos_cadastro <= 25:
    print('\033[32mE-mails não lidos Cadastro', emails_nao_lidos_cadastro, '\033[m') # 32m é o código da cor verde

else:
    print('\033[31mE-mails não lidos Cadastro:', emails_nao_lidos_cadastro, '\033[m') # 31m é o código da cor vermelha

##################################################################################################################################
##################################################################################################################################
#################################################### FESTO - ARGENTINA ###########################################################

root_folder = outlook.folders['AR Festo MBX Ventas MDP']

print(Fore.LIGHTYELLOW_EX + str(root_folder) + Style.RESET_ALL)
   
entrada = root_folder.Folders.Item("Bandeja de entrada")

festo_ar = entrada.UnReadItemCount
    
print("E-mails não lidos Festo - Argentina:", festo_ar)     

##################################################################################################################################
#################################################### FINAL LER E-MAILS ###########################################################
##################################################################################################################################

##################################################################################################################################
##################################################### APP SALES TOOL #############################################################
##################################################################################################################################

# Configurar o Edge Driver
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
edge_options.add_argument(r"--profile-directory=Profile 1")
driver = Edge(options=edge_options, executable_path=fr"C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\MATHEUSmsedgedriver.exe")


# Abrir a página desejada
driver.get("https://festo-my.sharepoint.com/personal/br0gfsj_festo_net/Lists/Sales%20Tool%20Solicitao%20rpida%20Contact%20Center/AllItems.aspx?FilterField1=Status&FilterValue1=Em%20aberto&FilterType1=Choice")

try:
    # Esperar até que a palavra 'Count' esteja presente ou atingir o tempo limite
    count_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Contagem')]")))
    
    time.sleep(3)
    
    # Duplo clique no elemento
    ActionChains(driver).double_click(count_element).perform()

    # Simular o Ctrl+C para copiar o texto atualizado para a área de transferência
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()

    # Obter o texto atualizado do clipboard usando pyperclip
    texto_atualizado = pyperclip.paste()
    print("Texto Atualizado:", texto_atualizado)

    # Ajustar a expressão regular para corresponder ao padrão esperado
    re_match = re.search(r"(\d+)", texto_atualizado)

    # Extrair o resultado diretamente sem verificar se houve correspondência
    salestool = int(re_match.group(1)) 
    print("Sales Tool em aberto:", salestool)


except TimeoutException:
    salestool = 0
    print("Sales Tool em aberto:", salestool)


finally:
    # Fechar o navegador
    driver.quit()


##################################################################################################################################
####################################################### APP CARTEIRA #############################################################
##################################################################################################################################

# Configurar o Edge Driver
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
edge_options.add_argument(r"--profile-directory=Profile 1")
driver = Edge(options=edge_options, executable_path=fr"C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\MATHEUSmsedgedriver.exe")


# Abrir a página desejada
driver.get("https://festo-my.sharepoint.com/personal/br0gfsj_festo_net/Lists/Sales%20tool%20Solicitao%20rpida%20Carteira/AllItems.aspx?xsdata=MDV8MDJ8fGMxOGI0ODgyMzMyMjQwOTRiYTE0MDhkYzY4NTBlYzNlfGExYWU4OWZiMjFiOTQwYmY5ZDgyYTEwYWU4NWEyNDA3fDB8MHw2Mzg0OTk5NDM5NDk5NzQwMzZ8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPakE0TWpFNU5qTTJMVGd3TkdNdE5HRmpZaTA1TUdSbUxXVXhNR05rWWpBeU9EY3haRjgzWWpjeE56aG1ZeTAzTWpNekxUUmpObVl0T0RKaE9TMHdaV1V3WWpJME9UQmpaVFZBZFc1eExtZGliQzV6Y0dGalpYTXZiV1Z6YzJGblpYTXZNVGN4TkRNNU56VTVOREEwTkE9PXwzNzg0MGM4ODgwYTc0N2IxYmExNDA4ZGM2ODUwZWMzZXxlM2FhOWM1ZDY3NmY0NzdjYTcwNjMxMjY3MzRmYjM4Ng%3D%3D&sdata=bW4wdjJnc3hTZTJpeUpUNVI5bTdGcGtJOTFzWWZXNXFNWGRObmVtdHF5UT0%3D&ovuser=a1ae89fb-21b9-40bf-9d82-a10ae85a2407%2Cbr0mslvr%40festo.net&OR=Teams-HL&CT=1714399698900&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yNDAzMjgyMTIwMCIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&FilterField1=Status&FilterValue1=Em%20aberto&FilterType1=Choice")

try:
    # Esperar até que a palavra 'Count' esteja presente ou atingir o tempo limite
    count_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Contagem')]")))
    
    time.sleep(3)
    
    # Duplo clique no elemento
    ActionChains(driver).double_click(count_element).perform()

    # Simular o Ctrl+C para copiar o texto atualizado para a área de transferência
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()

    # Obter o texto atualizado do clipboard usando pyperclip
    texto_atualizado = pyperclip.paste()
    print("Texto Atualizado:", texto_atualizado)

    # Ajustar a expressão regular para corresponder ao padrão esperado
    re_match = re.search(r"(\d+)", texto_atualizado)

    # Extrair o resultado diretamente sem verificar se houve correspondência
    app_carteira = int(re_match.group(1)) 
    print("APP CARTEIRA em aberto:", app_carteira)


except TimeoutException:
    app_carteira = 0
    print("APP CARTEIRA em aberto:", app_carteira)


finally:
    # Fechar o navegador
    driver.quit()

##################################################################################################################################
###################################################### APP POS VENDA #############################################################
##################################################################################################################################

# Configurar o Edge Driver
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
edge_options.add_argument(r"--profile-directory=Profile 1")
driver = Edge(options=edge_options, executable_path=fr"C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\MATHEUSmsedgedriver.exe")


# Abrir a página desejada
driver.get("https://festo-my.sharepoint.com/personal/br0gfsj_festo_net/Lists/Ps%20vendas%20%20Sales%20Tool/AllItems.aspx?xsdata=MDV8MDJ8fGMxOGI0ODgyMzMyMjQwOTRiYTE0MDhkYzY4NTBlYzNlfGExYWU4OWZiMjFiOTQwYmY5ZDgyYTEwYWU4NWEyNDA3fDB8MHw2Mzg0OTk5NDM5NDk5NzQwMzZ8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPakE0TWpFNU5qTTJMVGd3TkdNdE5HRmpZaTA1TUdSbUxXVXhNR05rWWpBeU9EY3haRjgzWWpjeE56aG1ZeTAzTWpNekxUUmpObVl0T0RKaE9TMHdaV1V3WWpJME9UQmpaVFZBZFc1eExtZGliQzV6Y0dGalpYTXZiV1Z6YzJGblpYTXZNVGN4TkRNNU56VTVOREEwTkE9PXwzNzg0MGM4ODgwYTc0N2IxYmExNDA4ZGM2ODUwZWMzZXxlM2FhOWM1ZDY3NmY0NzdjYTcwNjMxMjY3MzRmYjM4Ng%3D%3D&sdata=eHlIK2VtazRPRVB0QmZOWGlTbjJ0SWk1U1J0OUd4UktadUtPMTJWTm9zMD0%3D&ovuser=a1ae89fb-21b9-40bf-9d82-a10ae85a2407%2Cbr0mslvr%40festo.net&OR=Teams-HL&CT=1714399697044&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yNDAzMjgyMTIwMCIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&FilterField1=Status&FilterValue1=Em%20aberto&FilterType1=Choice")

try:
    # Esperar até que a palavra 'Count' esteja presente ou atingir o tempo limite
    count_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Contagem')]")))
    
    time.sleep(3)
    
    # Duplo clique no elemento
    ActionChains(driver).double_click(count_element).perform()

    # Simular o Ctrl+C para copiar o texto atualizado para a área de transferência
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()

    # Obter o texto atualizado do clipboard usando pyperclip
    texto_atualizado = pyperclip.paste()
    print("Texto Atualizado:", texto_atualizado)

    # Ajustar a expressão regular para corresponder ao padrão esperado
    re_match = re.search(r"(\d+)", texto_atualizado)

    # Extrair o resultado diretamente sem verificar se houve correspondência
    app_posvenda = int(re_match.group(1)) 
    print("APP Pos-Venda em aberto:", app_posvenda)


except TimeoutException:
    app_posvenda = 0
    print("APP Pos-Venda em aberto:", app_posvenda)


finally:
    # Fechar o navegador
    driver.quit()    

##################################################################################################################################
################################################# INDICADORES CRM - TELEFONE #####################################################
##################################################################################################################################

 # EXECUTA O PROGRAMA "Edge DRIVER"

edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
edge_options.add_argument(r"--profile-directory=Profile 1")
driver = Edge(options = edge_options, executable_path = fr"C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\MATHEUSmsedgedriver.exe")
   
driver.get("http://adevrt01.de.festo.net/DynamicView/?tabid=null&lid=8db71bbf-4fa4-e811-8ebb-005056a1f968&sid=f3311717-4849-4895-b976-fa6e79347364&tid=3189a3d6-8d59-4855-9a29-3853510c2bb1")

wait = WebDriverWait(driver, 3000)

wait = WebDriverWait(driver, 10)

try:
    report1 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div/div/div[2]/input')))
    report1 = driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div[2]/input')

    report1.clear()
    report1.send_keys("BR0CCES")

    report1 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div/div/div[3]/input')))
    report1 = driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div[3]/input')

    report1.clear()
    report1.send_keys("1234")

    report1 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="react-select-2-input"]')))
    report1 = driver.find_element_by_xpath('//*[@id="react-select-2-input"]')

    report1.send_keys("CC") 
    report1.send_keys(u'\ue007')

    report1 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div/div/div[5]/span/span')))
    report1 = driver.find_element_by_xpath('//*[@id="app"]/div/div/div/div[5]/span/span')

    report1.click()
    
except:
    pass

time.sleep(7)

    # Supondo que você já esteja na próxima página
try: 
    campo_selecao  = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[1]/div[1]/div/div/div[2]/div[1]/div/div/div[1]')))

    time.sleep(7)

    campo_selecao.click()

    time.sleep(2)

    pyautogui.press('enter')

    driver.find_element_by_xpath('//*[@id="app"]/div/div[1]/div[1]/div/div/div[2]/div[3]/span[1]').click()

except:
    print('Login já realizado')
    pass

time.sleep(10)

report1 = wait.until(EC.element_to_be_clickable((By.XPATH, r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[3]/div[10]')))

time.sleep(10)

recebidas = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[9]').text
recebidascare = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[1]/div[9]').text
recebidascomercial = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[2]/div[9]').text
ate20 = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[12]').text
nivelcare = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[1]/div[12]').text
nivelcomer = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[2]/div[12]').text
nivelserv = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[12]').text
nivelchat = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[3]/div[12]').text
volumechat = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[3]/div[9]').text
    
nivelcare = str(nivelcare)+'%'
nivelcomer = str(nivelcomer)+'%'
nivelserv = str(nivelserv)+'%'
nivelchat = str(nivelchat)+'%'

if recebidas == '-':
    recebidas=0
    recebidas = int(recebidas)

if recebidascare == '-':
    recebidascare=0
    recebidascare = int(recebidascare)

if recebidascomercial == '-':
    recebidascomercial=0
    recebidascomercial = int(recebidascomercial)

print('TELEFONE')
print('Ligações Recebidas:', recebidas) 
print('Ligações Recebidas Care:', recebidascare) 
print('Ligações Recebidas Comercial:', recebidascomercial) 
print('Ligações  atendidas até 20s:', ate20) 
print('Nível Care:', nivelcare) 
print('Nível Comercial:', nivelcomer) 
print('Nível de Serviço:', nivelserv) 
print('Nível Chat:', nivelchat) 
print('Volume Chat:', volumechat) 

time.sleep(2)
driver.quit()

##################################################################################################################################
################################################## CRIAÇÃO DOS ARQUIVOS CSV ######################################################
##################################################################################################################################

# Caminho do arquivo CSV desejado
caminho_arquivo_csv1 = r'C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\Numeros Atuais-Matheus.csv'
caminho_arquivo_csv2 = r'C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\Historico-Matheus.csv'

# Nome do arquivo CSV
nome_arquivo_csv1 = 'Numeros Atuais-Matheus.csv'
nome_arquivo_csv2 = 'Historico-Matheus.csv'

# Verifica se o arquivo CSV1 existe
if not os.path.isfile(nome_arquivo_csv1):
    # Se não existir, cria o arquivo 
    with open(caminho_arquivo_csv1, 'w', newline='') as arquivo_csv1:
        escritor_csv1 = csv.writer(arquivo_csv1)

# Verifica se o arquivo CSV2 existe
if not os.path.isfile(nome_arquivo_csv2):
    # Se não existir, cria o arquivo 
    with open(caminho_arquivo_csv2, 'w', newline='') as arquivo_csv2:
        escritor_csv2 = csv.writer(arquivo_csv2)
        
# Obtém a hora atual
hora_atual = datetime.datetime.now().strftime('%H:%M')

# Dados a serem escritos
dados = [hora_atual, entrada_de_documentos, pedidos_em_aberto, cotacoes_em_aberto, outros_em_aberto, cs_em_aberto, dealers_em_aberto, correcao_dpc_em_aberto, dpc_em_aberto, fila_cic, pos_venda_em_aberto, analise_em_aberto, refaturamento_em_aberto, pedidos_em_aberto_portal, cadastro_em_aberto_portal, cotacoes_em_aberto_portal, emails_nao_lidos_carteira, emails_nao_lidos_cadastro, festo_ar, salestool, recebidas, recebidascare, recebidascomercial, ate20, nivelcare, nivelcomer, nivelserv, nivelchat, volumechat, app_carteira, app_posvenda]

# Escreve no arquivo Numeros Atuais-Matheus.csv
with open(caminho_arquivo_csv1, 'w', newline='') as arquivo_csv1:
    escritor_csv1 = csv.writer(arquivo_csv1)
    escritor_csv1.writerow(['Horas', 'ENTRADA DE DOCUMENTOS', 'Pedidos', 'Cotacao', 'Outros', 'CS', 'Dealers', 'CORRECAO DPC', 'DPC', 'FILA CIC', 'POS VENDA', 'ANALISE', 'REFA', 'PEDIDOS PORTAL', 'CADASTRO PORTAL', 'COTACOES PORTAL', 'CARTEIRA', 'CADASTRO', 'FESTO AR', 'SALESTOOL', 'LIGACOES RECEBIDAS', 'LIGACOES RECEBIDAS CARE', 'LIGACOES RECEBIDAS COMERICAL', 'LIGACOES ATENDIDAS ATE 20S', 'NIVEL CARE', 'NIVEL COMERCIAL', 'NIVEL DE SERVICO', 'NIVEL CHAT', 'VOLUME CHAT', 'App Carteira', 'App Pos-Venda'])
    escritor_csv1.writerow(dados)

# Escreve no arquivo Historico-Matheus.csv
with open(caminho_arquivo_csv2, 'a', newline='') as arquivo_csv2:
    escritor_csv2 = csv.writer(arquivo_csv2)
    escritor_csv2.writerow(dados)

# Verifica o número de linhas no arquivo Historico-Matheus.csv
num_linhas = len(pd.read_csv(caminho_arquivo_csv2))
if num_linhas > 11: # O CSV2 fica com 13 linhas pq a primeira linha é a linha 0. Então com 13 linhas, na verdade são 12 linhas.
    with open(caminho_arquivo_csv2, 'w', newline='') as arquivo_csv2:
        escritor_csv2 = csv.writer(arquivo_csv2)
        escritor_csv2.writerow(['Horas', 'ENTRADA DE DOCUMENTOS', 'Pedidos', 'Cotacao', 'Outros', 'CS', 'Dealers', 'CORRECAO DPC', 'DPC', 'FILA CIC', 'POS VENDA', 'ANALISE', 'REFA', 'PEDIDOS PORTAL', 'CADASTRO PORTAL', 'COTACOES PORTAL', 'CARTEIRA', 'CADASTRO', 'FESTO AR', 'SALESTOOL', 'LIGACOES RECEBIDAS', 'LIGACOES RECEBIDAS CARE', 'LIGACOES RECEBIDAS COMERICAL', 'LIGACOES ATENDIDAS ATE 20S', 'NIVEL CARE', 'NIVEL COMERCIAL', 'NIVEL DE SERVICO', 'NIVEL CHAT', 'VOLUME CHAT', 'App Carteira', 'App Pos-Venda'])
        escritor_csv2.writerow(dados)

##################################################################################################################################
########################################################### THE END ##############################################################
##################################################################################################################################