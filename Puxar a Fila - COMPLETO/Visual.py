from selenium import webdriver
import time
import pyautogui
import pyperclip
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options 
import threading
import pymsgbox
import pandas as pd
import xlrd
from typing import Counter
import win32com.client as win32
import openpyxl
import unittest
from msedge.selenium_tools import Edge, EdgeOptions
from openpyxl import load_workbook, cell
import csv
from datetime import date, timedelta
from selenium.webdriver.support import expected_conditions as EC, select
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
import re 
import getpass



username=getpass.getuser()

list_cad_email = [
    'ATUALIZACAO CADASTRO',
    'BLOQUEIO PERIODO',
    'CAD. DISTRIBUIDOR',
    'CAD. NOVO - FORMS',
    'CAD. NOVO - FOX',
    'CAD. NOVO-OUTROS',
    'CONV I220 P/ K220',
    'DE PARA',
    'FÉRIAS COLETIVAS',
    'FERIAS CONSULTOR',
    'PAGAMENTO PORTAL',
    'POP UP PENDENTE',
    'TRANSF. CARTEIRA',
    'TRANSPORTADORA',
    'XML'
]

global horaagora
global root_folder
numcotport=0
numcad=0

yesterday = date.today()
hoje = str(yesterday.strftime('%d.%m.%Y'))
#-----------------------------------------------------------------------------
pymsgbox.alert('Abra o SAP na tela principal', 'Aviso')
#-----------------------------------------------------------------------------

SapGuiAuto = win32.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
global session
session = connection.Children(0)
session.findById("wnd[0]").maximize()

def robosaphora():
    global cotacao
    global pedido
    global outros
    global portal
    global customersolutions
    global posvenda
    global dpc
    global verificacaodpc
    global ccfila
    global entradadedocfax
    global dealers
    global analise
    verificacaodpc = 0
        
    for numfila in range(12):
        session.findById("wnd[0]/usr/subSC_SUB_3:Y02VSI_FAX_MONITOR:0300/tabsSC3_TABSTRIP/tabpSC3_TAB_6/ssubSC3_SUB1:Y02VSI_FAX_MONITOR:0310/cntlSC_HTML/shellcont/shell").sapEvent("","","sapevent:S_?SUBM=Busca")
        session.findById("wnd[1]/tbar[0]/btn[7]").press()
        session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_OPEN").select()
        
        if numfila ==0:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_ANG").select()

        if numfila ==1:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_BES").select()

        if numfila ==2:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_SON").select()

        if numfila ==3:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2226"
            #DELETAR

        if numfila ==4:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2224"

        if numfila ==5:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2221"

        if numfila ==6:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "DPC"

        if numfila ==7:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2227"

        if numfila ==8:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/radP_PROV").select()
    
        if numfila ==9:
            session.findById("wnd[1]/tbar[0]/btn[7]").press()
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2201"
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/ctxtO_DATUM-LOW").text = hoje

        if numfila ==10:
            session.findById("wnd[1]/usr/subSUB:SAPLY02VSI_FAX_MONITOR:0295/txtO_VKB_I-LOW").text = "2203"
        session.findById("wnd[1]").sendVKey(0)

        try:
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            if numfila ==0:
                cotacao = 0

            if numfila ==1:
                pedido = 0

            if numfila ==2:
                outros = 0

            if numfila ==3:
                portal = 0

            if numfila ==4:
                customersolutions = 0

            if numfila ==5:
                posvenda = 0
                
            if numfila ==6:
                dpc = 0

            if numfila ==7:
                dealers = 0              
            
            if numfila ==8:
                verificacaodpc = 0
           
            if numfila ==9:
                entradadedocfax = 0

            if numfila ==10:
                analise = 0  

        except:
            imgok=1
            while imgok:
                if pyautogui.locateCenterOnScreen(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\deimg3.png'):
                    imgok=0
                    pyautogui.moveTo(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\deimg3.png')
                    x, y = pyautogui.position()
                    x = x + 30
                    y = y + 30
                    pyautogui.doubleClick(x, y)
                    pyautogui.hotkey('ctrl','c')
                    time.sleep(2)

                    if pyperclip.paste() =="":
                        pyautogui.moveTo(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\deimg3.png')
                        x, y = pyautogui.position()
                        y = y + 20
                        pyautogui.doubleClick(x, y)
                        pyautogui.hotkey('ctrl','c')
            
                    if numfila ==0:
                        cotacao = pyperclip.paste()
                        print(cotacao)
                    if numfila ==1:
                        pedido = pyperclip.paste()
                        print(pedido)
                    if numfila ==2:
                        outros = pyperclip.paste()
                        print(outros)
                    if numfila ==3:
                        portal = pyperclip.paste()
                        print(portal)
                    if numfila ==4:
                        customersolutions = pyperclip.paste()
                        print(customersolutions)
                    if numfila ==5:
                        posvenda = pyperclip.paste()
                        print(posvenda)
                        
                    if numfila ==6:
                        dpc = pyperclip.paste()
                        print(dpc)

                    if numfila ==7:
                        dealers = pyperclip.paste()
                        print(dealers)

                    if numfila ==8:
                        verificacaodpc = pyperclip.paste()
                        print(verificacaodpc)

                    if numfila ==9:
                        entradadedocfax = pyperclip.paste()
                        print(entradadedocfax)

                    if numfila ==10:
                        analise = pyperclip.paste()
                        print(analise)

    ccfila = int(cotacao)+int(pedido)+int(outros)+int(verificacaodpc)

def robotelgora():
    global recebidas
    global ate20
    global logadocare
    global logadocomer
    global nivelcare
    global nivelcomer
    global nivelserv

    #-----------------------------------------------------------------------------
    #EXECUTA O PROGRAMA "Edge DRIVER"

    edge_options = EdgeOptions()
    edge_options.use_chromium = True
    edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
    edge_options.add_argument(r"--profile-directory=Profile 1")
    driver = Edge(options = edge_options, executable_path = fr"C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\msedgedriver.exe")
   
    #-----------------------------------------------------------------------------

    driver.get("http://adevrt01.de.festo.net/DynamicView/?tabid=null&lid=8db71bbf-4fa4-e811-8ebb-005056a1f968&sid=f3311717-4849-4895-b976-fa6e79347364&tid=3189a3d6-8d59-4855-9a29-3853510c2bb1")

    wait = WebDriverWait(driver, 3000)

    try:
        wait = WebDriverWait(driver, 10)
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

    report1 = wait.until(EC.element_to_be_clickable((By.XPATH, r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[3]/div[10]')))

    time.sleep(7)
    recebidas = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[9]').text
    ate20 = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[12]').text
    logadocare = 0
    logadocomer = 0
    nivelcare = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[1]/div[13]').text
    nivelcomer = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[2]/div[13]').text
    nivelserv = driver.find_element_by_xpath(r'//*[@id="RtServiceTable1_AGGrid"]/div/div/div[2]/div[2]/div[3]/div[2]/div/div/div[4]/div[13]').text
    
    nivelcare = str(nivelcare)+'%'
    nivelcomer = str(nivelcomer)+'%'
    nivelserv = str(nivelserv)+'%'
    if recebidas == '-':
        recebidas=0
    recebidas = int(recebidas)
    print("TELEFONE")
    print(recebidas)
    print(ate20)
    print(logadocare)
    print(logadocomer)
    print(nivelcare)
    print(nivelcomer)
    print(nivelserv)

    time.sleep(2)
    driver.quit()

def ddportal():
    
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    root_folder = outlook.folders['BR Festo MBX Marketplace']

    print(root_folder)
    entrada = root_folder.Folders.Item("Caixa de Entrada")
    sub1 = entrada.Folders.Item("2022")
    gui = sub1.Folders.Item("Pedidos - Guilherme")
    dai = sub1.Folders.Item("Pedidos - Daiana")
    let = sub1.Folders.Item("Pedidos - Letícia")
    cad = sub1.Folders.Item("PORTAL - CADASTRO").UnReadItemCount

    cot1portal = sub1.Folders.Item("COTAÇÕES").UnReadItemCount
    cot2portal = sub1.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Abner").UnReadItemCount
    cot3portal = sub1.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Thainara").UnReadItemCount
    global cotpo
    cotpo=cot1portal+cot2portal+cot3portal

    lis_gui = [] 

    for sub in gui.Folders:
        for mail in sub.Items:
            if mail.UnRead == True:
                lis_gui.append(mail.Subject)

    ped_gui = (len(lis_gui))
    #print (ped_gui)

    lis_dai = [] 

    for sub in dai.Folders:
        for mail in sub.Items:
            if mail.UnRead == True:
                lis_dai.append(mail.Subject)

    ped_dai = (len(lis_dai))
    #print (ped_dai)

    lis_let = [] 

    for sub in let.Folders:
        for mail in sub.Items:
            if mail.UnRead == True:
                lis_let.append(mail.Subject)

    ped_let = (len(lis_let))
    #print (ped_let)

    '''lis_cad = [] 

    for sub in cad.Folders:
        for mail in sub.Items:
            if mail.UnRead == True:
                lis_cad.append(mail.Subject)'''

    
    #print (ped_cad)
    global totalport
    totalport = (ped_gui + ped_dai + ped_let)
    print(totalport)
    global cad_mark
    cad_mark=cad
    #print (f'{totalport} Pedidos abertos no Portal')

def ddcarteira():

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    root_folder = outlook.folders['BR Festo MBX Carteira']

    print(root_folder)
    '''entrada = root_folder.Folders.Item("Caixa de Entrada").UnReadItemCount
    log1 = entrada.Folders.Item("1 - Assuntos Logísticos").UnReadItemCount
    antec2 = entrada.Folders.Item("2 - Antecipações").UnReadItemCount
    cancel = entrada.Folders.Item("3 - Cancelamento").UnReadItemCount
    bloq = entrada.Folders.Item("4 - Bloqueio de período").UnReadItemCount
    fup = entrada.Folders.Item("5 - FOLLOW UP").UnReadItemCount
    toppr = entrada.Folders.Item("7-Top Priority").UnReadItemCount'''
    entrada = root_folder.Folders.Item("Caixa de Entrada")
    len_cart= entrada.UnReadItemCount
    len_cart+= entrada.Folders.Item("1 - Assuntos Logísticos").UnReadItemCount
    len_cart+= entrada.Folders.Item("2 - Antecipações").UnReadItemCount
    len_cart+= entrada.Folders.Item("3 - Cancelamento").UnReadItemCount
    len_cart+= entrada.Folders.Item("4 - Bloqueio de período").UnReadItemCount
    len_cart+= entrada.Folders.Item("5 - FOLLOW UP").UnReadItemCount
    len_cart+= entrada.Folders.Item("7-Top Priority").UnReadItemCount

    global qntentcar
    '''lis_ent = [] 

    for sub in entrada.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    for sub in log1.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    for sub in antec2.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)
    
    for sub in cancel.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    for sub in bloq.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    for sub in fup.Items:
        
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    for sub in toppr.Items:
        if sub.UnRead == True:
            lis_ent.append(sub.Subject)

    qntentcar = (len(lis_ent))'''
    qntentcar=len_cart
    print(qntentcar) 


def cadecotp():
    global numcad
    global numcotport   
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    root_folder = outlook.folders['BR Festo MBX Cadastro']

    print(root_folder)
    entrada = root_folder.Folders.Item("Caixa de Entrada")
    count_unread = entrada.UnReadItemCount
    for i in list_cad_email:
        #print(entrada.Folders.Item(i+1).Name)
        count_unread =+entrada.Folders.Item(i).UnReadItemCount
    numcad=count_unread
    #sub1 = entrada.Folders.Item("CAD. NOVO - FORMS").UnReadItemCount
    

    #cot1portal = sub1.Folders.Item("COTAÇÕES").UnReadItemCount
    #cot2portal = sub1.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Abner").UnReadItemCount
    #cot3portal = sub1.Folders.Item("COTAÇÕES").Folders.Item("Cotações - Luanna").UnReadItemCount

    '''try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        for m in messages:
            assunto = m.Subject
            assunto = assunto.split(" ")
            if assunto[0] == "K7W02Q":
                #print('ok')
                print(assunto[1])
                numcad = assunto[1]
                m.Delete()

        for m in messages:
            assunto = m.Subject
            assunto = assunto.split(" ")
            if assunto[0] == "K7W01Q":
                #print('ok')
                print(assunto[1])
                numcotport = assunto[1]
                m.Delete()
    except:
        numcad = 0
        numcotport = 0 '''

def salvar():  
    
    try:
        data = pd.read_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Histórico.csv') 
        totallinhas=len(data)
        df = pd.read_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Números Atuais.csv')
        subligacoes = int(recebidas) - df['Ligacoes Recebidas'].iloc[0]
        subligacoes=int(subligacoes)
    except:
        totallinhas=20
        subligacoes = int(recebidas)

    if subligacoes <0:
        subligacoes = recebidas

    print("PRINT TOTAL LINHA")
    print(totallinhas)
    print('--------')
    print(subligacoes)
    if totallinhas >= 11:
        horaagora = 8
        horaagora = str(horaagora)+':00'
        df = pd.DataFrame([[horaagora,cotacao,pedido,outros,ccfila,portal,customersolutions,posvenda,dpc,verificacaodpc,subligacoes,ate20,logadocare,logadocomer,nivelcare,nivelcomer,nivelserv,entradadedocfax,totalport,cotpo,numcad,dealers,qntentcar,cad_mark,analise]])
        df.to_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Histórico.csv', index=False, mode='w', header=['Horas','Cotacao','Pedidos','Outros','CC Fila','FAX Portais','CS','Pos Venda','DPC','Verificacao DPC','Ligacoes Recebidas','Atendidas ate 20s','Logados - Care','Logados - Commercial','Nivel - Care','Nivel - Commercial','Nivel - CC','Entrada de Doc','Pedidos Portal','Cotacoes Portal','Cadastro','Dealers','Carteira','Cad. Portal','Analise 2203'])
        df2 = pd.DataFrame([[horaagora,cotacao,pedido,outros,ccfila,portal,customersolutions,posvenda,dpc,verificacaodpc,recebidas,ate20,logadocare,logadocomer,nivelcare,nivelcomer,nivelserv,entradadedocfax,totalport,cotpo,numcad,dealers,qntentcar,cad_mark,analise]])
        df2.to_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Números Atuais.csv', index=False, mode='w', header=['Horas','Cotacao','Pedidos','Outros','CC Fila','FAX Portais','CS','Pos Venda','DPC','Verificacao DPC','Ligacoes Recebidas','Atendidas ate 20s','Logados - Care','Logados - Commercial','Nivel - Care','Nivel - Commercial','Nivel - CC','Entrada de Doc','Pedidos Portal','Cotacoes Portal','Cadastro','Dealers','Carteira','Cad. Portal','Analise 2203'])
    else:
        horaagora = totallinhas + 8
        if horaagora == 18:
            horaagora ='17:30'
        else:
            horaagora = str(horaagora)+':00'
        df = pd.DataFrame([[horaagora,cotacao,pedido,outros,ccfila,portal,customersolutions,posvenda,dpc,verificacaodpc,subligacoes,ate20,logadocare,logadocomer,nivelcare,nivelcomer,nivelserv,entradadedocfax,totalport,cotpo,numcad,dealers,qntentcar,cad_mark,analise]])
        df.to_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Histórico.csv', index=False, mode='a', header=False)
        #'Cadastro'])
        df2 = pd.DataFrame([[horaagora,cotacao,pedido,outros,ccfila,portal,customersolutions,posvenda,dpc,verificacaodpc,recebidas,ate20,logadocare,logadocomer,nivelcare,nivelcomer,nivelserv,entradadedocfax,totalport,cotpo,numcad,dealers,qntentcar,cad_mark,analise]])
        df2.to_csv(fr'C:\Users\{username}\Festo\Customer Interaction Center - Fila Front\RPA\Números Atuais.csv', index=False, mode='w', header=['Horas','Cotacao','Pedidos','Outros','CC Fila','FAX Portais','CS','Pos Venda','DPC','Verificacao DPC','Ligacoes Recebidas','Atendidas ate 20s','Logados - Care','Logados - Commercial','Nivel - Care','Nivel - Commercial','Nivel - CC','Entrada de Doc','Pedidos Portal','Cotacoes Portal','Cadastro','Dealers','Carteira','Cad. Portal','Analise 2203'])


t2= threading.Thread(target=robotelgora)
t2.start()
robosaphora()
ddportal()
ddcarteira()
#cadecotp()
t2.join()
salvar()





