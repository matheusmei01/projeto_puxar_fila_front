import xmltojson, json, pandas
from lxml import html, etree
import pyautogui
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys  
from msedge.selenium_tools import Edge, EdgeOptions
import getpass
from selenium.common.exceptions import TimeoutException

username = getpass.getuser()

# Configurar o Edge Driver
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_argument(fr"--user-data-dir=C:\Users\{username}\AppData\Local\Microsoft\Edge\User Data\Default")
edge_options.add_argument(r"--profile-directory=Profile 1")
driver = Edge(options=edge_options, executable_path=fr"C:\Users\br0mslvr\Festo\CIC Management - Brasil - Fila Front\RPA\MATHEUSmsedgedriver.exe")

while True:  # Inicia um loop infinito

    driver.get("http://adevrt01.de.festo.net/DynamicView/?tabid=null&lid=8db71bbf-4fa4-e811-8ebb-005056a1f968&sid=f3311717-4849-4895-b976-fa6e79347364&tid=3189a3d6-8d59-4855-9a29-3853510c2bb1")
    # Acessa a URL especificada

    time.sleep(15)  # Espera 10 segundos para a página ser carregada

    def extract_name(name):  # Define uma função para extrair o nome do usuário. Pq na linha 61 o nome de usuário está vindo com o BR0xxxx, que é uma informação desnecessária.
        x = name.find('(')  # Encontra a posição do caractere '('
        y = name[:x-1]  # Obtém a substring antes do caractere '('
        return y  # Retorna a substring

    def retornar_tabelas(arquivo:str):  # Define uma função que retorna tabelas de um arquivo HTML
        
        df = pandas.DataFrame(columns=['Nome', 'Status'])  # Cria um DataFrame vazio na memória com as colunas 'Nome' e 'Status'

        htmldoc = html.fromstring(arquivo)  # Converte o arquivo HTML em um objeto lxml.html
        xml_doc = etree.tostring(htmldoc)  # Converte o objeto lxml.html em uma string XML
        json_ = xmltojson.parse(xml_doc)  # Converte a string XML em um objeto JSON

        for nomes_comercial in json.loads(json_)['html']['body']['div']['div']['div'][0]['div'][0]['div']['div']['div']['div'][5]['div']['div'][1]['div']['div']['div'][1]['div'][1]['div'][2]['div'][1]['div']['div']['div']:
            agentes_comercial = nomes_comercial['div'][0]['@title']
            status_agentes_comercial = nomes_comercial['div'][1]['@title']

            df.loc[len(df)] = [extract_name(agentes_comercial), status_agentes_comercial]
            #Estamos criando uma lista dentro desse dataframe (df) com os nomes e os status dos agentes
        
        for nomes_care in json.loads(json_)['html']['body']['div']['div']['div'][0]['div'][0]['div']['div']['div']['div'][1]['div']['div'][1]['div']['div']['div'][1]['div'][1]['div'][2]['div'][1]['div']['div']['div']:
            agentes_care = nomes_care['div'][0]['@title']
            status_agentes_care = nomes_comercial['div'][1]['@title']
            
            df.loc[len(df)] = [extract_name(agentes_care), status_agentes_care]
            #Estamos criando uma lista dentro desse dataframe (df) com os nomes e os status dos agentes

        return df

    driver.refresh()  # Recarrega a página

    time.sleep(10)  # Espera 10 segundos após recarregar a página para pegar os status de cada pessoa novamente

    variavel = retornar_tabelas(driver.page_source)  # Pega novamente o status das pessoas
    print(variavel)

    for usuario in variavel.iterrows():  # Aqui está iterando as linhas do df para imprimir os nomes dos agentes um em baixo do outro
        nome = usuario[1].Nome  # Obtém o nome do usuário dentro da coluna Nome. colocamos o [1] somente para pegar a informação dentro da linha. [1] é padrão.
        status = usuario[1].Status  # Obtém o status dentro da coluna Status

        if status not in ["free", "logged off", "interaction inbound"]:  # Verifica se o status do usuário não é "free", "logged off" ou "interaction inbound"
            
            link = 'https://teams.microsoft.com/_#/conversations/19:69d402e97da749bca4f378761a94248a@thread.v2?ctx=chat'
            # Define o link para a conversa no Teams

            print(f"Enviando mensagem de Pausa ou ACW para {nome} através do link {link}")  # Imprime uma mensagem informando que está enviando uma mensagem de Pausa ou ACW para o usuário

            driver.get(link)  # Abre o link do Teams 

            time.sleep(15)  # Espera 15 segundos para garantir que a página do Teams foi carregada completamente

            pyautogui.click(1787, 912)  # Clica na área de mensagem
            pyautogui.write(f"@{nome} ")  # Escreve a mensagem
            pyautogui.click(1815, 768)  # Clica na área da pessoa
            pyautogui.click(1787, 912)  # Clica na área de mensagem
            pyautogui.write('Atencao ACW / PAUSA')  # Escreve a mensagem
            pyautogui.press('enter')  # Envia a mensagem
            
            time.sleep(2)  # Espera 2 segundos antes de prosseguir

        else:
            print(f"{nome} não está em Pausa ou em ACW. Não é necessário enviar mensagem.")  # Imprime uma mensagem informando que o usuário não está em Pausa ou em ACW e que não é necessário enviar mensagem.

    time.sleep(5)  # Espera 5 segundos antes de reiniciar o loop

