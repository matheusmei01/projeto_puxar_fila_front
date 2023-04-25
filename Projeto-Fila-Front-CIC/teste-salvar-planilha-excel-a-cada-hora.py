import win32com.client
import re
import os
import datetime
import time

# cria uma instância do Excel
e = win32com.client.Dispatch("Excel.Application")
e.Visible = 1


while True:
    try:
        # define os valores a serem inseridos nas células
        pedidos = 281
        cotacao = 117
        outros = 4
        cs = 40
        correcaodpc = 6
        dpc = 60

        # adiciona um novo Workbook
        wb = e.Workbooks.Add()

        # insere os valores nas células do Workbook
        wb.Sheets(1).Cells(1, 1).Value = "Fila Front"
        wb.Sheets(1).Cells(1, 2).Value = datetime.datetime.now().strftime("%H:%M")

        wb.Sheets(1).Cells(2, 1).Value = "Pedidos"
        wb.Sheets(1).Cells(2, 2).Value = pedidos

        wb.Sheets(1).Cells(3, 1).Value = "Cotação"
        wb.Sheets(1).Cells(3, 2).Value = cotacao

        wb.Sheets(1).Cells(4, 1).Value = "Outros"
        wb.Sheets(1).Cells(4, 2).Value = outros

        wb.Sheets(1).Cells(5, 1).Value = "CS"
        wb.Sheets(1).Cells(5, 2).Value = cs

        wb.Sheets(1).Cells(6, 1).Value = "Correção DPC"
        wb.Sheets(1).Cells(6, 2).Value = correcaodpc

        wb.Sheets(1).Cells(7, 1).Value = "DPC"
        wb.Sheets(1).Cells(7, 2).Value = dpc

        wb.Sheets(1).Cells(8, 1).Value = "TOTAL"
        wb.Sheets(1).Cells(8, 2).Value = "=SUM(B2,B3,B4,B6)"

        # define o nome do arquivo a ser salvo
        file_name = "excel-para-dashboard_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = (r'C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\arquivos_excel', file_name)

        # salva o Workbook em um arquivo
        wb.SaveAs(file_path)

        # fecha o Workbook
        wb.Close()

        # encerra o Excel
        e.Quit()

        # log
        print(f"Arquivo salvo: {file_path}")

    except Exception as e:
        print(f"Erro ao salvar arquivo: {str(e)}")

    # aguarda 1 hora para executar novamente o loop
    time.sleep(3600)

# encerra o Excel
    e.Quit()
