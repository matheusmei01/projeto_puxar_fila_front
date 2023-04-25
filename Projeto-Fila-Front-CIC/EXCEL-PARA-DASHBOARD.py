import win32com.client
import re

# IMPORTANTE: A CADA HORA MUDAR O NOME DO ARQUIVO  DA LINHA 46
# EXEMPLO: excel-para-dashboard , DEPOIS ADICIONAR 1 NO NOME FINAL DO ARQUIVO, DEPOIS 2, 3, ETC
# ISSO É PARA SALVAR VARIAS PLANILHAS DE EXCEL A CADA HORA DO DIA.

e = win32com.client.Dispatch("Excel.Application")
e.Visible = 1
wb = e.Workbooks.Add()

# define os valores a serem inseridos nas células
pedidos = 281
cotacao = 117
outros = 4
cs = 40
correcaodpc = 6
dpc = 60

# insere os valores nas células do Workbook
wb.Sheets(1).Cells(1, 1).Value = "Fila Front"
wb.Sheets(1).Cells(1, 2).Value = "10h"

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

# salva o Workbook em um arquivo
wb.SaveAs(r'C:\Users\br0mslvr\Desktop\Projeto-Fila-Front-CIC\excel-para-dashboard.xlsx')

# fecha o Workbook
wb.Close()

# encerra o Excel
e.Quit()
