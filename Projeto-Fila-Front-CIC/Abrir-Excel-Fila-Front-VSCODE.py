import win32com.client
import re

e = win32com.client.Dispatch("Excel.Application")
e.Visible = 1
e.Workbooks.Add()

pedidos = 281
cotacao = 117
outros = 4
cs = 40
correcaodpc = 6
dpc = 60

e.Cells(1, 1).Value = "Fila Front"
e.Cells(1, 2).Value = "10h"  # tenho que alterar manualmente essa hora a cada hora (10h, 11h, 12h, etc)

e.Cells(2, 1).Value = "Pedidos"
e.Cells(2, 2).Value = pedidos

e.Cells(3, 1).Value = "Cotação"
e.Cells(3, 2).Value = cotacao

e.Cells(4, 1).Value = "Outros"
e.Cells(4, 2).Value = outros

e.Cells(5, 1).Value = "CS"
e.Cells(5, 2).Value = cs

e.Cells(6, 1).Value = "Correção DPC"
e.Cells(6, 2).Value = correcaodpc

e.Cells(7, 1).Value = "DPC"
e.Cells(7, 2).Value = dpc

e.Cells(8, 1).Value = "TOTAL"
e.Cells(8, 2).Value = "=SUM(B2,B3,B4,B6)"  # soma os valores diretamente na fórmula
