from openpyxl import load_workbook
import json
from openpyxl.descriptors import MinMax, Sequence
import time
import sys
from tkinter import filedialog

SelecArquivos = filedialog.askopenfilenames()
arquivos = ["C:/Users/alexandre.borges/Documents/TCU/pubs_secao3_2020-10_(outubro).xlsx"]

for item in SelecArquivos:
  print(item)
  wb = load_workbook(item)
  ws = wb.active
  col = ws['F']
  linha = 2

  dicionario = [
      "Banco do Brasil",
      "Banco do Nordeste",
      "Banco da Amazônia",
      "Caixa Econômica Federal",
      "Banco Central",
      "Casa da Moeda",
      "Comissão de Valores Mobiliários",
      "Susep",
      "BNDES",
      "Associação Brasileira de Fundos Garantidores",
      "Banco Nacional de Desenvolvimento Econômico e Social"
  ]


  def filtro(dicionario, max_row, linha):
      while linha <= ws.max_row:
          for item in dicionario:
              try:
                  if item in ws['F' + str(linha)].value:
                      print(ws['F' + str(linha)].value)
                      break
                  else:
                      if item == dicionario[-1]:
                          ws.delete_rows(linha)
                      else:
                          continue
              except:
                  ws.delete_rows(linha)
          linha = linha + 1

  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)
  filtro(dicionario, ws.max_row, linha)

  '''ultima = ws.max_row
  while ws['F' + str(ultima)].value != "":
    filtro(dicionario, ws.max_row, linha)'''
    
    
  wb.save(item)
