from openpyxl import load_workbook
import json
from openpyxl.descriptors import MinMax, Sequence
import time
import sys
wb = load_workbook(
    "C:/Users/alexandre.borges/Documents/TCU/dispensas_2020-09-21_09_46.xlsx")
ws = wb.active
col = ws['F']
linha = 2

dicionario = [
    "Comissão de Valores Mobiliários", "Casa da Moeda do Brasil",
    "Fundação Instituto Brasileiro de Geografia",
    "Instituto Nacional da Propriedade Industrial",
    "Instituto Nacional de Metrologia",
    "NAV Brasil Serviços de Navegação Aérea S.A",
    "Superintendência de Seguros Privados",
    "Centrais Elétricas Brasileiras S/A",
    "Centro de Pesquisa de Energia Elétrica",
    "Comissão Nacional de Energia Nuclear", "Eletrobrás Participações S.A",
    "Eletrobrás Termonuclear S/A", "Furnas Centrais Elétricas S/A",
    "Indústrias Nucleares do Brasil S/A", "Itaipu Binacional",
    "Nuclebrás Equipamentos Pesados S/A",
    "Agência Brasileira Gestora de Fundos Garantidores e Garantias S.A",
    "Agência Especial de Financiamento Industrial",
    "Banco Nacional de Desenvolvimento Econômico e Social", "BNDES",
    "Financiadora de Estudos e Projetos"
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
  
  
wb.save(
    "C:/Users/alexandre.borges/Documents/TCU/dispensas_2020-09-21_09_46.xlsx")
