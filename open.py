import os
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
print("BASE_DIR:", BASE_DIR)

caminho_excel = os.path.join(BASE_DIR, "UCG-FE-1231010B-26-04.xlsm")
print("CAMINHO COMPLETO:", caminho_excel)
print("EXISTE?", os.path.exists(caminho_excel))

app = xw.App(visible=True)
wb = app.books.open(caminho_excel)