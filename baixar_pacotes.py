from email import header
from minhas_funcoes.bases_de_dados import baseDeDadosGoogle
import requests
import pandas as pd
import xlwings as xw
import time
import datetime
import os
import ctypes
import numpy as np

instancia_teste = baseDeDadosGoogle()

app = xw.App(visible = False, add_book = False)

wb = xw.Book(r"Inventário de Pacotes.xlsm")
wb.activate()
ws = wb.sheets["base_forms"]
app.screen_updating = False
ws["a:z"].clear_contents()
ws["A1"].options(index=False).value = instancia_teste.preparar_tabela()
app.screen_updating = True
