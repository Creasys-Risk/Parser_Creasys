from datetime import date
import os
import re

from PyPDF2 import PdfReader
import pandas as pd

calendar_mini = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}

def process_bonds(rows: list[str]):
    bonds = []
    cartera_bonos = []
    for row in rows:
        if "Total" in row or "information" in row:
            break
        bonds.append(row)

    for i in range(int(len(bonds)/6)):
        index = i*6
        clean_row = f"{bonds[index]} {bonds[index+1]} {bonds[index+2]} {bonds[index+3]} {bonds[index+4]} {bonds[index+5]}"
        data = clean_row.split(" ")

        while not data[-10][-1].isdigit():
            data[-10] = data[-10][:-1]

        nemo = " ".join(data[0:-9])
        data = data[-9:]
        cantidad = data[0][9:]
        moneda = data[1]
        precio_compra = data[2][9:]
        _,precio_mercado = data[3].split(".", maxsplit=1)
        precio_mercado = precio_mercado[2:]
        _,valor_mercado = data[4].split(".", maxsplit=1)
        valor_mercado = valor_mercado[2:]
        interes_acum = data[5]

        cantidad = float(cantidad.replace(",", ""))
        precio_compra = float(precio_compra.replace(",", ""))
        precio_mercado = float(precio_mercado.replace(",", ""))
        valor_mercado = float(valor_mercado.replace(",", ""))
        interes_acum = float(interes_acum.replace(",", ""))

        cartera_bonos.append({
            "Fecha": "",
            "Nombre": "",
            "Rut": "",
            "Cuenta": "",
            "Nemotecnico": nemo,
            "Moneda": moneda,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": cantidad,
            "Precio_Mercado": precio_mercado,
            "Valor_Mercado": valor_mercado,
            "Precio_Compra": precio_compra,
            "Valor_Compra": "",
            "Interes_Acum": interes_acum,
            "Contraparte": "SANTANDER",
            "Clase_Activo": "BOND",
        })
        
    return cartera_bonos

def process_funds(rows: list[str]):
    funds = []
    cartera_fondos = []
    for row in rows:
        if "Total" in row or "information" in row:
            break
        funds.append(row)

    data = []
    for i in range(int(len(funds)/2)):
        index = i*2
        clean_row = f"{funds[index]} {funds[index+1]}"
        data = clean_row.split(" ")

        nemo = " ".join(data[:-10])
        nemo = f"{nemo} {data[-10][:12]}"
        cantidad = data[-10][12:]
        data = data[-9:]
        moneda = data[0]
        precio_compra = data[2]

        # Esta distinción se hace para separar las acciones de los fondos
        if "DIS -" in clean_row:
            valor_mercado = data[3]
            precio_mercado = data[5]
            clase_activo = "SHARE"
        else:
            precio_mercado = data[3]
            valor_mercado = data[6]
            clase_activo = "FUND"

        cantidad = float(cantidad.replace(",", ""))
        precio_compra = float(precio_compra.replace(",", ""))
        valor_mercado = float(valor_mercado.replace(",", ""))
        precio_mercado = float(precio_mercado.replace(",", ""))

        cartera_fondos.append({
            "Fecha": "",
            "Nombre": "",
            "Rut": "",
            "Cuenta": "",
            "Nemotecnico": nemo,
            "Moneda": moneda,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": cantidad,
            "Precio_Mercado": precio_mercado,
            "Valor_Mercado": valor_mercado,
            "Precio_Compra": precio_compra,
            "Valor_Compra": "",
            "Interes_Acum": "",
            "Contraparte": "SANTANDER",
            "Clase_Activo": clase_activo,
        })

    return cartera_fondos

folder_input = "./Input/"
folder_output = "./Output/"

files = [f for f in os.listdir(folder_input) if '.pdf' in f]

info_cartera: list[dict] = []
info_movimientos: list[dict] = []

for file in files:
    filename,_ = file.split('.pdf')
    reader = PdfReader(folder_input+file)

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    with open(f"{folder_output}{filename}.txt", "w", encoding="utf-8") as f:
        f.write(text)

    data = text.split("Total(%)\n")

    info_cartera = []
    for table in data[1:]:
        rows = table.split("\n")

        if "CASH" in rows[0]:
            continue
        if "ANNUAL" in rows[1]:
            info = process_bonds(rows)
        else:
            info = process_funds(rows)
        
        info_cartera += info

print(info_cartera)