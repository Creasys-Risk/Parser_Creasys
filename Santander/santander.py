from datetime import date
import os
from pathlib import Path
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

def process_bonds(rows: list[str], nombre: str, cuenta: str, fecha: date):
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
            "Fecha": fecha,
            "Nombre": nombre,
            "Rut": "",
            "Cuenta": cuenta,
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

def process_funds(rows: list[str], nombre: str, cuenta: str, fecha: date):
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
            "Fecha": fecha,
            "Nombre": nombre,
            "Rut": "",
            "Cuenta": cuenta,
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

def process_movements(rows: list[str], nombre: str, cuenta: str, fecha: date):
    row_aux = ''
    folio = ''
    movements: list[str] = []
    cartera_movimientos: list[dict] = []

    for row in rows:
        row_aux += f" {row}"

        if ' - ' in row_aux:
            movements.append(row_aux)
            row_aux = ''

    for move in movements:
        move = move.split(" ")
        nemo = " ".join(move[:-6])
        move = move[-6:]

        if move[1] == '-':
            # Se trata del balance inicial de la cuenta
            folio = nemo
            continue

        if len(move[0]) > 3:
            nemo += f" {move[0][:-3]}"
            move[0] = move[0][-3:]

        fecha_liquidacion = move[1]
        day, month, year = fecha_liquidacion.split('-')
        fecha_liquidacion = date(int(f"20{year}"), calendar_mini[month], int(day))

        fecha_movimiento = move[2]
        day, month, year = fecha_movimiento.split('-')
        fecha_movimiento = date(int(f"20{year}"), calendar_mini[month], int(day))

        moneda = move[0]
        monto = move[3] if move[3] != '-' else move[4]
        monto = float(monto.replace(',', ''))
        descripcion = "Deposito" if move[3] != '-' else "Retiro"

        cartera_movimientos.append({
            "Fecha Movimiento": fecha_movimiento,
            "Fecha Liquidación": fecha_liquidacion,
            "Nombre": nombre,
            "RUT": "",
            "Cuenta": cuenta,
            "Nemotecnico": nemo,
            "Moneda": moneda,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": "",
            "Precio": "",
            "Monto": monto,
            "Comision": 0,
            "IVA": 0,
            "Tipo": "Caja",
            "Descripcion": descripcion,
            "Concepto": "MOVIMIENTO",
            "Folio": folio,
            "Contraparte": "SANTANDER",
        })

    return cartera_movimientos

def Santander_Parser(input: Path, output: Path):
    info_cartera: list[dict] = []
    info_movimientos: list[dict] = []
    fecha: date

    for file in input.iterdir():
        if ".pdf" not in file.name:
            continue

        reader = PdfReader(file)

        text = ''

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text()

        _,aux = text.split("Global Vision of the Market and Your Account ")
        aux,_ = aux.split("\n", maxsplit=1)
        _,aux = aux.split(" to ")
        aux = aux.replace(" - ", " ")
        aux = aux.split(" ")
        
        day, month, year = aux[0].split("-")
        fecha = date(int(f"20{year}"), calendar_mini[month], int(day))
        cuenta= aux[1]
        nombre = " ".join(aux[2:])

        data_cartera = text.split("Total(%)\n")

        for table in data_cartera[1:]:
            rows = table.split("\n")

            if "CASH" in rows[0]:
                continue
            if "ANNUAL" in rows[1]:
                info = process_bonds(rows, nombre, cuenta, fecha)
            else:
                info = process_funds(rows, nombre, cuenta, fecha)
            
            info_cartera += info

        data_movimientos = text.split("Transactions\nDetail ccy. Value Date Booking Date Deposit Withdraws Balance\n")
        del data_movimientos[0]

        for table in data_movimientos:
            rows,_ = table.split("\nPlease see important information on the last page")
            rows = rows.split("\n")
            
            info = process_movements(rows, nombre, cuenta, fecha)

            info_movimientos += info

    df_cartera = pd.DataFrame(info_cartera)
    df_movimientos = pd.DataFrame(info_movimientos)

    with pd.ExcelWriter(output / f"InformeSantander_{fecha.strftime("%Y%m%d")}.xlsx", engine="openpyxl") as writer:
        df_cartera.to_excel(writer, index=False, sheet_name="Cartera")
        df_movimientos.to_excel(writer, index=False, sheet_name="Movimientos")

if __name__ == "__main__":
    folder_input = Path("Input")
    folder_output = Path("Output")
    Santander_Parser(folder_input, folder_output)