from datetime import date
import os
from pathlib import Path
import re

from PyPDF2 import PdfReader
import pandas as pd

def Bice_parser(input: Path, output: Path):
    info_cartera: list[dict] = []
    info_movimientos: list[dict] = []

    for file in input.iterdir():
        if ".pdf" not in file.name :
            continue

        reader = PdfReader(file)

        text = ''

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text()

        _,nombre = text.split("Cliente: ")
        nombre,_ = nombre.split("Rut", maxsplit=1)
        nombre = nombre.replace(".", "")

        _,rut = text.split("Rut:  ")
        rut,_ = rut.split("\n", maxsplit=1)

        _,fecha = text.split("Fecha Emisión: ", maxsplit=1)
        fecha,_ = fecha.split("\n", maxsplit=1)
        dia, mes, ano = fecha.split("-")
        fecha = date(int(ano), int(mes), int(dia))

        text_cartera: list[str] = []
        text_movimientos: list[str] = []

        if "DETALLE DE CARTERAS" in text:
            text_cartera = text.split("DETALLE DE CARTERAS")[1:]

        for data in text_cartera:
            if "DETALLE DE MOVIMIENTOS" in data:
                text_movimientos = data.split("DETALLE DE MOVIMIENTOS")[1:]
                data,_ = data.split("DETALLE DE MOVIMIENTOS", maxsplit=1)
            
            data_carteras = data.split("Detalle Cartera")[1:]

            for rows in data_carteras:
                rows = rows.split("\n")
                clase_activo = rows.pop(0)
                clase_activo = clase_activo[1:]
                clase_activo, moneda = clase_activo.split(" en ")

                if 'US' in moneda:
                    moneda = 'USD'
                else:
                    moneda = 'CLP'
                
                rows = rows[2:]
                nemotecnico = ''
                cuenta = ''
                nemotecnico = ''
                cantidad = ''
                precio_mercado = ''
                valor_mercado = ''
                precio_compra = ''
                valor_compra = ''

                aux_renta_fija = ''

                for row in rows:
                    if clase_activo == 'Renta Fija':
                        if "Subtotal" in row:
                            continue
                        if not re.search(r"[a-zA-Z]", row):
                            continue
                        if 'glosario' in row:
                            break
                        
                        aux_renta_fija += row
                        if aux_renta_fija[-1].isdigit():
                            nemotecnico, aux_renta_fija = aux_renta_fija.split("Fecha Compra : ")
                            nemotecnico,_ = nemotecnico.split(" (")
                            row_data = aux_renta_fija.split(" ")
                            fecha_compra = row_data.pop(0)
                            cuenta = fecha_compra[10:]
                            fecha_compra = fecha_compra[:10]
                            cantidad = float(row_data[0].replace('.', '').replace(',', '.'))
                            precio_compra = float(row_data[4].replace('.', '').replace(',', '.'))
                            precio_mercado = float(row_data[6].replace('.', '').replace(',', '.'))
                            valor_compra = ""
                            valor_mercado = float(row_data[7].replace('.', '').replace(',', '.'))

                            info_cartera.append({
                                "Fecha": fecha,
                                "Nombre": nombre,
                                "Rut": rut,
                                "Cuenta": cuenta,
                                "Nemotecnico": nemotecnico,
                                "Moneda": moneda,
                                "ISIN": "",
                                "CUSIP": "",
                                "Cantidad": cantidad,
                                "Precio_Mercado": precio_mercado,
                                "Valor_Mercado": valor_mercado,
                                "Precio_Compra": precio_compra,
                                "Valor_Compra": valor_compra,
                                "Interes_Acum": "",
                                "Contraparte": "BICE",
                                "Clase_Activo": clase_activo,
                            })

                            nemotecnico = ''
                            cuenta = ''
                            nemotecnico = ''
                            cantidad = ''
                            precio_mercado = ''
                            valor_mercado = ''
                            precio_compra = ''
                            valor_compra = ''
                            aux_renta_fija = ''

                    elif clase_activo == "Renta Variable":
                        if 'Subtotal' in row:
                            continue
                        if not re.search(r"[a-zA-Z]", row):
                            continue
                        if 'glosario' in row:
                            break

                        row_data = row.split(" ")
                        nemotecnico = row_data[0]
                        cuenta = row_data[1]
                        cantidad = float(row_data[2].replace('.', '').replace(',', '.'))
                        cantidad += float(row_data[3].replace('.', '').replace(',', '.'))
                        cantidad += float(row_data[4].replace('.', '').replace(',', '.'))
                        precio_compra = float(row_data[5].replace('.', '').replace(',', '.'))
                        precio_mercado = float(row_data[6].replace('.', '').replace(',', '.'))
                        valor_mercado = float(row_data[8].replace('.', '').replace(',', '.'))
                        valor_compra = ''
                    
                    elif clase_activo == "Fondos Mutuos":
                        if 'Subtotal' in row:
                            continue
                        if not re.search(r"[a-zA-Z]", row):
                            continue
                        if 'glosario' in row:
                            break
                        
                        row_data = row.split(" ")
                        nemotecnico = " ".join(row_data[:-5])
                        row_data = row_data[-5:]
                        cuenta = row_data[0]
                        precio_mercado = float(row_data[1].replace('.', '').replace(',', '.'))
                        cantidad = float(row_data[2].replace('.', '').replace(',', '.')) 
                        cantidad+= float(row_data[3].replace('.', '').replace(',', '.'))
                        valor_mercado = float(row_data[4].replace('.', '').replace(',', '.'))

                    elif clase_activo == "Intermediación Financiera":
                        if "Subtotal" in row:
                            continue
                        if not re.search(r"\d+", row) and not re.search(r"[a-zA-Z]", row):
                            continue
                        if 'glosario' in row:
                            break

                        row_data = row.split(" ")
                        nemotecnico = " ".join(row_data[:-9])
                        if nemotecnico == "":
                            nemotecnico = info_cartera[-1]["Nemotecnico"]
                        row_data = row_data[-9:]
                        cuenta = row_data[0]
                        cantidad = float(row_data[3].replace('.', '').replace(',', '.'))
                        precio_compra = float(row_data[4].replace('.', '').replace(',', '.'))
                        precio_mercado = float(row_data[5].replace('.', '').replace(',', '.'))
                        valor_mercado = float(row_data[8].replace('.', '').replace(',', '.'))
                        valor_compra = float(row_data[6].replace('.', '').replace(',', '.'))

                    elif clase_activo == "Depósitos a Plazo":
                        if "Subtotal" in row:
                            continue
                        if not re.search(r"\d+", row) and not re.search(r"[a-zA-Z]", row):
                            continue
                        if 'valorizados' in row:
                            continue
                        if 'glosario' in row:
                            break
                        
                        row_data = row.split(" ")
                        nemotecnico = row_data[-7]
                        cuenta = rut
                        valor_mercado = float(row_data[-2].replace('.', '').replace(',', '.'))
                        cantidad = float(row_data[-1].replace('.', '').replace(',', '.'))

                    else:
                        print(f"Tipo de activo desconocido: {clase_activo}")
                        continue

                    if clase_activo != 'Renta Fija':
                        info_cartera.append({
                            "Fecha": fecha,
                            "Nombre": nombre,
                            "Rut": rut,
                            "Cuenta": cuenta,
                            "Nemotecnico": nemotecnico,
                            "Moneda": moneda,
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": cantidad,
                            "Precio_Mercado": precio_mercado,
                            "Valor_Mercado": valor_mercado,
                            "Precio_Compra": precio_compra,
                            "Valor_Compra": valor_compra,
                            "Interes_Acum": "",
                            "Contraparte": "BICE",
                            "Clase_Activo": clase_activo,
                        })

                        nemotecnico = ''
                        cuenta = ''
                        nemotecnico = ''
                        cantidad = ''
                        precio_mercado = ''
                        valor_mercado = ''
                        precio_compra = ''
                        valor_compra = ''
        
        if len(text_movimientos) == 0 and "DETALLE DE MOVIMIENTOS" in text:
            text_movimientos = text.split("DETALLE DE MOVIMIENTOS")[1:]

        for data in text_movimientos:
            data_movimientos = data.split("ovimientos de ")[1:]

            for rows in data_movimientos:
                rows = rows.split("\n")
                clase_activo = rows.pop(0)
                clase_activo, moneda = clase_activo.split(" en ")
                _, cuenta = rows.pop(0).split(': ')
                del rows[0] # Títulos

                if 'US' in moneda:
                    moneda = 'USD'
                else:
                    moneda = 'CLP'

                for row in rows:
                    if "glosario" in row:
                        break
                    if clase_activo == 'Caja':
                        if "SALDO INICIAL" in row:
                            continue
                        if "SIN MOVIMIENTOS" in row:
                            break
                        if "SALDO FINAL" in row:
                            break

                        row_data = row.split(" ")
                        fecha_movimiento = row_data.pop(0)
                        dia_mov, mes_mov, ano_mov = fecha_movimiento.split('-')
                        fecha_movimiento = date(int(ano_mov), int(mes_mov), int(dia_mov))
                        folio = row_data.pop(0)

                        tipo = []
                        index_corte = 0

                        for i,d in enumerate(row_data):
                            if d == "":
                                index_corte = i
                                break
                            tipo.append(d)
                        
                        tipo = " ".join(tipo)
                        row_data = row_data[index_corte:]
                        row_data = [aux for aux in row_data if aux != '']

                        nemotecnico = ''
                        if re.search(r"[a-zA-Z]", row_data[0]):
                            nemotecnico = row_data.pop(0)

                        abono = float(row_data[0].replace('.', '').replace(',', '.'))
                        cargo = float(row_data[1].replace('.', '').replace(',', '.'))
                        monto = abono - cargo

                        info_movimientos.append({
                            "Fecha Movimiento": fecha_movimiento,
                            "Fecha Liquidación": fecha_movimiento,
                            "Nombre": nombre,
                            "RUT": rut,
                            "Cuenta": cuenta,
                            "Nemotecnico": nemotecnico,
                            "Moneda": moneda,
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": "",
                            "Precio": "",
                            "Monto": monto,
                            "Comision": 0,
                            "IVA": 0,
                            "Tipo": clase_activo,
                            "Descripcion": tipo,
                            "Concepto": "MOVIMIENTO",
                            "Folio": folio,
                            "Contraparte": "BICE",
                        })
                    elif clase_activo == "Títulos":
                        row_data = row.split(" ")
                        if len(row_data) < 2:
                            continue
                        
                        fecha_movimiento = row_data.pop(0)
                        dia_mov, mes_mov, ano_mov = fecha_movimiento.split('-')
                        fecha_movimiento = date(int(ano_mov), int(mes_mov), int(dia_mov))
                        folio = row_data.pop(0)
                        tipo = f"{row_data[0]} {row_data[1]}"
                        row_data = row_data[2:]

                        nemotecnico = []
                        index_corte = 0

                        for i,d in enumerate(row_data):
                            if d == "":
                                index_corte = i
                                break
                            nemotecnico.append(d)

                        nemotecnico = " ".join(nemotecnico)
                        row_data = row_data[index_corte:]
                        row_data = [aux for aux in row_data if aux != '']

                        cantidad = float(row_data[0].replace('.', '').replace(',', '.'))
                        precio = float(row_data[1].replace('.', '').replace(',', '.'))
                        monto = float(row_data[2].replace('.', '').replace(',', '.'))
                        
                        info_movimientos.append({
                            "Fecha Movimiento": fecha_movimiento,
                            "Fecha Liquidación": fecha_movimiento,
                            "Nombre": nombre,
                            "RUT": rut,
                            "Cuenta": cuenta,
                            "Nemotecnico": nemotecnico,
                            "Moneda": moneda,
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": cantidad,
                            "Precio": precio,
                            "Monto": monto,
                            "Comision": 0,
                            "IVA": 0,
                            "Tipo": clase_activo,
                            "Descripcion": tipo,
                            "Concepto": "OPERACIÓN",
                            "Folio": folio,
                            "Contraparte": "BICE",
                        })
                    else:
                        print(f"Tipo de movimiento desconocido: {clase_activo}")
                        continue

    df_cartera = pd.DataFrame(info_cartera)
    df_movimientos = pd.DataFrame(info_movimientos)

    with pd.ExcelWriter(output / f"Informe_{fecha.strftime("%Y%m%d")}.xlsx", engine="openpyxl") as writer:
        df_cartera.to_excel(writer, index=False, sheet_name="Cartera")
        df_movimientos.to_excel(writer, index=False, sheet_name="Movimientos")

if __name__ == "__main__":
    folder_input = Path("Input")
    folder_output = Path("Output")
    Bice_parser(folder_input, folder_output)
