import re
from PyPDF2 import PdfReader
import pandas as pd
import datetime

def extract_data_portfolio(type_data: str, text: str, filename: str, fecha: datetime.date, cuenta: str, nombre: str) -> list[dict]:
    if type_data not in text:
        print(f"No se encontró información de {type_data} en el archivo: {filename}.pdf")
        return[]
    
    info = []
    full_data = text.split(type_data.upper())

    for data in full_data[1:]:
        if f"Total {type_data.title()}" in data:
            data,_ = data.split(f"Total {type_data.title()}")
        data = data.split("Total")

        sub_instrumento = ''
        for row in data:
            if '%' not in row and 'TasaCosto Valor Mercado' not in row:
                continue

            if row[0] != '%':
                if '%' in row:
                    aux,_ = row.split('\n%')
                else:
                    aux,_ = row.split('\nTasaCosto Valor Mercado')
                _,aux = aux.split('\n')
                sub_instrumento = aux
            else:
                aux = row.split("\n")
                if len(aux) < 3:
                    continue
                for d in aux[2:-1]:
                    d = d.split(" ")
                    if type_data == "RENTA FIJA":
                        nemo = " ".join(d[0:-12])
                        moneda = d[-2]
                        cantidad = float(d[-8].replace(".","").replace(",","."))
                        precio_mercado = float(d[-5].replace(".","").replace(",","."))
                        valor_mercado = float(d[-4].replace(".","").replace(",","."))
                        precio_compra = float(d[-7].replace(".","").replace(",","."))
                        valor_compra = float(d[-6].replace(".","").replace(",","."))
                        cusip = d[-9]
                    else:
                        nemo = " ".join(d[1:-7])
                        moneda = d[-2]
                        cantidad = float(d[-7].replace(".","").replace(",","."))
                        precio_mercado = float(d[-4].replace(".","").replace(",","."))
                        valor_mercado = float(d[-3].replace(".","").replace(",","."))
                        precio_compra = float(d[-6].replace(".","").replace(",","."))
                        valor_compra = float(d[-5].replace(".","").replace(",","."))
                        cusip = ""
                    info.append({
                        "Fecha": fecha,
                        "Nombre": nombre,
                        "Rut": "",
                        "Cuenta": cuenta,
                        "Nemotecnico": nemo,
                        "Moneda": moneda,
                        "ISIN": "",
                        "CUSIP": cusip,
                        "Cantidad": cantidad,
                        "Precio_Mercado": precio_mercado,
                        "Valor_Mercado": valor_mercado,
                        "Precio_Compra": precio_compra,
                        "Valor_Compra": valor_compra,
                        "Interes_Acum": "",
                        "Contraparte": "BANCHILE",
                        "Clase_Activo": f"{sub_instrumento.upper()}",
                    })

    return info

def extract_data_movement(text: str, filename: str, fecha: datetime.date, cuenta: str, nombre: str) -> list[dict]:
    _,data = text.split("CUENTA CORRIENTE ")
    data,_ = data.split("*", maxsplit=1)
    data = data.split("\n")

    if data[1] == '00/00/0000 0,0000':
        print(f"No se encontró información de movimientos en el archivo: {filename}.pdf")
        return []

    info = []
    for d in data[1:-1]:
        aux = d.split(" ")
        instrumento=aux.pop(0)
        
        emisor = ""
        index_aux = 0
        for i,c in enumerate(aux):
            if "/" in c:
                index_aux = i
                break
            else:
                emisor += f"{c} "
        
        aux = aux[index_aux:]
        fecha_liquidacion = datetime.datetime.strptime(aux.pop(0), "%d/%m/%Y").date() 
        
        movimiento = ""
        index_aux = 0
        for i,c in enumerate(aux):
            if "/" in c:
                index_aux = i
                break
            else:
                movimiento += f"{c} "

        aux = aux[index_aux:]
        fecha_movimiento = datetime.datetime.strptime(aux.pop(0), "%d/%m/%Y").date() 
        monto = float(aux[0].replace(".","").replace(",","."))

        if len(aux) == 5:
            precio = float(aux[1].replace(".","").replace(",","."))
            moneda = aux[2]
            unidades = float(aux[3].replace(".","").replace(",","."))
            saldo = float(aux[4].replace(".","").replace(",","."))
            concepto = "OPERACION"
        else:
            precio = ""
            moneda = aux[1]
            unidades = float(aux[2].replace(".","").replace(",","."))
            saldo = float(aux[3].replace(".","").replace(",","."))
            concepto = "MOVIMIENTO"

        info.append({
            "Fecha Movimiento": fecha_movimiento,
            "Fecha Liquidación": fecha_liquidacion,
            "Nombre": nombre,
            "RUT": "",
            "Cuenta": cuenta,
            "Nemotecnico": instrumento,
            "Moneda": moneda,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": unidades,
            "Precio": precio,
            "Monto": monto,
            "Comision": 0,
            "IVA": 0,
            "Tipo": movimiento,
            "Descripcion": movimiento,
            "Concepto": concepto,
            "Folio": "",
            "Contraparte": "BANCHILE",
        })

    return info

def BanChile_Parser(filename: str) -> tuple[list[dict], list[dict]]:
    reader = PdfReader(f"./input/{filename}.pdf")

    text_full = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text_full += page.extract_text()

    with open("output_banchile.txt", "w", encoding="utf-8") as f:
        f.write(text_full)

    data_cartera = []
    data_movimientos = []
    text_full = text_full.split("Período del Estado de Cuenta:")

    for text in text_full[1:]:
        _,fecha = text.split("CARTERA DE INVERSIONES  AL ", maxsplit=1)
        fecha,_ = fecha.split("\n", maxsplit=1)
        day, month, year = fecha.split("/")
        fecha = datetime.date(day=int(day), month=int(month), year=int(year))

        _,nombre = text.split("Su Asesor de Inversiones:\n")
        nombre,_ = nombre.split("At.", maxsplit=1)
        nombre = nombre.replace("\n", " ")
        nombre = nombre.split(" ")

        for index, aux in enumerate(nombre):
            aux = aux[1:]
            uppercase_index = [i for i, c in enumerate(aux) if c.isupper()]
            if len(uppercase_index) != 0:
                nombre = " ".join([aux[uppercase_index[0]:]] + nombre[index+1:])
                break

        cuenta,_ = text.split("Subcuenta:")
        cuenta = cuenta.split("\n")[-1]
        cuenta = [int(num) for num in re.findall(r'\d+', cuenta)][0]

        data_cartera+= extract_data_portfolio("FONDOS MUTUOS", text, filename, fecha, cuenta, nombre)
        data_cartera+= extract_data_portfolio("RENTA FIJA", text, filename, fecha, cuenta, nombre)
        data_cartera+= extract_data_portfolio("RENTA VARIABLE", text, filename, fecha, cuenta, nombre)
        data_cartera+= extract_data_portfolio("OTRAS INVERSIONES", text, filename, fecha, cuenta, nombre)
        data_movimientos = extract_data_movement(text, filename, fecha, cuenta, nombre)

    return (data_cartera, data_movimientos)

if __name__ == "__main__":
    import os

    files = os.listdir("./input")
    data_cartera = []
    data_movimientos = []

    for file in files:
        if file == ".gitkeep":
            continue
        filename,_ = file.split('.')
        cartera, movimientos = BanChile_Parser(filename)
        data_cartera += cartera
        data_movimientos += movimientos
        
    date: datetime.date = data_cartera[0]["Fecha"]
    df_cartera = pd.DataFrame(data_cartera)
    df_movimientos = pd.DataFrame(data_movimientos)

    with pd.ExcelWriter(f"./output/Informe_{date.strftime('%YYYY%MM%DD')}.xlsx", engine="openpyxl") as writer:
        df_cartera.to_excel(writer, index=False, sheet_name="Cartera")
        df_movimientos.to_excel(writer, index=False, sheet_name="Movimientos")
