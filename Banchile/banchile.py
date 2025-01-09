import re
from PyPDF2 import PdfReader
import pandas as pd
import datetime

def extract_data(type_data: str, text: str, filename: str) -> list[dict]:
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
                    nemo = " ".join(d[1:-7])
                    moneda = d[-2]
                    cantidad = float(d[-7].replace(".","").replace(",","."))
                    precio_mercado = float(d[-4].replace(".","").replace(",","."))
                    valor_mercado = float(d[-3].replace(".","").replace(",","."))
                    precio_compra = float(d[-6].replace(".","").replace(",","."))
                    valor_compra = float(d[-5].replace(".","").replace(",","."))
                    info.append({
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
                        "Valor_Compra": valor_compra,
                        "Interes_Acum": "",
                        "Contraparte": "",
                        "Clase_Activo": f"{type_data.title()}; {sub_instrumento.title()}",
                    })

    return info

def BanChile_Parser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text_full = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text_full += page.extract_text()

    with open("output_banchile.txt", "w", encoding="utf-8") as f:
        f.write(text_full)

    data = []
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

        data+= extract_data("FONDOS MUTUOS", text, filename)
        data+= extract_data("RENTA FIJA", text, filename)
        data+= extract_data("RENTA VARIABLE", text, filename)
        data+= extract_data("OTRAS INVERSIONES", text, filename)

    df = pd.DataFrame(data)
    df = df.assign(Fecha=fecha)
    df = df.assign(Cuenta=cuenta)
    df = df.assign(Nombre=nombre)
    df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == "__main__":
    import os

    files = os.listdir("./input")

    for file in files:
        if file == ".gitkeep":
            continue
        filename,_ = file.split('.')
        BanChile_Parser(filename)