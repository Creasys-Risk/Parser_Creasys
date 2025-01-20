import datetime
import re
import pandas as pd
from PyPDF2 import PdfReader

def Extract_Equity(text: str, filename: str, account: str, date: datetime.date) -> list[dict]:
    if "Equity Detail\n" not in text:
        print(f"No se encontró información de Equity en el archivo: {filename}.pdf")
        return []

    _,data = text.split("Equity Detail\n")

    info = data.split("Unrealized Est. Annual Inc.")
    del info[-1]

    equity_rows: list[str] = []
    for i in info:
        lines = i.split("\n")
        for index, line in enumerate(lines):
            if "%" in line and "$" not in line:
                if "Equity" in line:
                    _,line = line.split("Equity")
                row = line + " " + lines[index+1] if '-' in lines[index+1] else line + " " + lines[index+2]
                equity_rows.append(row)

    results = []
    for row in equity_rows:
        row_data = row.split(" ")
        row_data = [x for x in row_data if x != '' and x != ')']
        row_data = [x.replace('(', '-') for x in row_data]
        
        price = float(row_data[0].replace(',', ''))
        quantity = float(row_data[1].replace(',', ''))
        value = float(row_data[2].replace(',', ''))
        nemotecnico = f"{row_data[-2]} {row_data[-1]}"

        results.append({
            "Fecha": date,
            "Nombre": "",
            "Rut": "",
            "Cuenta": account,
            "Nemotecnico": nemotecnico,
            "Moneda": "USD",
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": quantity,
            "Precio_Mercado": price,
            "Valor_Mercado": value,
            "Precio_Compra": "",
            "Valor_Compra": "",
            "Interes_Acum": 0,
            "Contraparte": "JP MORGAN",
            "Clase_Activo": "EQUITY",
        })

    return results

def Extract_Fixed_Income(text: str, filename: str, account: str, date: datetime.date) -> list[dict]:
    if "Cash & Fixed Income Detail" not in text:
        print(f"No se encontró información de Fixed Income en el archivo: {filename}.pdf")
        return []
    
    _,data = text.split("Cash & Fixed Income Detail")

    info = data.split("Unrealized Est. Annual Income")
    del info[0]
    del info[-1]

    fixed_income_rows: list[str] = []
    for i in info:
        lines = i.split("\n")
        row = ""
        for line in lines:
            if "Fixed Income" in line:
                aux = line.split("Fixed Income")
                line = aux[-1]
            if re.match(r'^\s{9}', line):
                if "This is the Annual Percentage Yield" not in row and "For the Period" not in row and "Page" not in row and "Price Quantity Value" not in row and "$" not in row:
                    fixed_income_rows.append(row)
                row = ""
            row += f" {line}"
        fixed_income_rows.append(row)

    results = []
    for row in fixed_income_rows:
        row_data = row.replace(f"{' '*6}", ";")
        row_data = row_data.split(";")
        row_data = [x for x in row_data if x != '']
        row_data = [x.replace('(', '-') for x in row_data]
        row_data = [x.replace(' )', '') for x in row_data]

        price = float(row_data[0].replace(',', '').replace(' ', ''))
        quantity = float(row_data[1].replace(',', '').replace(' ', ''))
        value = float(row_data[2].replace(',', '').replace(' ', ''))

        if re.search(r"[a-zA-Z]", row_data[4]):
            first_letter_index = next(i for i, char in enumerate(row_data[4]) if char.isalpha())
            nemotecnico = f"{row_data[4][first_letter_index:]}"
            unrealized_gain_loss = float(row_data[4][:first_letter_index].replace(',', '').replace(' ', ''))
            annual_income = 0.0
            annual_yield = 0.0
            accrued_interest = 0.0
        else:
            unrealized_gain_loss = float(row_data[4].replace(',', '').replace(' ', ''))

            aux = row_data[5].split(" ")
            aux = [x for x in aux if x != '']
            annual_income = aux[0]
            annual_yield = aux[1]
            nemotecnico = " ".join(aux[2:])

            aux_2 = row_data[6].split(" ")
            aux_2 = [x for x in aux_2 if x != '']
            accrued_interest = float(aux_2[0].replace(',', '').replace(' ', ''))
            nemotecnico += " " + " ".join(aux_2[1:])

        results.append({
            "Fecha": date,
            "Nombre": "",
            "Rut": "",
            "Cuenta": account,
            "Nemotecnico": nemotecnico,
            "Moneda": "USD",
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": quantity,
            "Precio_Mercado": price,
            "Valor_Mercado": value,
            "Precio_Compra": "",
            "Valor_Compra": "",
            "Interes_Acum": accrued_interest,
            "Contraparte": "JP MORGAN",
            "Clase_Activo": "FIXED INCOME",
        })

    return results

def JPM_Parser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    with open("output_jp_morgan.txt", "w", encoding="UTF-8") as f:
        f.write(text)

    if "ACCT." in text:
        _,account = text.split("ACCT. ", maxsplit=1)
        account,_ = account.split("\n", maxsplit=1)
    elif "Primary Account:" in text:
        _,account = text.split("Primary Account: ", maxsplit=1)
        account,_ = account.split("\n", maxsplit=1)
    else:
        _,account = text.split("²", maxsplit=1)
        account,_ = account.split("\n", maxsplit=1)

    _,date_aux = text.split("For the Period ", maxsplit=1)
    date_aux,_ = date_aux.split("\n", maxsplit=1)
    date_aux = date_aux.split(" ")
    date_aux = date_aux[2]

    month,day,year = date_aux.split("/")
    day = int(day)
    month = int(month)
    year = int(f"20{year}")

    date = datetime.date(year, month, day)

    results_equity = Extract_Equity(text, filename, account, date)
    results_fixed_income = Extract_Fixed_Income(text, filename, account, date)

    results = results_equity + results_fixed_income

    if len(results) > 0:
        df = pd.DataFrame(results_equity + results_fixed_income)
        df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == "__main__":
    import os

    files = os.listdir("./input")

    for file in files:
        if file == ".gitkeep":
            continue
        filename,_ = file.split('.pdf')
        JPM_Parser(filename)
