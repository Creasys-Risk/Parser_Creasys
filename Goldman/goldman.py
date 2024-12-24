import re
import pandas as pd
from datetime import date
from PyPDF2 import PdfReader

calendar = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}

def GenerateResultList(client_info: list[str]) -> list[str]:
    client_result = []
    for client_data in client_info:
        if "TOTAL PORTFOLIO" in client_data:
            client_data,_ = client_data.split("TOTAL", maxsplit=1)

        if "in PercentageEstimated" not in client_data and "Annual Income" not in client_data:
            continue

        _,client_data = client_data.split("Annual Income", maxsplit=1)

        if "Statement Detail" in client_data:
            client_data,_ = client_data.split("Statement Detail")

        client_list = client_data.split("\n")

        for index, aux in enumerate(client_list):
            if "Period Ended" in aux:
                aux,_ = aux.split("Period Ended")
            if re.search(r"\d[.]", aux):
                if re.search(r"[a-zA-Z]", aux):
                    client_result.append(aux)
                elif index+1 < len(client_list):
                    client_list[index+1] = aux + ' ' + client_list[index+1]

    return client_result

def GoldmanParser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    # with open("output.txt", "w") as f:
    #     f.write(text)

    _,date_list = text.split("Period Covering ", maxsplit=1)
    date_list,_ = date_list.split("\n", maxsplit=1)
    date_list = date_list.split(" ")
    date_document = date(int(date_list[2].replace(',', '')), calendar[date_list[4]], int(date_list[5].replace(',', '')))

    _, portfolios = text.split("PORTFOLIO INFORMATION\n", maxsplit=1)
    portfolios,data = portfolios.split("DUPLICATE COPIES OF THIS ACCOUNT STATEMENT ARE BEING SENT TO:\n", maxsplit=1)
    _,portfolios = portfolios.split("INDIVIDUAL PORTFOLIOS\n")
    portfolios = portfolios.split("\n")

    # portfolios = re.split(r"\d{2,3}", portfolios)

    clients: list[str] = []
    portfolio_numbers: dict[str, str] = {}

    for line in portfolios:
        if "XXX-XX" in line:
            lines = re.split(r"(\d{2,3})", line)
            portfolio_number = ""
            for aux in lines:
                if " XXX-XX" in aux:
                    portfolio_number = ""
                    client_name,_ = aux.split(" XXX-XX")
                    clients.append(client_name)
                else:
                    portfolio_number += aux
                    portfolio_numbers[client_name] = portfolio_number

    results_public_eq = []
    results_fixed_income = []

    for client in clients:
        client_data = data.split(client)
        del client_data[0]
        client_text = ""
        for aux in client_data:
            exit_clause = False
            if "THIS PAGE INTENTIONALLY LEFT BLANK" in aux:
                aux,_ = aux.split("THIS PAGE INTENTIONALLY LEFT BLANK", maxsplit=1)
                exit_clause = True
            if "See the Bank Disclosures for Specific Disclosure Information" in aux:
                aux,_ = aux.split("See the Bank Disclosures for Specific Disclosure Information", maxsplit=1)
                exit_clause = True
            client_text += aux

            if exit_clause:
                break

        client_text = client_text.replace(" FIXED INCOME", "")

        public_eq_client: list[str] = client_text.split("PUBLIC EQUITY")
        fixed_income_client: list[str] = client_text.split("FIXED INCOME")

        del public_eq_client[0]
        del fixed_income_client[0]
        
        public_eq_result: list[str] = GenerateResultList(public_eq_client)
        fixed_income_result: list[str] = GenerateResultList(fixed_income_client)

        for row in public_eq_result:
            if "TOTAL" in row:
                continue

            row = row.replace(",", "")
            row_data = row.split(" ")
            row_data = [x for x in row_data if x != '']
            row_data = [x.replace('(', '-') for x in row_data]
            row_data = [x.replace(')', '') for x in row_data]

            if len(row_data) < 8:
                continue

            try:
                float(row_data[3])
                quantity = float(row_data[0])
                market_price = float(row_data[1])
                market_value = float(row_data[2])
                mnemonic = ' '.join([x for x in row_data if re.search(r"[a-zA-Z]", x)])
            except ValueError:
                continue

            try:
                float(row_data[8])
                accrued_income = float(row_data[3])
                unit_cost = float(row_data[4])
                cost_basis = float(row_data[5])
            except ValueError:
                accrued_income = 0.0
                unit_cost = float(row_data[3])
                cost_basis = float(row_data[4])

            results_public_eq.append({
                "Fecha": date_document,
                "Nombre": client,
                "Rut": "",
                "Cuenta": portfolio_numbers[client],
                "Nemotecnico": mnemonic,
                "Moneda": "",
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": quantity,
                "Precio_Mercado": market_price,
                "Valor_Mercado": market_value,
                "Precio_Compra": unit_cost,
                "Valor_Compra": cost_basis,
                "Interes_Acum": accrued_income,
                "Contraparte": "GOLDMAN SACHS",
                "Clase_Activo": "PUBLIC EQUITY",
            })
        
        for row in fixed_income_result:
            if "TOTAL" in row:
                continue

            row = row.replace(",", "")
            row = row.replace(")", " ")
            row = row.replace("(", "-")
            row_data = row.split(" ")
            row_data = [x for x in row_data if x != '']

            if len(row_data) < 8:
                continue

            try:
                float(row_data[3])
                quantity = float(row_data[0])
                market_price = float(row_data[1])
                market_value = float(row_data[2])
                mnemonic = ' '.join([x for x in row_data if re.search(r"[a-zA-Z]", x)])
            except ValueError:
                continue

            try:
                float(row_data[8])
                accrued_income = float(row_data[3])
                unit_cost = float(row_data[4])
                cost_basis = float(row_data[5])
            except ValueError:
                accrued_income = 0.0
                unit_cost = float(row_data[3])
                cost_basis = float(row_data[4])

            results_fixed_income.append({
                "Fecha": date_document,
                "Nombre": client,
                "Rut": "",
                "Cuenta": portfolio_numbers[client],
                "Nemotecnico": mnemonic,
                "Moneda": "",
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": quantity,
                "Precio_Mercado": market_price,
                "Valor_Mercado": market_value,
                "Precio_Compra": unit_cost,
                "Valor_Compra": cost_basis,
                "Interes_Acum": accrued_income,
                "Contraparte": "GOLDMAN SACHS",
                "Clase_Activo": "FIXED INCOME",
            })

    df = pd.DataFrame(results_public_eq + results_fixed_income)
    df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == '__main__':
    import os

    files = os.listdir("./input")

    for file in files:
        if ".pdf" in file:
            filename,_ = file.split('.')
            GoldmanParser(filename)
