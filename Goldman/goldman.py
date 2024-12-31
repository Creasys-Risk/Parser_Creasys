import re
import pandas as pd
import math
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
            if aux == "":
                continue
            if re.match(r"\d", aux[0]) and "%" not in aux:
                client_result.append(aux)
            elif re.search(r"[a-zA-Z]", aux) and re.search(r"\d", aux) and len(client_result) > 0:
                client_result[-1] += " " + aux

    client_result_concat = []

    for index, data in enumerate(client_result):
        if not re.search(r"[a-zA-Z]", data) and index+1 < len(client_result):
            client_result[index+1] = data + " " + client_result[index+1]
        else:
            client_result_concat.append(data)

    return client_result_concat

def GenerateResultProduct(client_result: list[str], product_name: str, date_document: date, client: str, portfolio_number: str, currency: str) -> list[dict]:
    results_product = []
    
    for row in client_result:
        if "TOTAL" in row:
            continue

        try:
            first_letter_index = next(i for i, char in enumerate(row) if char.isalpha())
        except StopIteration:
            continue
        
        mnemonic = row[first_letter_index:]
        row = row[:first_letter_index]
        row = row.replace(",", "")
        row = row.replace(")", " ")
        row = row.replace("(", "-")
        row_data = row.split(" ")
        row_data = [x for x in row_data if x != '']

        if len(row_data) < 6:
            continue

        try:
            float(row_data[3])
            quantity = float(row_data[0])
            market_price = float(row_data[1])
            market_value = float(row_data[2])
        except ValueError:
            continue

        if math.isclose(round(quantity*float(row_data[3])),round(float(row_data[4])), rel_tol=1e-2) or math.isclose(round(quantity*float(row_data[3])/100), round(float(row_data[4])), rel_tol=1e-2):
            accrued_income = 0.0
            unit_cost = float(row_data[3])
            cost_basis = float(row_data[4])
        else:
            accrued_income = float(row_data[3])
            unit_cost = float(row_data[4])
            cost_basis = float(row_data[5])

        results_product.append({
            "Fecha": date_document,
            "Nombre": client,
            "Rut": "",
            "Cuenta": portfolio_number,
            "Nemotecnico": mnemonic,
            "Moneda": currency,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": quantity,
            "Precio_Mercado": market_price,
            "Valor_Mercado": market_value,
            "Precio_Compra": unit_cost,
            "Valor_Compra": cost_basis,
            "Interes_Acum": accrued_income,
            "Contraparte": "GOLDMAN SACHS",
            "Clase_Activo": product_name,
        })

    return results_product

def GoldmanParser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    _,date_list = text.split("Period Covering ", maxsplit=1)
    date_list,_ = date_list.split("\n", maxsplit=1)
    date_list = date_list.split(" ")
    date_document = date(int(date_list[2].replace(',', '')), calendar[date_list[4]], int(date_list[5].replace(',', '')))

    _,currency = text.split("BASE CURRENCY : ", maxsplit=1)
    currency,_ = currency.split("\n", maxsplit=1)
    _,currency = currency.split("(")
    currency = currency.replace(")", "")

    _, portfolios = text.split("PORTFOLIO INFORMATION\n", maxsplit=1)
    portfolios,data = portfolios.split("DUPLICATE COPIES OF THIS ACCOUNT STATEMENT ARE BEING SENT TO:\n", maxsplit=1)
    _,portfolios = portfolios.split("INDIVIDUAL PORTFOLIOS\n")
    portfolios = portfolios.split("\n")

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
                    if " " in portfolio_number:
                        portfolio_number,_ = portfolio_number.split(" ")
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
        # TODO: Agrega "CASH, DEPOSITS & MONEY MARKET FUNDS"

        del public_eq_client[0]
        del fixed_income_client[0]
        
        public_eq_result: list[str] = GenerateResultList(public_eq_client)
        fixed_income_result: list[str] = GenerateResultList(fixed_income_client)

        public_eq = GenerateResultProduct(public_eq_result, "PUBLIC EQUITY", date_document, client, portfolio_numbers[client], currency)
        fixed_income = GenerateResultProduct(fixed_income_result, "FIXED INCOME", date_document, client, portfolio_numbers[client], currency)

        results_public_eq += public_eq
        results_fixed_income += fixed_income

    df = pd.DataFrame(results_public_eq + results_fixed_income)
    df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == '__main__':
    import os

    files = os.listdir("./input")

    for file in files:
        if file == '.gitkeep':
            continue
        filename,_ = file.split('.')
        GoldmanParser(filename)
