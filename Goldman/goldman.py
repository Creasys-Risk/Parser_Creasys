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

calendar_mini = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

def ProcessPurchasesSales(client_info: list[str], client: str, portfolio_number: str, currency: str):
    client_result: list[str] = []
    result_purchases_sales: list[dict] = []
    for client_data in client_info:
        if "ActivitySettlement" in client_data:
            continue
        aux_list = client_data.split("Accrued Interest")
        for aux_data in aux_list:
            if "Period Ended" in aux_data:
                aux_data,_ = aux_data.split("Period Ended")
            if "TOTAL " in aux_data:
                aux_data,_ = aux_data.split("TOTAL ", maxsplit=1)
            if "Type of Activity" in aux_data:
                continue

            aux = aux_data.split("\n")

            if re.search(r"\d", aux[-1]) and not re.search(r"[a-zA-Z]", aux[-1]):
                del aux[-1]
            del aux[0]
            
            row_data = ""
            for row in aux:
                char = row.split(" ")
                if char[0] in ["Purchase", "Sale"] and row_data != "":
                    client_result.append(row_data)
                    row_data = row
                else:
                    row_data += " " + row
            if row_data != "":
                client_result.append(row_data)

    for row in client_result:
        if row[0] == " ":
            row = row[1:]
        row_data = row.split(" ")

        activity = row_data.pop(0)
        if activity not in ["Purchase", "Sale"]:
            continue

        trade_month = calendar_mini[row_data.pop(0)]
        trade_day = int(row_data.pop(0))
        trade_year = int(f"20{row_data.pop(0)}")
        trade_date = date(trade_year, trade_month, trade_day)

        settlement_month = calendar_mini[row_data.pop(0)]
        settlement_day = int(row_data.pop(0))
        if len(row_data[0]) > 2:
            settlement_year = int(f"20{row_data[0][0:2]}")
            row_data[0] = row_data[0][2:]
            quantity = row_data.pop(0)
            if row_data[0] == "":
                del row_data[0]
            price = row_data.pop(0)

            if re.search(r"\d", row[0]):
                commission = row_data.pop(0)
            else:
                commission = "0"
            fees = "0"
            for i, d in enumerate(row_data):
                if re.search(r"\d", d):
                    row_data = row_data[i:]
                    break
            amount = row_data.pop(0)
        else:
            settlement_year = int(f"20{row_data.pop(0)}")
            quantity = row_data.pop(0)
            price = row_data.pop(0)
            commission = "0"
            if re.search(r"[a-zA-Z]",row_data[1]):
                fees = "0"
                amount = row_data.pop(0)
            else:
                fees = row_data.pop(0)
                amount = row_data.pop(0)
        settlement_date = date(settlement_year, settlement_month, settlement_day)

        description = " ".join(row_data)
        quantity = float(quantity.replace(',', '').replace("(", "-").replace(")", ""))
        price = float(price.replace(',', '').replace("(", "-").replace(")", ""))
        commission = float(commission.replace(',', '').replace("(", "-").replace(")", ""))
        fees = float(fees.replace(',', '').replace("(", "-").replace(")", ""))
        amount = float(amount.replace(',', '').replace("(", "-").replace(")", ""))

        result_purchases_sales.append({
            "Fecha Movimiento": trade_date,
            "Fecha Liquidación": settlement_date,
            "Nombre": client,
            "RUT": "",
            "Cuenta": portfolio_number,
            "Nemotecnico": "",
            "Moneda": currency,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": quantity,
            "Precio": price,
            "Monto": amount,
            "Comision": commission,
            "IVA": 0,
            "Tipo": activity,
            "Descripcion": description,
            "Concepto": "OPERACIÓN",
            "Folio": "",
            "Contraparte": "GOLDMAN SACHS",
        })

    return result_purchases_sales

def ProcessTransactions(client_info: list[str], client: str, portfolio_number: str, currency: str):
    page_split = False
    activity_list = ['Deposit', 'Intr', 'Fee', 'Credi', 'Wire']
    client_result: list[str] = []
    result_transaction = []
    for client_data in client_info:
        aux_currency = client_data.split("\n")
        if "(" in aux_currency[1] and not re.search(r"\d", aux_currency[1]):
            _,currency = aux_currency[1].split("(")
            currency,_ = currency.split(")")
        if client_data.count("CLOSING BALANCE AS OF") == 1:
            if page_split:
                client_data,_ = client_data.split("CLOSING BALANCE AS OF")
                page_split = False
            else:
                _,client_data = client_data.split("CLOSING BALANCE AS OF")
                client_data,_ = client_data.split("Cash ActivityStatement Detail")
                page_split = True
        elif client_data.count("CLOSING BALANCE AS OF") == 2:
            _,client_data,_ = client_data.split("CLOSING BALANCE AS OF", maxsplit=2)
        client_list = client_data.split("\n")
        del client_list[0]
        del client_list[-1]

        for row in client_list:
            aux = row.split(" ")
            if aux[0] in activity_list:
                client_result.append(row)
            elif len(client_result) > 0:
                client_result[-1] += " " + row

    for row in client_result:
        aux = row.split(" ")
        if aux[0] == 'Intr':
            if aux[1] == 'Chgd':
                activity = f"{aux[0]} {aux[1]}"
                aux = aux[2:]
            else:
                activity = f"{aux[0]} {aux[1]}{aux[2]}"
                aux = aux[3:]
        elif aux[0] == 'Credi':
            activity = f"{aux[0]}{aux[1]}"
            aux = aux[2:]
        else:
            activity = aux[0]
            aux = aux[1:]
        
        month = calendar_mini[aux.pop(0)]
        day = int(aux.pop(0))
        year = int(f"20{aux.pop(0)}")
        effective_date = date(year, month, day)
        amount = float(aux.pop(0).replace(',', '').replace('(', '-').replace(')', ''))

        if re.search(r"\d", aux[0]) and "%" not in aux[0]:
            money_balance = float(aux.pop(0).replace(',', '').replace('(', '-').replace(')', ''))
        else:
            money_balance = ''

        nemo = " ".join(aux)

        result_transaction.append({
            "Fecha Movimiento": effective_date,
            "Fecha Liquidación": effective_date,
            "Nombre": client,
            "RUT": "",
            "Cuenta": portfolio_number,
            "Nemotecnico": "",
            "Moneda": currency,
            "ISIN": "",
            "CUSIP": "",
            "Cantidad": "",
            "Precio": "",
            "Monto": amount,
            "Comision": 0,
            "IVA": 0,
            "Tipo": activity,
            "Descripcion": nemo,
            "Concepto": "MOVIMIENTO",
            "Folio": "",
            "Contraparte": "GOLDMAN SACHS",
        })

    return result_transaction

def GenerateResultList(client_info: list[str]) -> list[str]:
    client_result = []
    for client_data in client_info:
        if "TOTAL PORTFOLIO" in client_data:
            client_data,_ = client_data.split("TOTAL", maxsplit=1)

        if "FIXED INCOME" in client_data:
            client_data,_ = client_data.split("FIXED INCOME", maxsplit=1)

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
            if (re.match(r"\d", aux[0]) or (aux[0] == '(' and re.match(r"\d", aux[1]))) and "%" not in aux:
                client_result.append(aux)
            elif re.search(r"[a-zA-Z]", aux) and re.search(r"\d", aux) and len(client_result) > 0:
                if re.search(r'\([A-Z]*\)', client_result[-1]) and re.search(r'\([A-Z]*\)', aux):
                    continue
                client_result[-1] += " " + aux

    client_result_concat = []

    for index, data in enumerate(client_result):
        if not re.search(r"[a-zA-Z]", data) and index+1 < len(client_result):
            client_result[index+1] = data + " " + client_result[index+1]
        elif not " " in data and len(client_result_concat) > 0:
            client_result_concat[-1] += " " + data
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

    with open("output_Goldman_Sachs.txt", "w") as f:
        f.write(text)

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
                    portfolio_numbers[client_name] = portfolio_number

    for k in portfolio_numbers:
        if " " in portfolio_numbers[k]:
            portfolio_numbers[k],_ = portfolio_numbers[k].split(" ")
    
    results_public_eq = []
    results_fixed_income = []
    results_cash_deposits = []
    results_movement = []

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
        cash_deposits_client: list[str] = client_text.split("CASH, DEPOSITS & MONEY MARKET FUNDS")
        transactions_client: list[str] = client_text.split("TRANSACTIONS AFFECTING CASH")
        purchases_sales_client: list[str] = client_text.split("PURCHASES & SALES")

        del public_eq_client[0]
        del fixed_income_client[0]
        del cash_deposits_client[0]
        del transactions_client[0]
        del purchases_sales_client[0]
        
        public_eq_result: list[str] = GenerateResultList(public_eq_client)
        fixed_income_result: list[str] = GenerateResultList(fixed_income_client)
        cash_deposits_result: list[str] = GenerateResultList(cash_deposits_client)

        public_eq = GenerateResultProduct(public_eq_result, "PUBLIC EQUITY", date_document, client, portfolio_numbers[client], currency)
        fixed_income = GenerateResultProduct(fixed_income_result, "FIXED INCOME", date_document, client, portfolio_numbers[client], currency)
        cash_deposits = GenerateResultProduct(cash_deposits_result, "CASH, DEPOSITS & MONEY MARKET FUNDS", date_document, client, portfolio_numbers[client], currency)

        transactions = ProcessTransactions(transactions_client, client, portfolio_numbers[client], currency)
        purchases_sales = ProcessPurchasesSales(purchases_sales_client, client, portfolio_numbers[client], currency)

        results_public_eq += public_eq
        results_fixed_income += fixed_income
        results_cash_deposits += cash_deposits
        
        results_movement += transactions
        results_movement += purchases_sales

    df_cartera = pd.DataFrame(results_public_eq + results_fixed_income + results_cash_deposits)
    df_movimientos = pd.DataFrame(results_movement)
    with pd.ExcelWriter(f"./output/{filename}.xlsx", engine="openpyxl") as writer:
        df_cartera.to_excel(writer, index=False, sheet_name="Cartera")
        df_movimientos.to_excel(writer, index=False, sheet_name="Movimientos")

if __name__ == '__main__':
    import os

    files = os.listdir("./input")

    for file in files:
        if file == '.gitkeep':
            continue
        filename,_ = file.split('.')
        GoldmanParser(filename)
