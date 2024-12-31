import re
import pandas as pd
from PyPDF2 import PdfReader

def Extract_Data(rows: list[str]) -> list[dict]:
    results = []
    for row in rows:
        row_data = row.split(" ")
        row_data = [x for x in row_data if x != '' and x != ')']
        row_data = [x.replace('(', '-') for x in row_data]
        
        price = float(row_data[0].replace(',', ''))
        quantity = float(row_data[1].replace(',', ''))
        value = float(row_data[2].replace(',', ''))
        nemotecnico = f"{row_data[-2]} {row_data[-1]}"
        results.append({
            "mnemonic": nemotecnico,
            "quantity": quantity,
            "market_price": price,
            "market_value": value,
            "accrued_income": 0.0,
        })

    return results

def Extract_Equity(text: str, filename: str) -> list[dict]:
    if "Equity Detail\n" not in text:
        print(f"No se encontró información de Equity en el archivo: {filename}.pdf")
        return [{}]

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

    results = Extract_Data(equity_rows)

    return results

def Extract_Fixed_Income(text: str, filename: str) -> list[dict]:
    if "Cash & Fixed Income Detail" not in text:
        print(f"No se encontró información de Fixed Income en el archivo: {filename}.pdf")
        return [{}]
    
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
                _,line = line.split("Fixed Income")
            if re.match(r'^\s{9}', line):
                if "This is the Annual Percentage Yield" not in row and "For the Period" not in row and "Page" not in row and "Price Quantity Value Average Cost" not in row:
                    fixed_income_rows.append(row)
                row = ""
            row += f" {line}"
        fixed_income_rows.append(row)

    for row in fixed_income_rows:
        row_data = row.replace(f"{' '*6}", ";")
        row_data = row_data.split(";")
        row_data = [x for x in row_data if x != '']
        row_data = [x.replace('(', '-') for x in row_data]
        row_data = [x.replace(' )', '') for x in row_data]

        print(row_data)

    # results = Extract_Data(fixed_income_rows)

    # return results

def JPM_Parser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    with open("output_jp_morgan.txt", "w") as f:
        f.write(text)

    results_equity = Extract_Equity(text, filename)
    results_fixed_income = Extract_Fixed_Income(text, filename)

    df = pd.DataFrame(results_equity)
    df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == "__main__":
    import os

    files = os.listdir("./input")

    for file in files:
        if file == ".gitkeep":
            continue
        filename,_ = file.split('.')
        # JPM_Parser(filename)

    JPM_Parser("20160831 GB (JPM W22161009)")
