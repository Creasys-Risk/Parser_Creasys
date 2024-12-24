from PyPDF2 import PdfReader
import pandas as pd

def JPM_Parser(filename: str):
    reader = PdfReader(f"./input/{filename}.pdf")

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    # with open("output_jp_morgan.txt", "w") as f:
    #     f.write(text)

    if "Equity Detail\n" not in text:
        print(f"No se encontró información de Equity en el archivo: {filename}.pdf")
        return 

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
            "mnemonic": nemotecnico,
            "quantity": quantity,
            "market_price": price,
            "market_value": value,
            "accrued_income": 0.0,
        })

    df = pd.DataFrame(results)
    df.to_excel(f"./output/{filename}.xlsx", index=False, engine="openpyxl")

if __name__ == "__main__":
    import os

    files = os.listdir("./input")

    for file in files:
        filename,_ = file.split('.')
        JPM_Parser(filename)
