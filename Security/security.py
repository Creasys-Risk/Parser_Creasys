import re
import datetime
import pandas as pd
from PyPDF2 import PdfReader
from pathlib import Path

def process_file_to_data(filepath: Path):
    info_cartera = []
    info_movimientos = []
    text = ""
    
    if filepath.suffix.lower() == ".pdf":
        reader = PdfReader(filepath)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    else:
        text = filepath.read_text(encoding="utf-8")

    fecha_cartera = ""
    m_fecha = re.search(r"Desde\s+el\s+(\d{2}-\d{2}-\d{4})\s+al\s+(\d{2}-\d{2}-\d{4})", text)
    if m_fecha:
        fecha_cartera = m_fecha.group(2)    
    nombre = ""
    m_nombre = re.search(r"Nombre\s*:\s*(.+?)\n", text)
    if m_nombre:
        nombre = m_nombre.group(1).strip()   
    rut = ""
    m_rut = re.search(r"Rut\s*:\s*([\d\.\-kK]+)", text)
    if m_rut:
        rut = m_rut.group(1).strip()
    clp_block_match = re.search(r"FONDOS SECURITY\s*-\s*CLP\s*(.*?)TOTAL", text, re.DOTALL)
    if clp_block_match:
        bloque_clp = clp_block_match.group(1)
        bloque_clp = re.sub(r"(?m)^(MENOR\s+VALOR|MAYOR\s+O\s+MENOR\s+VALOR)\s*$", "", bloque_clp)
        pattern_clp = re.compile(r"^(?!TOTAL)([A-Z0-9\s]+?)\s+S\s+(\d+)\s+(?:NO|SI)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)(?:\s+([\d\.,]+))?", re.MULTILINE)
        for match in pattern_clp.finditer(bloque_clp):
            nemotecnico = match.group(1).strip()
            cuenta = match.group(2).strip()
            cantidad = match.group(3).replace('.', '').replace(',', '.')
            precio_compra = match.group(4).replace('.', '').replace(',', '.')
            valor_compra = match.group(5).replace('.', '').replace(',', '.')
            precio_mercado = match.group(6).replace('.', '').replace(',', '.')
            valor_mercado = match.group(7).replace('.', '').replace(',', '.')
            info_cartera.append({
                "Fecha": fecha_cartera,
                "Nombre": nombre,
                "RUT": rut,
                "Cuenta": cuenta,
                "Nemotecnico": nemotecnico,
                "Moneda": "CLP",
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": cantidad,
                "Precio_Mercado": precio_mercado,
                "Valor_Mercado": valor_mercado,
                "Precio_Compra": precio_compra,
                "Valor_Compra": valor_compra,
                "interes_Acum": "",
                "Contraparte": "Security",
                "Clase_Activo": "Fondo Mutuo"
            })
    usd_block_match = re.search(r"FONDOS SECURITY\s*-\s*USD\s*(.*?)TOTAL", text, re.DOTALL)
    if usd_block_match:
        bloque_usd = usd_block_match.group(1)
        bloque_usd = re.sub(r"(?m)^(MENOR\s+VALOR|MAYOR\s+O\s+MENOR\s+VALOR)\s*$", "", bloque_usd)
        pattern_usd = re.compile(r"^(?!TOTAL)([A-Z0-9\s]+?)\s+S\s+(\d+)\s+(?:NO|SI)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)(?:\s+([\d\.,]+))?", re.MULTILINE)
        for match in pattern_usd.finditer(bloque_usd):
            nemotecnico = match.group(1).strip()
            cuenta = match.group(2).strip()
            cantidad = match.group(3).replace('.', '').replace(',', '.')
            precio_compra = match.group(4).replace('.', '').replace(',', '.')
            valor_compra = match.group(5).replace('.', '').replace(',', '.')
            precio_mercado = match.group(6).replace('.', '').replace(',', '.')
            valor_mercado = match.group(7).replace('.', '').replace(',', '.')
            info_cartera.append({
                "Fecha": fecha_cartera,
                "Nombre": nombre,
                "RUT": rut,
                "Cuenta": cuenta,
                "Nemotecnico": nemotecnico,
                "Moneda": "USD",
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": cantidad,
                "Precio_Mercado": precio_mercado,
                "Valor_Mercado": valor_mercado,
                "Precio_Compra": precio_compra,
                "Valor_Compra": valor_compra,
                "interes_Acum": "",
                "Contraparte": "Security",
                "Clase_Activo": "Fondo Mutuo"
            })
    acc_block_match = re.search(r"ACCIONES NACIONALES\s*(.*?)CUOTAS", text, re.DOTALL)
    if acc_block_match:
        bloque_acc = acc_block_match.group(1)
        for line in bloque_acc.splitlines():
            line = line.strip()
            if not line or line.upper().startswith("TOTAL"):
                continue
            if re.search(r"NOMBRE\s+ACCION", line):
                continue
            parts = re.split(r"\s+S\s+", line, maxsplit=1)
            if len(parts) != 2:
                continue
            nemotecnico = parts[0].strip()
            rest = parts[1].strip()
            tokens = re.split(r"\s+", rest)
            if len(tokens) < 12:
                continue
            cuenta = tokens[0]
            cantidad = tokens[1]
            precio_compra = tokens[5]
            moneda = tokens[6]
            valor_compra = tokens[7]
            precio_mercado = tokens[8]
            valor_mercado = tokens[10]
            info_cartera.append({
                "Fecha": fecha_cartera,
                "Nombre": nombre,
                "RUT": rut,
                "Cuenta": cuenta,
                "Nemotecnico": nemotecnico,
                "Moneda": moneda,
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": cantidad.replace('.', '').replace(',', '.'),
                "Precio_Mercado": precio_mercado.replace('.', '').replace(',', '.'),
                "Valor_Mercado": valor_mercado.replace('.', '').replace(',', '.'),
                "Precio_Compra": precio_compra.replace('.', '').replace(',', '.'),
                "Valor_Compra": valor_compra.replace('.', '').replace(',', '.'),
                "interes_Acum": "",
                "Contraparte": "Security",
                "Clase_Activo": "Acciones"
            })
    cuotas_block_match = re.search(r"NOMBRE\s+CUOTA\s+MANDATONUMERO\s*(.*?)(?=^TOTAL\s+\d)", text, re.DOTALL | re.MULTILINE)
    if cuotas_block_match:
        bloque_cuotas = cuotas_block_match.group(1)
        for line in bloque_cuotas.splitlines():
            line = line.strip()
            if not line or line.upper().startswith("TOTAL"):
                continue
            parts = re.split(r"\s+S\s+", line, maxsplit=1)
            if len(parts) != 2:
                continue
            nemotecnico_full = parts[0].strip()
            nemotecnico = nemotecnico_full.split(" - ")[0].strip()
            rest = parts[1].strip()
            tokens = re.split(r"\s+", rest)
            if len(tokens) < 13:
                continue
            cuenta = tokens[0]
            cantidad = tokens[1]
            precio_compra = tokens[5]
            moneda = tokens[6]
            valor_compra = tokens[7]
            precio_mercado = tokens[8]
            valor_mercado = tokens[11]
            info_cartera.append({
                "Fecha": fecha_cartera,
                "Nombre": nombre,
                "RUT": rut,
                "Cuenta": cuenta,
                "Nemotecnico": nemotecnico,
                "Moneda": moneda,
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": cantidad.replace('.', '').replace(',', '.'),
                "Precio_Mercado": precio_mercado.replace('.', '').replace(',', '.'),
                "Valor_Mercado": valor_mercado.replace('.', '').replace(',', '.'),
                "Precio_Compra": precio_compra.replace('.', '').replace(',', '.'),
                "Valor_Compra": valor_compra.replace('.', '').replace(',', '.'),
                "interes_Acum": "",
                "Contraparte": "Security",
                "Clase_Activo": "Fondos de inversion"
            })
    deuda_block_match = re.search(r"INSTRUMENTOS DE DEUDA NACIONALES\s*(.*?)TOTAL", text, re.DOTALL)
    if deuda_block_match:
        bloque_deuda = deuda_block_match.group(1)
        lines = bloque_deuda.splitlines()
        start_index = None
        for i, line in enumerate(lines):
            if "MERCADO" in line.upper():
                start_index = i
                break
        if start_index is not None:
            for line in lines[start_index+1:]:
                line = line.strip()
                if not line or line.upper().startswith("TOTAL"):
                    continue
                tokens = re.split(r"\s+", line)
                if "UF" not in tokens:
                    continue
                idx_uf = tokens.index("UF")
                if len(tokens) <= idx_uf + 7:
                    continue
                if len(tokens) < 2:
                    continue
                if tokens[1].upper() == "S":
                    nemotecnico = tokens[0]
                else:
                    nemotecnico = tokens[0] + " " + tokens[1]
                cuenta = ""
                for token in tokens[2:idx_uf]:
                    m_cuenta = re.match(r"^(\d+)", token)
                    if m_cuenta:
                        cuenta = m_cuenta.group(1)
                        break
                cantidad = tokens[idx_uf+1].replace('.', '').replace(',', '.')
                precio_compra = tokens[idx_uf+4].replace('.', '').replace(',', '.')
                valor_compra = tokens[idx_uf+5].replace('.', '').replace(',', '.')
                precio_mercado = tokens[idx_uf+6].replace('.', '').replace(',', '.')
                valor_mercado = tokens[idx_uf+7].replace('.', '').replace(',', '.')
                info_cartera.append({
                    "Fecha": fecha_cartera,
                    "Nombre": nombre,
                    "RUT": rut,
                    "Cuenta": cuenta,
                    "Nemotecnico": nemotecnico,
                    "Moneda": "UF",
                    "ISIN": "",
                    "CUSIP": "",
                    "Cantidad": cantidad,
                    "Precio_Mercado": precio_mercado,
                    "Valor_Mercado": valor_mercado,
                    "Precio_Compra": precio_compra,
                    "Valor_Compra": valor_compra,
                    "interes_Acum": "",
                    "Contraparte": "Security",
                    "Clase_Activo": "Instrumento de Deuda"
                })
    movs_match = re.search(r"INFORME DE TRANSACCIONES\s*(.*)", text, re.DOTALL)
    if movs_match:
        bloque_movs = movs_match.group(1)
        for line in bloque_movs.splitlines():
            line = line.strip()
            if not line:
                continue
            if not re.match(r"^\d{2}/\d{2}/\d{4}", line):
                continue
            tokens = line.split()
            if len(tokens) < 14:
                continue
            fecha_mov = tokens[0]
            fecha_liq = tokens[1]
            tipo_mov = tokens[3] + " " + tokens[4]
            cuenta = tokens[5]
            folio = tokens[6]
            nemotecnico = " ".join(tokens[7:10]).split("-")[0]
            cantidad_orig = tokens[10]
            cantidad = float(cantidad_orig.replace('.', '').replace(',', '.'))
            moneda_val = tokens[11]
            precio_orig = tokens[12]
            precio = float(precio_orig.replace('.', '').replace(',', '.'))
            monto_orig = tokens[13]
            monto = float(monto_orig.replace('.', '').replace(',', '.'))
            comision_orig = tokens[14] if len(tokens) > 14 else ""
            comision = float(comision_orig.replace('.', '').replace(',', '.')) if comision_orig else ""
            info_movimientos.append({
                "Fecha Movimiento": fecha_mov,
                "Fecha Liquidación": fecha_liq,
                "Nombre": nombre,
                "RUT": rut,
                "Cuenta": cuenta,
                "Nemotecnico": nemotecnico,
                "Moneda": moneda_val,
                "ISIN": "",
                "CUSIP": "",
                "Cantidad": cantidad,
                "Precio": precio,
                "Monto": monto,
                "Comision": comision,
                "tipo": tipo_mov,
                "descripcion": "",
                "Concepto": "",
                "Folio": folio,
                "Contraparte": "Security",
                "Cantidad_orig": cantidad_orig,
                "Precio_orig": precio_orig,
                "Monto_orig": monto_orig,
                "Comision_orig": comision_orig
            })
    return info_cartera, info_movimientos

def get_decimal_format(num_str):
    m = re.search(r",(\d+)$", num_str)
    if m:
        decimals = len(m.group(1))
        return "#,##0." + "0" * decimals
    else:
        return "#,##0"

def Security_Parser(input: Path, output: Path):
    all_cartera, all_movs = [], []
    
    for file in input.glob("*"):
        if file.is_file() and file.suffix.lower() in (".pdf", ".txt"):
            cartera, movs = process_file_to_data(file)
            all_cartera.extend(cartera)
            all_movs.extend(movs)
    
    if not all_cartera and not all_movs:
        return
    
    df_cartera = pd.DataFrame(all_cartera)
    df_movs = pd.DataFrame(all_movs)
    
    cartera_cols = ["Fecha","Nombre","RUT","Cuenta","Nemotecnico","Moneda","ISIN","CUSIP","Cantidad",
                    "Precio_Mercado","Valor_Mercado","Precio_Compra","Valor_Compra","interes_Acum",
                    "Contraparte","Clase_Activo"]
    df_cartera = df_cartera.reindex(columns=cartera_cols)
    
    movs_cols = ["Fecha Movimiento","Fecha Liquidación","Nombre","RUT","Cuenta","Nemotecnico","Moneda",
                 "ISIN","CUSIP","Cantidad","Precio","Monto","Comision","tipo","descripcion","Concepto",
                 "Folio","Contraparte"]
    df_movs = df_movs.reindex(columns=movs_cols)
    
    numeric_cols_cartera = ["Cantidad","Precio_Mercado","Valor_Mercado","Precio_Compra","Valor_Compra"]
    df_cartera[numeric_cols_cartera] = df_cartera[numeric_cols_cartera].apply(pd.to_numeric, errors="coerce")
    
    numeric_cols_movs = ["Cantidad","Precio","Monto","Comision"]
    df_movs[numeric_cols_movs] = df_movs[numeric_cols_movs].apply(pd.to_numeric, errors="coerce")
    
    fecha_archivo = ""
    if not df_cartera.empty:
        fecha_archivo = df_cartera.iloc[0]["Fecha"]
    elif not df_movs.empty:
        fecha_archivo = df_movs.iloc[0]["Fecha Movimiento"]
    
    try:
        parsed_date = datetime.datetime.strptime(fecha_archivo, "%d-%m-%Y")
        fecha_str = parsed_date.strftime("%Y%m%d")
    except:
        fecha_str = datetime.datetime.today().strftime("%Y%m%d")
    
    output.mkdir(parents=True, exist_ok=True)
    out_file = output / f"InformeSecurity_{fecha_str}.xlsx"
    
    with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
        if not df_cartera.empty:
            df_cartera.to_excel(writer, sheet_name="Cartera", index=False)       
        if not df_movs.empty:
            df_movs.to_excel(writer, sheet_name="Movimientos", index=False)
            workbook = writer.book
            worksheet_mov = writer.sheets["Movimientos"]
            
            num_fields = {
                "Cantidad": {"orig": "Cantidad_orig", "col": df_movs.columns.get_loc("Cantidad")},
                "Precio": {"orig": "Precio_orig", "col": df_movs.columns.get_loc("Precio")},
                "Monto": {"orig": "Monto_orig", "col": df_movs.columns.get_loc("Monto")},
                "Comision": {"orig": "Comision_orig", "col": df_movs.columns.get_loc("Comision")}
            }
            
            for row_num, (_, row) in enumerate(df_movs.iterrows(), start=1):
                for field, info in num_fields.items():
                    orig_value = row.get(info["orig"], "")
                    if isinstance(orig_value, str) and orig_value:
                        fmt_str = get_decimal_format(orig_value)
                        cell_format = workbook.add_format({'num_format': fmt_str})
                        worksheet_mov.write_number(row_num, info["col"], row[field], cell_format)

if __name__ == "__main__":
    Security_Parser(Path("Input"), Path("Output"))