import os
import re
import datetime
import pandas as pd
from decimal import Decimal, InvalidOperation
from PyPDF2 import PdfReader

FOLDER_INPUT = "./Input/"
FOLDER_OUTPUT = "./Output/"
FOLDER_TXT = os.path.join(FOLDER_OUTPUT, "Textos")
os.makedirs(FOLDER_OUTPUT, exist_ok=True)
os.makedirs(FOLDER_TXT, exist_ok=True)

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def parse_number(num_str):
    if not num_str:
        return None
    s = num_str.replace("%", "").replace("$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(Decimal(s))
    except InvalidOperation:
        return None

def unify_spaces(s):
    return re.sub(r"\s+", " ", s).strip()

def get_fecha_global(text):
    m = re.search(r"al\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def convert_fecha(fecha_str):
    meses = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
             "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}
    parts = fecha_str.upper().split(" DE ")
    if len(parts) == 3:
        try:
            return datetime.date(
                int(parts[2].strip()),
                meses.get(parts[1].strip(), 0),
                int(parts[0].strip())
            )
        except:
            return None
    return None

def get_metadata_block(text):
    fecha_global = get_fecha_global(text)
    fecha_obj = convert_fecha(fecha_global) if fecha_global else None
    mr = re.search(r"(\d{1,2}\.\d{3}\.\d{3}-[\dkK])", text)
    rut = mr.group(1) if mr else ""
    nombre = ""
    cuenta = ""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        l = line.strip()
        if re.search(r"Servicio al Cliente", l, re.IGNORECASE):
            if i + 1 < len(lines):
                cand = lines[i+1].strip()
                if cand and not re.match(r"^(Atención a|Dirección|RUT|N° Cuenta)", cand, re.IGNORECASE):
                    nombre = cand
        if re.match(r"^Nombre\s*$", l, re.IGNORECASE):
            if i + 1 < len(lines):
                cand2 = lines[i+1].strip()
                if cand2 and not re.match(r"^(Atención a|Dirección|RUT|N° Cuenta)", cand2, re.IGNORECASE):
                    nombre = cand2
        if re.match(r"^Cuenta", l, re.IGNORECASE):
            if i + 1 < len(lines):
                c_line = lines[i+1].strip()
                if re.match(r"^\d+$", c_line):
                    cuenta = c_line
        if not cuenta:
            mc = re.search(r"N°\s*Cuenta\s+(\d+)", l, re.IGNORECASE)
            if mc:
                cuenta = mc.group(1).strip()
    try:
        cuenta = float(cuenta) if cuenta != "" else ""
    except:
        cuenta = ""
    return fecha_obj, nombre.strip(), rut.strip(), cuenta

def fix_section_headers(text):
    patterns = [
        r"Inversiones En Fondos De Inversión\s+Locales En CLP",
        r"Fondo Serie\s+Nro\. Cuotas\s+Precio Compra Precio \(\*\)\s*Valorización\s*%\s*Cartera\s*Inversiones\s*En\s*Fondos\s*De Inversión\s+Locales En CLP",
        r"Inversiones En Fondos Mutuos\s+Locales En CLP",
        r"Inversiones En Fondos Mutuos\s+Locales En USD",
        r"Inversiones En Renta Fija\s+Locales En CLP",
        r"Inversiones En Renta Fija\s+Internacional En USD",
        r"Inversiones En Fondos Mutuos\s+Internacionales En USD",
        r"Inversiones En Acciones\s+Locales En CLP",
        r"Total Inversiones En Fondos De Inversión\s+Locales En CLP",
        r"Total Inversiones En Fondos Mutuos\s+Locales En CLP"
    ]
    for pat in patterns:
        text = re.sub(pat, lambda m: "\n" + m.group(0), text, flags=re.IGNORECASE)
    text = re.sub(r"(Inversiones En Fondos Mutuos Locales En (CLP|USD))(?=\d)", r"\1\n", text, flags=re.IGNORECASE)
    return text

def detect_currency_dynamic(line: str, default_moneda: str) -> str:
    up_line = line.upper()
    if "USD" in up_line:
        return "USD"
    if "UF" in up_line and "$" in up_line:
        if re.search(r"\d+,\d+%\s*", up_line):
            return "UF"
        else:
            return "CLP"
    if "UF" in up_line:
        return "UF"
    if "$" in up_line:
        return "CLP"
    return default_moneda

def parse_fondos_inversion_local_clp(line):
    p = re.compile(
        r"^(?P<nemotecnico>.+?)\s+"
        r"(?P<cantidad>[\d\.,]+)\s+\$\s*(?P<precio_compra>[\d\.,]+)\s+\$\s*(?P<precio_mercado>[\d\.,]+)\s+\$\s*(?P<valor_mercado>[\d\.,]+).*"
    )
    m = p.match(line.strip())
    if not m:
        return None
    record = {
        "Nemotecnico": unify_spaces(m.group("nemotecnico")),
        "Cantidad": m.group("cantidad"),
        "Precio_Compra": m.group("precio_compra"),
        "Precio_Mercado": m.group("precio_mercado"),
        "Valor_Mercado": m.group("valor_mercado")
    }
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record[campo]) is None:
            return None
    return record

def parse_fondos_mutuos_locales_clp(line):
    fm_idx = line.find("FM ")
    if fm_idx == -1:
        return None
    sub = line[fm_idx:].strip()
    p = re.compile(
        r"^(?P<nemotecnico>FM\s+.+?)(?:\s+(?P<serie>[A-Z]))?\s+"
        r"(?P<cantidad>[\d\.,]+)\s+\$\s*"
        r"(?P<precio_compra>[\d\.,]+)\s+\$\s*"
        r"(?P<precio_mercado>[\d\.,]+)\s+"
        r"(?P<porcentaje>[\d\.,]+)\s*%\s+\$\s*"
        r"(?P<valor_mercado>[\d\.,]+)"
    )
    m = p.match(sub)
    if not m:
        return None
    nem = m.group("nemotecnico").strip()
    serie = m.group("serie")
    if serie:
        nem = nem + " " + serie
    record = {
        "Nemotecnico": nem,
        "Cantidad": m.group("cantidad"),
        "Precio_Compra": m.group("precio_compra"),
        "Precio_Mercado": m.group("precio_mercado"),
        "Valor_Mercado": m.group("valor_mercado")
    }
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record[campo]) is None:
            return None
    return record

def parse_fmius_line_regex(line):
    pattern = re.compile(
        r"^(?P<nemotecnico>.+?)\s+"
        r"(?P<p1>[\d,]+%)\s+"
        r"(?P<valor_mercado>[\d\.,]+)\s+"
        r"(?P<p2>[\d,]+%)\s+"
        r"(?P<cantidad>[\d\.,]+)\s+"
        r"(?P<precio_compra>[\d\.,]+)\s+"
        r"(?P<precio_mercado>[\d\.,]+)\s+"
        r"(?P<custodio>\S+)$"
    )
    m = pattern.match(line.strip())
    if not m:
        return None
    record = {
        "Nemotecnico": m.group("nemotecnico").strip(),
        "Valor_Mercado": m.group("valor_mercado").strip(),
        "Cantidad": m.group("cantidad").strip(),
        "Precio_Compra": m.group("precio_compra").strip(),
        "Precio_Mercado": m.group("precio_mercado").strip()
    }
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record.get(campo, "")) is None:
            return None
    return record

def parse_fondos_mutuos_internacionales_usd_block(lines, start_idx):
    data = []
    i = start_idx + 1
    while i < len(lines):
        line = lines[i].strip()
        if not line or re.search(r"^Total Inversiones En Fondos Mutuos\s+Internacionales En USD", line, re.IGNORECASE):
            break
        parsed = parse_fmius_line_regex(line)
        if parsed:
            parsed["Moneda"] = "USD"
            parsed["Clase_Activo"] = "Fondos Mutuos Internacionales"
            data.append(parsed)
        i += 1
    return data, i

def parse_renta_fija_internacional_usd_new(line):
    tokens = line.split()
    if not tokens or tokens[0].lower()=="total" or tokens[0].replace(",","").replace(".","").isdigit():
        return None
    try:
        usd_index = tokens.index("USD")
        instrument = tokens[0]
        cantidad = tokens[usd_index+2]
        precio_compra = tokens[usd_index+3]
        precio_mercado = tokens[usd_index+4]
        valor_mercado = tokens[usd_index+6]
    except:
        return None
    if any(parse_number(x) is None for x in [cantidad,precio_compra,precio_mercado,valor_mercado]):
        return None
    return {"Nemotecnico":instrument,"Cantidad":cantidad,"Precio_Compra":precio_compra,"Precio_Mercado":precio_mercado,"Valor_Mercado":valor_mercado}

def parse_renta_fija_internacional_usd_block(lines, start_idx):
    data=[]
    i=start_idx
    while i<len(lines):
        line=lines[i].strip()
        if not line or is_new_section_header(line):
            break
        rec=parse_renta_fija_internacional_usd_new(line)
        if rec:
            rec["Moneda"]="USD"
            rec["Clase_Activo"]="Renta Fija  Internacional"
            data.append(rec)
        i+=1
    return data,i

def parse_fondos_mutuos_locales_usd(line):
    pattern=re.compile(
        r"^(?:[\d,]+\s*%)\s+"
        r"(?P<nemotecnico>FM\s+.+?)(?:\s+(?P<serie>[A-Z]))?\s+"
        r"(?P<cantidad>[\d\.,]+)\s+(?:[\d,]+\s*%)\s+"
        r"(?P<valor_mkt>[\d\.,]+)\s+USD\s+"
        r"(?P<precio_compra>[\d\.,]+)\s+USD\s+"
        r"(?P<precio_mercado>[\d\.,]+)\s+USD"
    )
    m=pattern.search(line)
    if not m:
        return None
    nem=m.group("nemotecnico").strip()
    serie=m.group("serie")
    if serie:
        nem=nem+" "+serie
    record={"Nemotecnico":nem,"Cantidad":m.group("cantidad").strip(),"Valor_Mercado":m.group("valor_mkt").strip(),"Precio_Compra":m.group("precio_compra").strip(),"Precio_Mercado":m.group("precio_mercado").strip()}
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record[campo]) is None:
            return None
    return record

def parse_fondos_mutuos_usd_block(lines, start_idx):
    data=[]
    i=start_idx
    while i<len(lines):
        line=lines[i].strip()
        if not line or is_new_section_header(line):
            break
        rec=parse_fondos_mutuos_locales_usd(line)
        if rec:
            rec["Moneda"]="USD"
            rec["Clase_Activo"]="Fondos Mutuos"
            data.append(rec)
        i+=1
    return data,i

def parse_renta_fija_local_clp(line):
    p=re.compile(
        r"^(?P<nemo>\S+)\s+"
        r"(?P<emisor>.+?)\s+"
        r"(?P<rating>\S+)\s+"
        r"(?P<moneda>UF|CLP)\s+"
        r"(?P<f_vencimiento>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<cantidad>[\d\.,]+)\s+"
        r"(?P<precio_compra>[\d\.,]+)%\s+"
        r"(?P<precio_mercado>[\d\.,-]+)%\s+[\d\.,]+\s*%\s+\$\s+"
        r"(?P<valor_mercado>[\d\.,]+)"
    )
    m=p.match(line.strip())
    if not m:
        return None
    record={
        "Nemotecnico":m.group("nemo"),
        "Moneda":m.group("moneda"),
        "Cantidad":m.group("cantidad"),
        "Precio_Compra":m.group("precio_compra"),
        "Precio_Mercado":m.group("precio_mercado"),
        "Valor_Mercado":m.group("valor_mercado")
    }
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record[campo]) is None:
            return None
    return record

def parse_acciones_local_clp(line):
    if any(line.strip().startswith(prefix) for prefix in ("Total","Fondo","FI","Detalle","Instrumento")):
        return None
    p=re.compile(
        r"^(?P<nem>\S+)\s+"
        r"(?P<cant>[\d\.,]+)\s+\$\s*(?P<pcompra>[\d\.,]+)\s+\$\s*(?P<pmercado>[\d\.,]+)\s+"
        r"(?P<varprecio>[-\d\.,]+)\s*%\s+\$\s*(?P<valormercado>[\d\.,]+)"
    )
    m=p.search(line.strip())
    if not m:
        return None
    record={
        "Nemotecnico":m.group("nem"),
        "Cantidad":m.group("cant"),
        "Precio_Compra":m.group("pcompra"),
        "Precio_Mercado":m.group("pmercado"),
        "Valor_Mercado":m.group("valormercado")
    }
    for campo in ["Cantidad","Precio_Compra","Precio_Mercado","Valor_Mercado"]:
        if parse_number(record[campo]) is None:
            return None
    return record

section_patterns=[
    r"(?:Inversiones En|Fondo Serie\s+Nro\. Cuotas).*Fondos De Inversión\s+Locales En CLP",
    r"Inversiones En Fondos Mutuos\s+Locales En CLP",
    r"Inversiones En Fondos Mutuos\s+Internacionales En USD",
    r"Inversiones En Renta Fija\s+Internacional\s+En\s+USD",
    r"Inversiones En Renta Fija\s+Locales En CLP",
    r"Inversiones En Acciones\s+Locales En CLP",
    r"Inversiones En Fondos Mutuos Locales En USD"
]

def is_new_section_header(line):
    return any(re.search(pat, line, re.IGNORECASE) for pat in section_patterns)

def parse_section(lines,start_idx,section_title,total_title,func,clase,default_moneda):
    data=[]
    i=start_idx
    while i<len(lines):
        line=lines[i].strip()
        if re.search(total_title,line,re.IGNORECASE) or is_new_section_header(line) or re.search(r"^(Nombre|Detalle De Movimientos)",line,re.IGNORECASE):
            break
        rec=func(line)
        if rec:
            rec["Clase_Activo"]=clase
            if "Moneda" in rec and rec["Moneda"]:
                rec["Moneda"]=detect_currency_dynamic(line,rec["Moneda"])
            else:
                rec["Moneda"]=detect_currency_dynamic(line,default_moneda)
            data.append(rec)
        i+=1
    return data,i

def parse_all_investment_sections(lines):
    sections=[
        {"section_title":r"Inversiones En Fondos De Inversión\s+Locales En CLP","total_title":r"^Total Inversiones En Fondos De Inversión\s+Locales En CLP","func":parse_fondos_inversion_local_clp,"clase":"Fondos De Inversión","moneda":"CLP"},
        {"section_title":r"Inversiones En Fondos Mutuos\s+Locales En CLP","total_title":r"^Total Inversiones En Fondos Mutuos\s+Locales En CLP","func":parse_fondos_mutuos_locales_clp,"clase":"Fondos Mutuos","moneda":"CLP"},
        {"section_title":r"Inversiones En Fondos Mutuos\s+Internacionales En USD","total_title":r"^Total Inversiones En Fondos Mutuos\s+Internacionales En USD","func":parse_fondos_mutuos_internacionales_usd_block,"clase":"Fondos Mutuos Internacionales","moneda":"USD","block":True},
        {"section_title":r"Inversiones En Renta Fija\s+Internacional\s+En\s+USD","total_title":r"^Instrumento\s+Emisor\s+Rating\s+Moneda\s+F\.Vencimiento","func":parse_renta_fija_internacional_usd_block,"clase":"Renta Fija  Internacional","moneda":"USD","block":True},
        {"section_title":r"Inversiones En Renta Fija\s+Locales En CLP","total_title":r"^Total Inversiones En Renta Fija\s+Locales En CLP","func":parse_renta_fija_local_clp,"clase":"Renta Fija","moneda":"UF"},
        {"section_title":r"Inversiones En Acciones\s+Locales En CLP","total_title":r"^Total Inversiones En Acciones\s+Locales En CLP","func":parse_acciones_local_clp,"clase":"Acciones","moneda":"CLP"},
        {"section_title":r"Inversiones En Fondos Mutuos Locales En USD","total_title":r"^Total Inversiones En Fondos Mutuos Locales En USD","func":parse_fondos_mutuos_usd_block,"clase":"Fondos Mutuos","moneda":"USD","block":True}
    ]
    all_invs=[]
    i=0
    n=len(lines)
    while i<n:
        line=lines[i].strip()
        matched=False
        for sec in sections:
            if re.search(sec["section_title"],line,re.IGNORECASE):
                if sec.get("block"):
                    block_data,new_i=sec["func"](lines,i+1)
                else:
                    block_data,new_i=parse_section(lines,i+1,sec["section_title"],sec["total_title"],sec["func"],sec["clase"],sec["moneda"])
                all_invs.extend(block_data)
                i=new_i
                matched=True
                break
        if not matched:
            i+=1
    return all_invs

def parse_movements_in_block(lines,start_idx,nombre,rut,cuenta):
    movs=[]
    i=start_idx
    usd_regex=re.compile(
        r"^(?P<fecha_mov>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<tipo>\w+)\s+"
        r"(?P<instrumento>(?:\S+\s+){3}\S+)\s+"
        r"(?P<doc>\S+)\s+"
        r"(?P<fecha_liq>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<monto>[\d\.,-]+)\s+"
        r"(?P<cantidad>[\d\.,-]+)\s+"
        r"(?P<precio>[\d\.,-]+)(?:\s+.*)?$",
        re.IGNORECASE
    )
    clp_regex=re.compile(
        r"^(?P<fecha_mov>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<tipo>\w+)\s+"
        r"(?P<instrumento>.+?)\s+"
        r"(?P<doc>\S+)\s+"
        r"(?P<fecha_liq>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<cantidad>[\d\.,-]+)\s+"
        r"(?P<precio>[\d\.,-]+)\s+\$\s*\S+\s+\$\s*(?P<monto>[+-]?[\d\.,-]+)(?:\s+.*)?$",
        re.IGNORECASE
    )
    while i<len(lines):
        l=lines[i].strip()
        if re.search(r"^(Cartola de Movimientos|Nombre|Detalle De Movimientos|Inversiones En |Ganancias / Pérdidas)",l,re.IGNORECASE):
            break
        reg=usd_regex if "USD" in l.upper() else clp_regex
        m=reg.match(l)
        if m:
            monto_val=parse_number(m.group("monto"))
            uds_val=parse_number(m.group("cantidad"))
            px_val=parse_number(m.group("precio"))
            if uds_val is None or px_val is None or monto_val is None:
                i+=1
                continue
            movs.append({
                "Fecha Movimiento":m.group("fecha_mov"),
                "Fecha Liquidación":m.group("fecha_liq"),
                "Nombre":nombre,
                "RUT":rut,
                "Cuenta":cuenta,
                "Nemotecnico":unify_spaces(m.group("instrumento")),
                "Moneda":"USD" if "USD" in l.upper() else "CLP",
                "ISIN":"",
                "CUSIP":"",
                "Cantidad":uds_val,
                "Precio":px_val,
                "Monto":monto_val,
                "Comision":"",
                "tipo":m.group("tipo"),
                "descripcion":"",
                "Concepto":"",
                "Folio":m.group("doc"),
                "Contraparte":"BTG pactual"
            })
        i+=1
    return movs,i

def process_file_to_data(file_path):
    if file_path.lower().endswith(".pdf"):
        text=extract_text_from_pdf(file_path)
    else:
        with open(file_path,"r",encoding="utf-8") as f:
            text=f.read()
    text=fix_section_headers(text)
    base_name=os.path.basename(file_path)
    dbg_txt_path=os.path.join(FOLDER_TXT,base_name+".txt")
    with open(dbg_txt_path,"w",encoding="utf-8") as dbg:
        dbg.write(text)
    fecha_obj,nombre_global,rut_global,cuenta_global=get_metadata_block(text)
    if not fecha_obj:
        fecha_obj=""
    lines=text.splitlines()
    inversiones=parse_all_investment_sections(lines)
    cartera=[]
    for inv in inversiones:
        registro={}
        registro["Fecha"]=fecha_obj
        registro["Nombre"]=re.sub(r"\.$","",re.sub(r"^Cuenta","",nombre_global).strip())
        registro["RUT"]=rut_global
        registro["Cuenta"]=cuenta_global
        registro["Nemotecnico"]=inv.get("Nemotecnico","")
        registro["Moneda"]=inv.get("Moneda","CLP")
        registro["ISIN"]=""
        registro["CUSIP"]=""
        registro["Cantidad"]=parse_number(inv.get("Cantidad",""))
        registro["Precio_Mercado"]=parse_number(inv.get("Precio_Mercado",""))
        registro["Valor_Mercado"]=parse_number(inv.get("Valor_Mercado",""))
        registro["Precio_Compra"]=parse_number(inv.get("Precio_Compra",""))
        if registro["Moneda"].upper() in ["CLP","USD"]:
            if registro["Cantidad"] is not None and registro["Precio_Compra"] is not None:
                registro["Valor_Compra"]=registro["Cantidad"]*registro["Precio_Compra"]
            else:
                registro["Valor_Compra"]=""
        else:
            registro["Valor_Compra"]=""
        registro["interes_Acum"]=""
        registro["Contraparte"]="BTG pactual"
        registro["Clase_Activo"]=inv.get("Clase_Activo","")
        cartera.append(registro)
    movs=[]
    i=0
    while i<len(lines):
        l=lines[i].strip()
        if re.match(r"^Cartola de Movimientos",l,re.IGNORECASE):
            parsed,new_i=parse_movements_in_block(lines,i+1,registro["Nombre"],rut_global,cuenta_global)
            movs.extend(parsed)
            i=new_i
        else:
            i+=1
    return cartera,movs

def main():
    all_cartera=[]
    all_movs=[]
    for fname in os.listdir(FOLDER_INPUT):
        if fname.lower().endswith((".pdf",".txt")):
            fpath=os.path.join(FOLDER_INPUT,fname)
            c,m=process_file_to_data(fpath)
            all_cartera.extend(c)
            all_movs.extend(m)
    if not all_cartera and not all_movs:
        return
    df_cartera=pd.DataFrame(all_cartera)
    df_movs=pd.DataFrame(all_movs)
    cartera_cols=["Fecha","Nombre","RUT","Cuenta","Nemotecnico","Moneda","ISIN","CUSIP","Cantidad","Precio_Mercado","Valor_Mercado","Precio_Compra","Valor_Compra","interes_Acum","Contraparte","Clase_Activo"]
    for col in cartera_cols:
        if col not in df_cartera.columns:
            df_cartera[col]=""
    df_cartera=df_cartera[cartera_cols]
    movs_cols=["Fecha Movimiento","Fecha Liquidación","Nombre","RUT","Cuenta","Nemotecnico","Moneda","ISIN","CUSIP","Cantidad","Precio","Monto","Comision","tipo","descripcion","Concepto","Folio","Contraparte"]
    for col in movs_cols:
        if col not in df_movs.columns:
            df_movs[col]=""
    df_movs=df_movs[movs_cols]
    numeric_cols_cartera=["Cantidad","Precio_Mercado","Valor_Mercado","Precio_Compra"]
    for col in numeric_cols_cartera:
        df_cartera[col]=pd.to_numeric(df_cartera[col],errors="coerce")
    numeric_cols_movs=["Cantidad","Precio","Monto"]
    for col in numeric_cols_movs:
        df_movs[col]=pd.to_numeric(df_movs[col],errors="coerce")
    if not df_cartera.empty and df_cartera["Fecha"].iloc[0]!="":
        if isinstance(df_cartera["Fecha"].iloc[0],datetime.date):
            df_cartera["Fecha"]=pd.to_datetime(df_cartera["Fecha"])
        file_date=df_cartera["Fecha"].iloc[0].strftime("%Y%m%d")
    else:
        file_date=datetime.datetime.today().strftime("%Y%m%d")
    out_file=os.path.join(FOLDER_OUTPUT,f"InformeBTG_{file_date}.xlsx")
    with pd.ExcelWriter(out_file,engine="xlsxwriter",date_format="dd/mm/yyyy",datetime_format="dd/mm/yyyy") as writer:
        if not df_cartera.empty:
            df_cartera.to_excel(writer,sheet_name="Cartera",index=False)
            workbook=writer.book
            worksheet=writer.sheets["Cartera"]
            date_format=workbook.add_format({"num_format":"dd/mm/yyyy"})
            worksheet.set_column(df_cartera.columns.get_loc("Fecha"),df_cartera.columns.get_loc("Fecha"),None,date_format)
            worksheet.set_column(df_cartera.columns.get_loc("Cantidad"),df_cartera.columns.get_loc("Cantidad"),None,workbook.add_format({"num_format":"0.0000"}))
            worksheet.set_column(df_cartera.columns.get_loc("Precio_Mercado"),df_cartera.columns.get_loc("Precio_Mercado"),None,workbook.add_format({"num_format":"0.0000"}))
            worksheet.set_column(df_cartera.columns.get_loc("Precio_Compra"),df_cartera.columns.get_loc("Precio_Compra"),None,workbook.add_format({"num_format":"0.0000"}))
            usd_format_valormercado=workbook.add_format({"num_format":"0.00"})
            clp_format_valormercado=workbook.add_format({"num_format":"0"})
            moneda_col_c=df_cartera.columns.get_loc("Moneda")
            valor_col_c=df_cartera.columns.get_loc("Valor_Mercado")
            for row_num,row_data in enumerate(df_cartera.itertuples(index=False),start=1):
                currency=row_data[moneda_col_c]
                valor=row_data[valor_col_c]
                if pd.notnull(valor):
                    if str(currency).upper()=="USD":
                        worksheet.write_number(row_num,valor_col_c,valor,usd_format_valormercado)
                    else:
                        worksheet.write_number(row_num,valor_col_c,valor,clp_format_valormercado)
        if not df_movs.empty:
            df_movs.to_excel(writer,sheet_name="Movimientos",index=False)
            workbook=writer.book
            worksheet=writer.sheets["Movimientos"]
            worksheet.set_column(df_movs.columns.get_loc("Cantidad"),df_movs.columns.get_loc("Cantidad"),None,workbook.add_format({"num_format":"0.0000"}))
            worksheet.set_column(df_movs.columns.get_loc("Precio"),df_movs.columns.get_loc("Precio"),None,workbook.add_format({"num_format":"0.0000"}))
            usd_format_monto=workbook.add_format({"num_format":"0.00"})
            clp_format_monto=workbook.add_format({"num_format":"0"})
            moneda_col_m=df_movs.columns.get_loc("Moneda")
            monto_col_m=df_movs.columns.get_loc("Monto")
            for row_num,row_data in enumerate(df_movs.itertuples(index=False),start=1):
                currency=row_data[moneda_col_m]
                monto_val=row_data[monto_col_m]
                if pd.notnull(monto_val):
                    if str(currency).upper()=="USD":
                        worksheet.write_number(row_num,monto_col_m,monto_val,usd_format_monto)
                    else:
                        worksheet.write_number(row_num,monto_col_m,monto_val,clp_format_monto)
    print(f"Informe generado: {out_file} a las {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__=="__main__":
    main()
