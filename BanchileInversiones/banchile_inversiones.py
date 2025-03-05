import os
import re
import datetime
import pandas as pd
from decimal import Decimal, InvalidOperation
from PyPDF2 import PdfReader

folder_input = "./Input/"
folder_output = "./Output/"
folder_txt = "./Output/Textos/"
os.makedirs(folder_txt, exist_ok=True)

pass_dict = {}
with open(os.path.join(folder_input, "pass.txt"), "r") as f:
    for line in f:
        k, v = line.strip().split(":")
        pass_dict[k] = v.strip()

def parse_number(num_str, is_percentage=False):
    if not num_str:
        return None
    original = num_str.replace("%", "").strip()
    s = original.replace(".", "").replace(",", ".")
    try:
        d = Decimal(s)
        if "," in original:
            dec_places = len(original.split(",")[1])
            quantizer = Decimal("1." + "0" * dec_places)
            d = d.quantize(quantizer)
        return d
    except InvalidOperation:
        return Decimal("0")

def extract_text_from_pdf(file_path, pass_key):
    reader = PdfReader(file_path)
    if reader.is_encrypted:
        reader.decrypt(pass_dict.get(pass_key, ""))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    txt_path = os.path.join(folder_txt, os.path.basename(file_path) + ".txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)
    return text

def clean_nemotecnico(nemo: str) -> str:
    nemo = re.sub(r"\([^)]*\d{2}/\d{2}/\d{4}[^)]*\)", "", nemo)
    nemo = re.sub(r"\([^)]*%[^)]*\)", "", nemo)
    nemo = re.sub(r"-\d+,\d+%", "", nemo)
    nemo = re.sub(r"REG SB%.*? B/E", "REG SB", nemo)
    nemo = nemo.replace("(PERSHING) (USA)", "")
    return re.sub(r"\s+", " ", nemo).strip().upper()

def shorten_nemotecnico(nemo: str) -> str:
    if " - " in nemo:
        return nemo.split(" - ")[0].strip()
    return nemo.strip()

def parse_rfi_two_lines(line1: str, line2_combined: str) -> dict:
    pat_line1 = re.compile(
        r"^(?P<fecha>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<cantidad>[\d\.,]+)\s+(?P<moneda1>CLP|USD|UF)\s+"
        r"(?P<price_mercado>[\d\.,]+)\s+(?P<moneda2>CLP|USD|UF)\s+"
        r"(?P<valor_mercado>[\d\.,]+)\s+(?P<moneda3>CLP|USD|UF)\s+"
        r"(?P<price_compra>[\d\.,]+)\s+(?P<moneda4>CLP|USD|UF)\s+"
        r"(?P<valor_compra>[\d\.,]+)"
    )
    m = pat_line1.match(line1.strip())
    if not m:
        return {}
    fecha_adq = m.group("fecha")
    cantidad = parse_number(m.group("cantidad"))
    price_mercado = parse_number(m.group("price_mercado"))
    valor_mercado = parse_number(m.group("valor_mercado"))
    price_compra = parse_number(m.group("price_compra"))
    valor_compra = parse_number(m.group("valor_compra"))
    moneda = m.group("moneda1")

    pat_isin = re.compile(r"(ISIN#\S+)", re.IGNORECASE)
    match_isin = pat_isin.search(line2_combined)
    if match_isin:
        isin_val = match_isin.group(1).replace("ISIN#", "")
    else:
        isin_val = ""

    cleaned = re.sub(r"ISIN#\S+", "", line2_combined, flags=re.IGNORECASE)
    cleaned = re.sub(r"[A-Z]{3}\s*[+-]?[\d\.,-]+\s*", "", cleaned, flags=re.IGNORECASE)
    nemotecnico = cleaned.strip().split("Subtotal")[0].strip()
    nemotecnico = clean_nemotecnico(nemotecnico)

    return {
        "Fecha_Adquisicion": fecha_adq,
        "Cantidad": cantidad,
        "Precio_Compra": price_compra,
        "Valor_Compra": valor_compra,
        "Precio_Mercado": price_mercado,
        "Valor_Mercado": valor_mercado,
        "Nemotecnico": nemotecnico,
        "ISIN": isin_val,
        "Clase_Activo": "Renta Fija Internacional",
        "Moneda": moneda
    }

def BanChile_Parser(filename):
    pdf_path = os.path.join(folder_input, f"{filename}.pdf")
    txt_path = os.path.join(folder_txt, f"{filename}.txt")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No existe {filename}.pdf")

    try:
        _, pass_key, _ = filename.split("_", 2)
    except ValueError:
        pass_key = ""

    if os.path.exists(txt_path):
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()
    else:
        text = extract_text_from_pdf(pdf_path, pass_key)

    nombre_match = re.search(r"ancla:.*\|.*\|.*\n([A-ZÑÁÉÍÓÚ ]+)", text)
    nombre = " ".join(nombre_match.group(1).split()) if nombre_match else "Desconocido"

    fecha_match = re.search(r"al\s+(\d{2}/\d{2}/\d{4})", text)
    fecha_dt = datetime.datetime.strptime(fecha_match.group(1), "%d/%m/%Y") if fecha_match else None

    lines = text.splitlines()
    cartera = []
    movimientos = []

    clase_activo_map = {
        "FONDOS MUTUOS": "FONDOS MUTUOS",
        "FONDOS DE INVERSIÓN": "FONDOS INVERSIÓN",
        "ACCIONES": "ACCIONES",
        "RENTA FIJA INTERNACIONAL": "Renta Fija Internacional",
        "RENTA FIJA": "RENTA FIJA",
        "RENTA FIJA CLP": "RENTA FIJA",
        "RENTA FIJA PESOS": "RENTA FIJA"
    }

    renta_fija_uf_pat = re.compile(
        r"^(?P<fecha_line>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<t_emision>[\d,]+)%\s+"
        r"(?P<t_compra>[\d,]+)%\s+"
        r"UF\s+"
        r"(?P<cantidad>[\d.,]+)\s+"
        r"UF\s+"
        r"(?P<valor_compra>[\d.,]+)\s+"
        r"(?P<t_mercado>[-\d,]+)%\s+"
        r"UF\s+"
        r"(?P<valor_mercado>[\d.,]+)"
    )

    renta_fija_clp_pat = re.compile(
        r"^(?P<fecha_venc>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<t_emision>[\d,]+%)\s+"
        r"(?P<t_compra>[\d,]+%)\s+"
        r"\$\s*(?P<valor_nominal>[\d\.]+)\s+"
        r"\$\s*(?P<monto_inicio>[\d\.]+)\s+"
        r"(?P<t_mercado>[\d,]+%)\s+"
        r"\$\s*(?P<valor_mercado>[\d\.]+)"
        r"(?:\s+\$\s*(?P<ganancia>[\d\.]+))?\s+"
        r"(?P<nemotecnico>.*?)\s+\(\d{2}/\d{2}/\d{4}\)"
    )

    date_pat = re.compile(r"^\d{2}/\d{2}/\d{4}")
    dividend_pat = re.compile(r"^Pago de Dividendo", re.IGNORECASE)
    pattern_standard = re.compile(
        r"^(?P<fecha>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<tipo>\w+)\s+"
        r"(?:(?:USD\s+)?\$?\s*(?P<precio>[\d\.,-]+|--))\s+"
        r"(?P<cantidad>[\d\.,-]+|--)\s+"
        r"(?:(?:USD\s+)?\$?\s*(?P<monto>[\d\.,-]+|--))\s+"
        r"(?:\$?\s*)?(?P<descripcion>.*)$",
        re.IGNORECASE
    )
    pattern_date = re.compile(r"^(?P<fecha_liq>\d{2}/\d{2}/\d{4})$")
    mode = None
    current_currency = None
    current_cuenta = None
    current_clase_activo = ""
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        if re.search(r"DETALLE DE POSICIONES EN", line, re.IGNORECASE):
            mode = "posiciones"
            current_currency = "CLP" if "PESOS" in line.upper() else "USD" if "DÓLARES" in line.upper() else "UF" if "UF" in line.upper() else None
            cta_match = re.search(r"CUENTA[:\s]*(\d+)", line, re.IGNORECASE)
            current_cuenta = cta_match.group(1) if cta_match else "0"
            i += 1
            continue

        if re.search(r"DETALLE DE MOVIMIENTOS DE CAJA", line, re.IGNORECASE):
            mode = "movimientos"
            cta_match = re.search(r"CUENTA[:\s]*(\d+)", line, re.IGNORECASE)
            current_cuenta = cta_match.group(1) if cta_match else "0"
            i += 1
            continue

        if mode == "posiciones":
            for sub, clase_mapped in clase_activo_map.items():
                if sub in line.upper():
                    current_clase_activo = clase_mapped

            if "RENTA FIJA INTERNACIONAL" in line.upper():
                current_clase_activo = "Renta Fija Internacional"
                i += 1
                while i < len(lines):
                    current_line = lines[i].strip()
                    if re.search(r"DETALLE DE (POSICIONES|MOVIMIENTOS)", current_line, re.IGNORECASE):
                        break
                    if re.match(r"^\d{2}/\d{2}/\d{4}", current_line):
                        line1 = current_line
                        line2_combined = []
                        j = i + 1
                        while j < len(lines):
                            next_line = lines[j].strip()
                            if re.match(r"^\d{2}/\d{2}/\d{4}", next_line) or re.search(r"DETALLE DE", next_line, re.IGNORECASE):
                                break
                            line2_combined.append(next_line)
                            j += 1
                        line2_combined = " ".join(line2_combined)
                        parsed = parse_rfi_two_lines(line1, line2_combined)
                        if parsed:
                            record = {
                                "Fecha": fecha_dt.strftime("%d/%m/%Y") if fecha_dt else "",
                                "Nombre": nombre,
                                "Cuenta": current_cuenta,
                                "Nemotecnico": parsed["Nemotecnico"],
                                "Moneda": parsed["Moneda"],
                                "ISIN": parsed["ISIN"],
                                "CUSIP": "",
                                "Cantidad": parsed["Cantidad"],
                                "Precio_Mercado": parsed["Precio_Mercado"],
                                "Valor_Mercado": parsed["Valor_Mercado"],
                                "Precio_Compra": parsed["Precio_Compra"],
                                "Valor_Compra": parsed["Valor_Compra"],
                                "Clase_Activo": parsed["Clase_Activo"],
                                "Contraparte": "BanChile"
                            }
                            cartera.append(record)
                        i = j - 1
                    i += 1
                continue

            if current_currency == "UF":
                m_uf = renta_fija_uf_pat.match(line)
                if m_uf:
                    cantidad = parse_number(m_uf.group("cantidad"))
                    valor_compra = parse_number(m_uf.group("valor_compra"))
                    precio_compra = parse_number(m_uf.group("t_compra"), True)
                    precio_mercado = parse_number(m_uf.group("t_mercado"), True)
                    valor_mercado = parse_number(m_uf.group("valor_mercado"))

                    j = i + 1
                    clp_value = None
                    nemotecnico_line = ""
                    while j < len(lines):
                        nxt = lines[j].strip()
                        if not nxt or re.match(r"^\d{2}/\d{2}/\d{4}", nxt) or re.search(r"DETALLE DE POSICIONES EN", nxt, re.IGNORECASE) or re.search(r"DETALLE DE MOVIMIENTOS DE CAJA", nxt, re.IGNORECASE):
                            break
                        if nxt.startswith("$"):
                            m_clp = re.search(r"^\$\s*([\d\.,]+)", nxt)
                            if m_clp:
                                clp_value = parse_number(m_clp.group(1))
                            m_nemo = re.search(r"^\$\s*[\d\.,]+UF\s*[-\d\.,]+\s*([A-Z0-9]+)\s*-", nxt)
                            if m_nemo:
                                nemotecnico_line = m_nemo.group(1).strip()
                        j += 1

                    if clp_value is not None:
                        valor_mercado = clp_value

                    if nemotecnico_line:
                        nemotecnico = nemotecnico_line
                    else:
                        nemo_lines = []
                        for k in range(i+1, j):
                            nemo_lines.append(lines[k].strip())
                        nemotecnico_text = " ".join(nemo_lines)
                        m_nemo2 = re.search(r"([A-Z0-9]{3,})\s*-", nemotecnico_text)
                        if m_nemo2:
                            nemotecnico = m_nemo2.group(1).strip()

                    rec_fecha = fecha_dt.strftime("%d/%m/%Y") if fecha_dt else ""
                    record = {
                        "Fecha": rec_fecha,
                        "Nombre": nombre,
                        "Cuenta": current_cuenta,
                        "Nemotecnico": nemotecnico,
                        "Moneda": "UF",
                        "ISIN": "",
                        "CUSIP": "",
                        "Cantidad": cantidad,
                        "Precio_Mercado": precio_mercado,
                        "Valor_Mercado": valor_mercado,
                        "Precio_Compra": precio_compra,
                        "Valor_Compra": valor_compra,
                        "Clase_Activo": current_clase_activo or "RENTA FIJA",
                        "Contraparte": "BanChile"
                    }
                    cartera.append(record)
                    i = j
                    continue

            if current_currency == "CLP":
                m_clp = renta_fija_clp_pat.match(line)
                if m_clp and "RENTA FIJA" in current_clase_activo.upper():
                    record = {
                        "Fecha": fecha_dt.strftime("%d/%m/%Y") if fecha_dt else "",
                        "Nombre": nombre,
                        "Cuenta": current_cuenta,
                        "Nemotecnico": shorten_nemotecnico(m_clp.group("nemotecnico")),
                        "Moneda": "CLP",
                        "ISIN": "",
                        "CUSIP": "",
                        "Cantidad": parse_number(m_clp.group("valor_nominal")),
                        "Precio_Compra": parse_number(m_clp.group("t_compra"), True),
                        "Valor_Compra": parse_number(m_clp.group("monto_inicio")),
                        "Precio_Mercado": parse_number(m_clp.group("t_mercado"), True),
                        "Valor_Mercado": parse_number(m_clp.group("valor_mercado")),
                        "Clase_Activo": "RENTA FIJA",
                        "Contraparte": "BanChile"
                    }
                    cartera.append(record)
                    i += 1
                    continue

            if current_currency in ("CLP", "USD"):
                if current_clase_activo.upper() == "RENTA FIJA" and re.match(r"^(\d{2}/\d{2}/\d{4}|[\d,.]+%)", line):
                    record_line = line
                    j = i + 1
                    while j < len(lines):
                        nxt = lines[j].strip()
                        if (re.match(r"^(\d{2}/\d{2}/\d{4}|[\d,.]+%)", nxt) and "$" in nxt) or re.search(r"DETALLE DE", nxt, re.IGNORECASE):
                            break
                        record_line += " " + nxt
                        j += 1

                    parts = [p.strip() for p in re.split(r"\$", record_line) if p.strip()]
                    if len(parts) >= 5:
                        p0 = parts[0].split()
                        t_emision = ""
                        t_compra = ""
                        if p0:
                            if re.match(r"^\d{2}/\d{2}/\d{4}$", p0[0]):
                                percentages = [s for s in p0[1:] if '%' in s]
                                if len(percentages) >= 1:
                                    t_emision = percentages[0]
                                if len(percentages) >= 2:
                                    t_compra = percentages[1]
                            else:
                                percentages = [s for s in p0 if '%' in s]
                                if len(percentages) >= 1:
                                    t_compra = percentages[0]
                                if len(percentages) >= 2:
                                    t_emision, t_compra = percentages[0], percentages[1]

                        valor_nominal = parts[1]
                        p2 = parts[2].split()
                        monto_inicio = p2[0] if p2 else ""
                        t_mercado = p2[1] if len(p2) > 1 else ""
                        valor_actual = parts[3]
                        p4 = parts[4].split()
                        nemotecnico = " ".join(p4[1:]) if len(p4) >= 2 else parts[4]

                        rec_fecha = fecha_dt.strftime("%d/%m/%Y") if fecha_dt else ""
                        record = {
                            "Fecha": rec_fecha,
                            "Nombre": nombre,
                            "Cuenta": current_cuenta,
                            "Nemotecnico": shorten_nemotecnico(nemotecnico),
                            "Moneda": current_currency,
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": parse_number(valor_nominal),
                            "Precio_Compra": parse_number(t_compra, True),
                            "Valor_Compra": parse_number(monto_inicio),
                            "Precio_Mercado": parse_number(t_mercado, True),
                            "Valor_Mercado": parse_number(valor_actual),
                            "Clase_Activo": current_clase_activo or "RENTA FIJA",
                            "Contraparte": "BanChile"
                        }
                        cartera.append(record)
                    i = j
                    continue
                else:
                    delimiter = r"\$" if current_currency == "CLP" else r"USD"
                    parts = [p.strip() for p in re.split(delimiter, line) if p.strip()]
                    if len(parts) >= 7:
                        cantidad_str = parts[0].split()[0] if parts[0] else "0"
                        precio_compra_str = parts[1].split()[0] if len(parts) > 1 else "0"
                        valor_compra_str = parts[2].split()[0] if len(parts) > 2 else "0"
                        precio_mercado_str = parts[3].split()[0] if len(parts) > 3 else "0"
                        split_last = parts[6].split()
                        valor_mercado_str = split_last[0] if split_last else "0"
                        rest_text = " ".join(split_last[1:]) if len(split_last) > 1 else ""
                        nemotecnico = re.sub(r"\s*\d+([.,]\d+)?%\s*$", "", rest_text).strip()

                        rec_fecha = fecha_dt.strftime("%d/%m/%Y") if fecha_dt else ""
                        record = {
                            "Fecha": rec_fecha,
                            "Nombre": nombre,
                            "Cuenta": current_cuenta,
                            "Nemotecnico": shorten_nemotecnico(nemotecnico),
                            "Moneda": current_currency,
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": parse_number(cantidad_str),
                            "Precio_Compra": parse_number(precio_compra_str),
                            "Valor_Compra": parse_number(valor_compra_str),
                            "Precio_Mercado": parse_number(precio_mercado_str),
                            "Valor_Mercado": parse_number(valor_mercado_str),
                            "Clase_Activo": current_clase_activo or "SIN CLASE",
                            "Contraparte": "BanChile"
                        }
                        cartera.append(record)
                    i += 1
                    continue

        elif mode == "movimientos":
            if date_pat.match(line) or dividend_pat.match(line):
                consolidated = [line]
                i2 = i + 1
                while i2 < len(lines):
                    l2 = lines[i2].strip()
                    if re.search(r"DETALLE DE POSICIONES EN", l2, re.IGNORECASE) or re.search(r"DETALLE DE MOVIMIENTOS DE CAJA", l2, re.IGNORECASE) or date_pat.match(l2) or dividend_pat.match(l2):
                        break
                    consolidated[-1] += " " + l2
                    i2 += 1

                rec_line = consolidated[0]
                if dividend_pat.match(rec_line):
                    fecha_mov = fecha_liq = "??/??/????"
                    monto_match = re.search(r"\$([\d\.,-]+)", rec_line)
                    monto = parse_number(monto_match.group(1)) if monto_match else None
                    if fecha_mov != "??/??/????" and fecha_liq != "??/??/????":
                        movimientos.append({
                            "Fecha Movimiento": fecha_mov,
                            "Fecha Liquidación": fecha_liq,
                            "Nombre": nombre,
                            "RUT": "",
                            "Cuenta": current_cuenta,
                            "Nemotecnico": "",
                            "Moneda": "CLP",
                            "ISIN": "",
                            "CUSIP": "",
                            "Cantidad": None,
                            "Precio": None,
                            "Monto": monto,
                            "Comision": "",
                            "tipo": "Abono",
                            "descripcion": rec_line,
                            "Concepto": "",
                            "Folio": "",
                            "Contraparte": "BanChile"
                        })
                    i = i2
                    continue

                m_std = pattern_standard.match(rec_line)
                if m_std:
                    try:
                        fecha_mov = datetime.datetime.strptime(m_std.group("fecha"), "%d/%m/%Y").strftime("%d/%m/%Y")
                    except:
                        fecha_mov = "??/??/????"
                    desc = m_std.group("descripcion").strip()
                    fecha_liq_pat = re.search(r"(\d{2}/\d{2}/\d{4})$", desc)
                    if fecha_liq_pat:
                        fecha_liq_str = fecha_liq_pat.group(1)
                        try:
                            fecha_liq = datetime.datetime.strptime(fecha_liq_str, "%d/%m/%Y").strftime("%d/%m/%Y")
                            desc = desc.replace(fecha_liq_str, "").strip()
                        except:
                            fecha_liq = "??/??/????"
                    else:
                        fecha_liq = fecha_mov
                    if fecha_mov == "??/??/????" or fecha_liq == "??/??/????":
                        i = i2
                        continue
                    
                    desc_clean = re.sub(r"^\$?\s*-?[\d\.,]+\s*", "", desc).strip()
                    desc_clean = re.split(r"\s+\d{2}/\d{2}/\d{4}", desc_clean)[0].strip()
                    nemotecnico = ""
                    descripcion = desc_clean
                    if ":" in desc_clean:
                        split_desc = desc_clean.split(":", 1)
                        desc_tmp = split_desc[0].strip()
                        nemo_candidate = re.sub(r"\s*\d{2}/\d{2}/\d{4}.*", "", split_desc[1].strip()).strip()
                        if nemo_candidate and not re.match(r"^(\$|USD|CLP|UF)\b", nemo_candidate):
                            descripcion = desc_tmp
                            nemotecnico = shorten_nemotecnico(clean_nemotecnico(nemo_candidate))
                    movimientos.append({
                        "Fecha Movimiento": fecha_mov,
                        "Fecha Liquidación": fecha_liq,
                        "Nombre": nombre,
                        "RUT": "",
                        "Cuenta": current_cuenta,
                        "Nemotecnico": nemotecnico,
                        "Moneda": "USD" if "USD" in rec_line else "CLP",
                        "ISIN": "",
                        "CUSIP": "",
                        "Cantidad": parse_number(m_std.group("cantidad")) if m_std.group("cantidad") != "--" else None,
                        "Precio": parse_number(m_std.group("precio")) if m_std.group("precio") != "--" else None,
                        "Monto": parse_number(m_std.group("monto")) if m_std.group("monto") != "--" else None,
                        "Comision": "",
                        "tipo": m_std.group("tipo"),
                        "descripcion": descripcion,
                        "Concepto": "",
                        "Folio": "",
                        "Contraparte": "BanChile"
                    })
                i = i2
                continue
        i += 1

    return cartera, movimientos, nombre

if __name__ == "__main__":
    os.makedirs(folder_output, exist_ok=True)
    all_cartera = []
    all_movimientos = []
    for file in os.listdir(folder_input):
        if not file.lower().endswith(".pdf"):
            continue
        filename_noext = os.path.splitext(file)[0]
        try:
            cartera, movimientos, nombre = BanChile_Parser(filename_noext)
            if cartera:
                all_cartera.extend(cartera)
            if movimientos:
                all_movimientos.extend(movimientos)
        except Exception as e:
            print(f"Error procesando {file}: {str(e)}")
    if all_cartera or all_movimientos:
        fecha_informe = datetime.datetime.now()
        if all_cartera:
            primera_fecha = all_cartera[0].get("Fecha", "")
            if primera_fecha:
                try:
                    fecha_informe = datetime.datetime.strptime(primera_fecha, "%d/%m/%Y")
                except:
                    pass

        fecha_formateada = fecha_informe.strftime("%Y%m%d")
        output_path = os.path.join(folder_output, f"Informe_{fecha_formateada}.xlsx")
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            if all_cartera:
                df_cartera = pd.DataFrame(all_cartera)
                df_cartera["Rut"] = ""
                df_cartera["interes_Acum"] = ""
                final_cols = [
                    "Fecha",
                    "Nombre",
                    "Rut",
                    "Cuenta",
                    "Nemotecnico",
                    "Moneda",
                    "ISIN",
                    "CUSIP",
                    "Cantidad",
                    "Precio_Mercado",
                    "Valor_Mercado",
                    "Precio_Compra",
                    "Valor_Compra",
                    "interes_Acum",
                    "Contraparte",
                    "Clase_Activo"
                ]
                df_cartera = df_cartera.reindex(columns=final_cols)
                df_cartera = df_cartera.dropna(how="all")
                df_cartera.to_excel(writer, index=False, sheet_name="Cartera")
                workbook = writer.book
                worksheet = writer.sheets["Cartera"]
                numeric_cols = ["Cantidad", "Precio_Mercado", "Valor_Mercado", "Precio_Compra", "Valor_Compra"]
                for col_name in numeric_cols:
                    if col_name in df_cartera.columns:
                        col_idx = df_cartera.columns.get_loc(col_name)
                        for row_idx, value in enumerate(df_cartera[col_name], start=1):
                            if isinstance(value, Decimal):
                                exponent = value.as_tuple().exponent
                                dec_places = -exponent if exponent < 0 else 0
                                num_format = "0" + ("." + "0" * dec_places if dec_places > 0 else "")
                                cell_format = workbook.add_format({'num_format': num_format})
                                try:
                                    numeric_val = float(value)
                                except:
                                    numeric_val = 0.0
                                worksheet.write_number(row_idx, col_idx, numeric_val, cell_format)

            if all_movimientos:
                df_movimientos = pd.DataFrame(all_movimientos)
                df_movimientos = df_movimientos.dropna(how="all")
                df_movimientos.to_excel(writer, index=False, sheet_name="Movimientos")
                workbook = writer.book
                worksheet = writer.sheets["Movimientos"]
                numeric_cols_mov = ["Cantidad", "Precio", "Monto"]
                for col_name in numeric_cols_mov:
                    if col_name in df_movimientos.columns:
                        col_idx = df_movimientos.columns.get_loc(col_name)
                        for row_idx, value in enumerate(df_movimientos[col_name], start=1):
                            if isinstance(value, Decimal):
                                exponent = value.as_tuple().exponent
                                dec_places = -exponent if exponent < 0 else 0
                                num_format = "0" + ("." + "0" * dec_places if dec_places > 0 else "")
                                cell_format = workbook.add_format({'num_format': num_format})
                                try:
                                    numeric_val = float(value)
                                except:
                                    numeric_val = 0.0
                                worksheet.write_number(row_idx, col_idx, numeric_val, cell_format)
        print(output_path)
    else:
        print("No se encontraron datos para generar el informe")
