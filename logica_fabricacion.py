import pandas as pd
import pdfplumber
import PyPDF2
import re
import re
from collections import defaultdict
import math
import io

def procesar_csv_lantek(csv_file) -> dict:
    """
    Parses Lantek's 'Listado de hojas de taller' CSV export using standard csv module.
    Returns the same structure as procesar_pdf:
    {'1': {'W-42': 2.0, ...}, '2': ...}
    """
    csv_file.seek(0)
    try:
        content = csv_file.read().decode('latin1')
    except Exception:
        csv_file.seek(0)
        try:
            content = csv_file.read().decode('utf-8', errors='ignore')
        except AttributeError:
            csv_file.seek(0)
            content = csv_file.read()
            
    import csv
    import io
    import collections
    import re
    
    f_str = io.StringIO(content)
    reader = csv.reader(f_str)
    
    config_inventario = {}
    config_aliases = {}
    hojas_vistas = []
    
    for row in reader:
        # Check if the row has enough columns
        if len(row) < 56:
            continue
            
        # Filter only main data rows
        if row[4] != "DATOS DE LA CHAPA":
            continue
            
        cnc_raw = str(row[7]).strip().upper()
        ref = row[10].strip().upper()
        
        # Extract number of sheets of this setup
        try:
            cant_chapas = float(row[16])
        except (ValueError, TypeError):
            cant_chapas = 1.0
        if cant_chapas == 0: 
            cant_chapas = 1.0
            
        if ref not in config_inventario:
            config_inventario[ref] = collections.defaultdict(float)
            hojas_vistas.append(ref)
            
            # Construir todos los alias posibles para poder cruzar los datos
            aliases = set([ref, ref.replace(" ", "")])
            if cnc_raw and cnc_raw != 'NAN':
                aliases.add(cnc_raw)
            match_hoja = re.search(r'\((\d+)\)', ref)
            if match_hoja:
                aliases.add(match_hoja.group(1))
            else:
                aliases.add(str(len(hojas_vistas)))
            config_aliases[ref] = aliases
            
        # Extract piece definition
        pieza_raw = row[51].strip()
        
        # Valid prefixes: W, WB, FL, ST, PC, PL, etc.
        match_pieza = re.search(r'((?:W|WB|FL|ST|PC|PL)-\S+)', pieza_raw)
        if not match_pieza:
            continue
        pieza = match_pieza.group(1).strip()
        
        # Last column is total quantity (replace comma for dot in spanish decimals)
        try:
            total_piezas = float(row[55].replace('"', '').replace(',', '.').strip())
        except (ValueError, TypeError):
            total_piezas = 0.0
            
        # Distribute into qty per individual sheet
        qty_per_sheet = total_piezas / cant_chapas
        config_inventario[ref][pieza] += qty_per_sheet

    pdf_inventario = {}
    for ref_key, inv_dict in config_inventario.items():
        for alias in config_aliases[ref_key]:
            pdf_inventario[alias] = inv_dict

    if pdf_inventario:
        pdf_inventario['_raw_text_1'] = f"CSV LECTURA OK (ALIASES ACTIVADOS): Extraídas {len(hojas_vistas)} hojas únicas de Lantek."

    return pdf_inventario

def procesar_nesting_file(file) -> dict:
    name = getattr(file, "name", "").lower()
    if name.endswith('.csv'):
        return procesar_csv_lantek(file)
    else:
        return procesar_pdf(file)

def procesar_pdf(pdf_file) -> dict:
    """
    Recibe un objeto archivo PDF (de Streamlit) y retorna un diccionario:
    {
      'hoja_X': {'W-42': 2.0, 'FL-66': 1.0, ...},
      ...
    }
    """
    pdf_inventario = {}
    
    # Intentar con pdfplumber primero
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            hoja_num = str(i + 1)
            pdf_inventario[hoja_num] = defaultdict(float)
            
            text = page.extract_text()
            
            # Si hay problemas de fuente CID (caracteres corruptos), intentamos con PyPDF2
            if text and "(cid:" in text:
                break # Rompemos y probamos el Plan B
                
            if i == 0:
                pdf_inventario['_raw_text_1'] = text[:1000] if text else "TEXTO VACIO"
                
            if not text:
                continue
                
            cant_match = re.search(r"Cantidad\s+(\d+)", text, re.IGNORECASE)
            cantidad = float(cant_match.group(1)) if cant_match else 1.0
            if cantidad == 0: cantidad = 1.0
            
            for line in text.split('\n'):
                match = re.search(r"((?:W|WB|FL)-\S+)\s+(\d+)", line)
                if match:
                    pieza = match.group(1).strip()
                    chapas_totales = float(match.group(2))
                    pdf_inventario[hoja_num][pieza] = chapas_totales / cantidad
                    
    # PLAN B: Si pdfplumber fracasó por fuentes CID
    if pdf_inventario and pdf_inventario.get('1') == defaultdict(float) and "(cid:" in pdf_inventario.get('_raw_text_1', ''):
        pdf_file.seek(0)
        reader = PyPDF2.PdfReader(pdf_file)
        pdf_inventario = {}
        for i, page in enumerate(reader.pages):
            hoja_num = str(i + 1)
            pdf_inventario[hoja_num] = defaultdict(float)
            text = page.extract_text()
            
            if i == 0:
                pdf_inventario['_raw_text_1'] = "FALLBACK PYPDF2:\n" + (text[:1000] if text else "VACIO")
                
            if not text:
                continue
                
            cant_match = re.search(r"Cantidad\s+(\d+)", text, re.IGNORECASE)
            cantidad = float(cant_match.group(1)) if cant_match else 1.0
            if cantidad == 0: cantidad = 1.0
            
            for line in text.split('\n'):
                match = re.search(r"((?:W|WB|FL)-\S+)\s+(\d+)", line)
                if match:
                    pieza = match.group(1).strip()
                    chapas_totales = float(match.group(2))
                    pdf_inventario[hoja_num][pieza] = chapas_totales / cantidad
                    
    return pdf_inventario

def generar_inventario_cortado(csv_file, pdf_dicts) -> dict:
    """
    Lee el CSV de Airtable y genera el inventario real de piezas cortadas.

    RUTA A — Pieza Única: fila con 'Tipo de Corte' == 'Pieza Unica'.
      El nombre de la pieza viene de la columna DETALLE/CNC (ej: FL-4),
      y la cantidad de 'Cantidad Cortada'. Va directo al inventario sin
      necesitar nesting.

    RUTA B — Chapa de Nesting normal: cruce con el PDF/CSV de Lantek.
    """
    df_csv = pd.read_csv(csv_file)
    inventario = defaultdict(float)
    
    # Normalizar columnas
    cols = [str(c).lower().strip() for c in df_csv.columns]
    df_csv.columns = cols
    
    col_cnc        = None
    col_hoja       = None
    col_cant       = None
    col_tipo_corte = None   # "Tipo de Corte" → "Pieza Unica" / "Chapa Entera"
    col_detalle    = None   # columna con nombre de pieza (DETALLE / Prog CNC)
    
    for c in df_csv.columns:
        if 'cnc' in c:
            col_cnc = c
        if 'detalle' in c or 'descripcion' in c or 'nombre' in c:
            col_detalle = c   # columna alternativa con nombre de pieza
        if 'hoja' in c:
            col_hoja = c
        if ('cantidad cortada' in c or 'cortada' in c) and 'peso' not in c:
            col_cant = c
        if 'tipo' in c and 'corte' in c:
            col_tipo_corte = c

    # La columna con el nombre de la pieza es DETALLE si existe, sino CNC
    col_nombre_pieza = col_detalle or col_cnc
        
    debug_info = {
        "Columnas detectadas": cols,
        "Col Hoja": col_hoja,
        "Col CNC / Detalle": col_nombre_pieza,
        "Col Cant Cortada": col_cant,
        "Col Tipo de Corte": col_tipo_corte,
        "Total Hojas PDF": sum(len(d) for d in pdf_dicts) if pdf_dicts else 0,
        "Muestra Datos de PDF (Hoja 1)": pdf_dicts[0].get('1', {}) if pdf_dicts else {},
        "Raw Text Pagina 1": pdf_dicts[0].get('_raw_text_1', "Vacio") if pdf_dicts else "Sin PDF aportado",
        "Resultados Validos (Cant > 0)": [],
        "Piezas Únicas Sumadas": []
    }
    
    if (not col_hoja and not col_nombre_pieza) or not col_cant:
        return {}, debug_info
        
    for _, row in df_csv.iterrows():
        try:
            cant_cortada_raw = row[col_cant]
            if pd.isna(cant_cortada_raw):
                continue
            cant_cortada = float(cant_cortada_raw)
            if cant_cortada <= 0:
                continue

            # -------------------------------------------------------
            # RUTA A: PIEZA EXPLÍCITA ("Tipo de Corte" == "Pieza Unica" o Nombre Válido)
            # El nombre de la pieza viene del campo DETALLE/CNC.
            # Se suma directamente al inventario sin cruzar con nesting.
            # -------------------------------------------------------
            es_pieza_explicita = False
            if col_tipo_corte and pd.notna(row[col_tipo_corte]):
                tipo_val = str(row[col_tipo_corte]).strip().lower()
                if 'unica' in tipo_val or 'única' in tipo_val or 'suelta' in tipo_val:
                    es_pieza_explicita = True

            pieza_raw = ""
            if col_nombre_pieza and pd.notna(row[col_nombre_pieza]):
                pieza_raw = str(row[col_nombre_pieza]).strip().upper()
                if pieza_raw and pieza_raw not in ('NAN', '', '-'):
                    # Si tiene pinta de nombre de pieza (W-, FL-, etc.) asumimos que es pieza explícita
                    if re.match(r'^(W|WB|FL|ST|PC|PL|C|V)-', pieza_raw) or not pdf_dicts:
                        es_pieza_explicita = True

            if es_pieza_explicita and pieza_raw and pieza_raw not in ('NAN', '', '-'):
                # Normalizar a base (FL-4-A → FL-4)
                base_m = re.match(r'^([A-Z]+-\d+)', pieza_raw)
                pieza_key = base_m.group(1) if base_m else pieza_raw
                inventario[pieza_key] += cant_cortada
                debug_info["Piezas Únicas Sumadas"].append(
                    f"{pieza_key} += {cant_cortada:.0f} (pieza explícita directa)"
                )
                continue   # No procesar por la ruta de nesting

            # -------------------------------------------------------
            # RUTA B: CHAPA DE NESTING (cruce con PDF/CSV Lantek)
            # -------------------------------------------------------
            hoja_str = ""
            if col_hoja and pd.notna(row[col_hoja]):
                hoja_str = str(int(float(row[col_hoja])))
                
            cnc_str = ""
            if col_cnc and pd.notna(row[col_cnc]):
                val_cnc = str(row[col_cnc]).strip().upper()
                if val_cnc and val_cnc != 'NAN':
                    cnc_str = val_cnc
                    
            # Claves para cruzar con el diccionario Lantek
            claves_a_buscar = []
            if cnc_str: 
                claves_a_buscar.append(cnc_str)
                claves_a_buscar.append(cnc_str.replace(" ", ""))
            if hoja_str: 
                claves_a_buscar.append(hoja_str)
            
            # Buscar en todos los diccionarios de PDF subidos esta hoja/cnc
            matriz_hoja = {}
            for p_dict in pdf_dicts:
                for clave in claves_a_buscar:
                    if clave in p_dict:
                        matriz_hoja = p_dict[clave]
                        break
                if matriz_hoja:
                    break
                    
            if not matriz_hoja:
                debug_info["Resultados Validos (Cant > 0)"].append(str(claves_a_buscar) + " (Sin Match PDF)")
                continue
                
            # Añadir al inventario total
            suma = 0
            for pieza, qty_per_sheet in matriz_hoja.items():
                inventario[pieza] += qty_per_sheet * cant_cortada
                suma += 1
            
            debug_info["Resultados Validos (Cant > 0)"].append(str(hoja_str) + f" (Match PDF: {suma} tipos de pieza)")
                
        except (ValueError, TypeError):
            continue
            
    return dict(inventario), debug_info


def calcular_armables(excel_file, inventario) -> dict:
    """
    Recibe el Excel BOM y el inventario real.
    Determina cuantas piezas principales se pueden armar.
    Retorna (df_resultado, metricas)
    """
    # Leer excel crudo
    df_raw = pd.read_excel(excel_file, sheet_name=0, header=None)
    
    # Formato BOM típico de Cristhian
    # Asumiremos la columna ESTRUCTURA CRITICA (A=Cant_Conjunto, B=Nombre, C=Cant_Pos, D=Posicion)
    # Buscaremos la primera fila donde haya un W-, WB- o FL- para "anclar" los datos.
    
    # Limpiado dinámico
    re_pieza_critica = re.compile(r'^(W|WB|FL)-')
    
    requerimientos_conjunto = defaultdict(lambda: defaultdict(float))
    cantidades_conjunto_pedidas = defaultdict(float)
    
    col_conjunto = 1    # Nombre viga (ej. V-112)
    col_cant_pos = 2    # Cant por conjunto (ej. 2)
    col_pos = 3         # Nombre pos (ej. FL-66)
    
    ultima_viga = "DESCONOCIDA"
    
    for _, row in df_raw.iterrows():
        # Extracción segura
        viga_val = str(row[col_conjunto]).strip() if pd.notna(row[col_conjunto]) else ""
        pos_val = str(row[col_pos]).strip() if pd.notna(row[col_pos]) else ""
        
        try:
            cant_pos = float(row[col_cant_pos])
        except (ValueError, TypeError):
            cant_pos = 0.0
            
        # Detectar si esta línea es la cabecera de un ensamblaje
        # Regla robusta: Col 0 es numérico (cantidad pedida) y Col 3 (posición) está vacía
        es_ensamblaje = False
        try:
            cant_pedida_raw = float(row[0])
            pos_val_lower = pos_val.lower()
            if cant_pedida_raw > 0 and (pos_val_lower == 'nan' or pos_val_lower == ''):
                es_ensamblaje = True
                cant_conj_pedida = cant_pedida_raw
        except (ValueError, TypeError):
            pass

        if es_ensamblaje:
            nombre_col2 = str(row[2]).strip() if pd.notna(row[2]) else ""
            nombre_col1 = str(row[1]).strip() if pd.notna(row[1]) else ""
            
            # Algunos proyectistas ponen el nombre en Col C (2), otros en Col B (1)
            # Priorizamos Col 2 si no es numérico ni vacío (ej. "C1", "C2")
            # Si Col 2 está vacío o es un número puro, usamos Col 1 (ej "V-112")
            if nombre_col2 and nombre_col2.lower() != 'nan':
                # Validar que no sea un simple número de cantidad
                is_num = False
                try:
                    float(nombre_col2)
                    is_num = True
                except ValueError:
                    pass
                if not is_num:
                    ultima_viga = nombre_col2
                elif nombre_col1 and nombre_col1.lower() != 'nan':
                    ultima_viga = nombre_col1
            elif nombre_col1 and nombre_col1.lower() != 'nan':
                ultima_viga = nombre_col1
                
            cantidades_conjunto_pedidas[ultima_viga] = cant_conj_pedida
                
        if re_pieza_critica.match(pos_val) and cant_pos > 0:
            requerimientos_conjunto[ultima_viga][pos_val] += cant_pos

    # ---------------------------------------------
    # Algoritmo de Armado: Cruce de Inventario
    # ---------------------------------------------
    resultados_armables = []
    
    # Clonar para descontar
    inv_disponible = dict(inventario)
    
    total_conjuntos = len(requerimientos_conjunto)
    completos_count = 0
    
    for conjunto, requerimientos in requerimientos_conjunto.items():
        # Calcular cuantos de "conjunto" podemos armar con el stock inicial
        posibles_armar_iter = []
        for pos, cant_req in requerimientos.items():
            stock = inv_disponible.get(pos, 0.0)
            
            # Chequeo de sub-piezas derivadas (ej: FL-105-A, FL-105-B)
            patron_sub = re.compile(f'^{re.escape(pos)}[-]?([a-zA-Z])$', re.IGNORECASE)
            subpiezas_disp = {}
            for k, val in inv_disponible.items():
                m = patron_sub.match(k)
                if m:
                    subpiezas_disp[m.group(1).upper()] = val
            
            # Asumimos que se requieren al menos las partes A y B para ensamblar la pieza entera
            stock_compuesto = 0.0
            if 'A' in subpiezas_disp and 'B' in subpiezas_disp:
                stock_compuesto = min(subpiezas_disp.values())
            
            stock_total = stock + stock_compuesto
            if cant_req > 0:
                posibles = math.floor(stock_total / cant_req)
                posibles_armar_iter.append(posibles)
                
        if posibles_armar_iter:
            cantidad_armable = min(posibles_armar_iter)
        else:
            cantidad_armable = 0
            
        # Si no logramos leer la cantidad del Excel, asumimos 1 como base mínima.
        # Fallback previo tomaba max(cantidad_armable, 1) creando un falso llenado de pedidos infinitos
        pedido = cantidades_conjunto_pedidas.get(conjunto, 1.0)
        
        # Cantidad real a procesar para descuento de inventario: No podemos armar mas del pedido 
        # para no robarle inventario irresponsablemente a los ensamblajes de debajo.
        armados_reales = min(cantidad_armable, pedido)
        
        # Descontar del inventario general para que no se re-use el material
        if armados_reales > 0:
            for pos, cant_req in requerimientos.items():
                a_descontar = armados_reales * cant_req
                
                stock_directo = inv_disponible.get(pos, 0.0)
                descuento_directo = min(a_descontar, stock_directo)
                inv_disponible[pos] = stock_directo - descuento_directo
                
                restante = a_descontar - descuento_directo
                if restante > 0:
                    patron_sub = re.compile(f'^{re.escape(pos)}[-]?([a-zA-Z])$', re.IGNORECASE)
                    for k in list(inv_disponible.keys()):
                        if patron_sub.match(k):
                            inv_disponible[k] -= restante
                
        # Estado
        if cantidad_armable >= pedido:
            estado = "✅ ARMABLE AL 100%"
            completos_count += 1
        elif cantidad_armable > 0:
            estado = f"⚠️ PARCIAL ({cantidad_armable} de {int(pedido)})"
        else:
            estado = "❌ INCOMPLETO"
            
        # Por qué falta? (Si es incompleto o parcial)
        faltantes = []
        if cantidad_armable < pedido:
            for pos, cant_req in requerimientos.items():
                req_total_actual = cant_req * 1 # check para 1 unidad más
                stock_base = inv_disponible.get(pos, 0.0)
                
                patron_sub = re.compile(f'^{re.escape(pos)}[-]?([a-zA-Z])$', re.IGNORECASE)
                subpiezas_disp = {}
                for k, val in inv_disponible.items():
                    m = patron_sub.match(k)
                    if m:
                        subpiezas_disp[m.group(1).upper()] = val
                        
                stock_compuesto = 0.0
                if 'A' in subpiezas_disp and 'B' in subpiezas_disp:
                    stock_compuesto = min(subpiezas_disp.values())
                    
                stock_total = stock_base + stock_compuesto
                
                if stock_total < req_total_actual:
                    if len(subpiezas_disp) > 0 and stock_base == 0:
                        faltan_partes = []
                        # Determinar que partes faltan (asumimos que mínimo requiere A y B)
                        for letra in ['A', 'B']:
                            if subpiezas_disp.get(letra, 0.0) < req_total_actual:
                                faltan_partes.append(f"{pos}-{letra}")
                        if faltan_partes:
                            faltantes.append("Falta " + " y ".join(faltan_partes))
                        else:
                            faltantes.append(f"Falta {pos}")
                    else:
                        faltantes.append(f"Falta {pos}")
        
        obs = ", ".join(faltantes) if faltantes else "OK"
        
        # Guardar la row
        dict_req = " | ".join([f"{k}:{v}" for k,v in requerimientos.items()])
        resultados_armables.append({
            "Pieza / Ensamblaje": conjunto,
            "Cant. Pedida Lista": int(pedido),
            "Cant. que se PUEDE Armar": int(cantidad_armable),
            "Estado": estado,
            "Requerimiento (W/FL)": dict_req,
            "Observación de Faltantes": obs
        })
        
    df_res = pd.DataFrame(resultados_armables)
    
    metricas = {
        'total_conjuntos': total_conjuntos,
        'completos': completos_count
    }
    
    return df_res, metricas

def resumen_maquinas_airtable(csv_file) -> pd.DataFrame:
    """
    Lee el CSV original de Airtable y agrupa el PESO total y PESO CORTADO
    por máquina asignada, basándose en las columnas booleanas (1. GUILLOTINA, etc).
    """
    df = pd.read_csv(csv_file)
    # Detectar maquinas asumiendo el prefijo número + punto (ej. "1. MÁQUINA")
    maquinas = [c for c in df.columns if hasattr(c, 'startswith') and any(c.startswith(f"{i}.") for i in range(1, 10))]
    
    resumen = []
    
    peso_col = 'PESO' if 'PESO' in df.columns else None
    peso_cortado_col = 'PESO CORTADO' if 'PESO CORTADO' in df.columns else None
    cant_chapas_col = next((c for c in df.columns if 'cant' in str(c).lower() and 'chapas' in str(c).lower()), None)
    
    if peso_col:
        df[peso_col] = pd.to_numeric(df[peso_col], errors='coerce').fillna(0)
    if peso_cortado_col:
        df[peso_cortado_col] = pd.to_numeric(df[peso_cortado_col], errors='coerce').fillna(0)
    if cant_chapas_col:
        df[cant_chapas_col] = pd.to_numeric(df[cant_chapas_col], errors='coerce').fillna(1)
        
    for maq in maquinas:
        # Airtable exporta checkboxes como "checked"
        mascara = df[maq].notna() & df[maq].astype(str).str.contains('checked', case=False, na=False)
        df_maq = df[mascara]
        
        if peso_col and cant_chapas_col:
            peso_total = (df_maq[peso_col] * df_maq[cant_chapas_col]).sum()
        else:
            peso_total = df_maq[peso_col].sum() if peso_col else 0.0
            
        peso_cortado = df_maq[peso_cortado_col].sum() if peso_cortado_col else 0.0
        
        if peso_total > 0 or peso_cortado > 0:
            avance = (peso_cortado / peso_total * 100) if peso_total > 0 else 0.0
            resumen.append({
                "Máquina": maq.split(".", 1)[-1].strip() if "." in maq else maq,
                "Peso Total Asignado (Kg)": round(peso_total, 2),
                "Peso Cortado (Kg)": round(peso_cortado, 2),
                "Progreso": f"{avance:.1f}%"
            })
            
    return pd.DataFrame(resumen)

def generar_recomendaciones_corte(df_resultados, pdf_dicts, csv_file_bytes=None) -> list:
    """
    Toma el resultado INCOMPLETO, extrae los componentes faltantes (W/FL),
    y escanea el dict de Lantek para rankear las hojas que deberian priorizarse
    para destrabar mayor variedad de faltantes.
    """
    import re
    from collections import defaultdict
    import pandas as pd
    
    # 1. Leer el CSV para saber qué hojas/CNCs ya están cortadas y mapear datos extra
    mapa_cnc_datos = {}
    cncs_cortados = set()
    hojas_cortadas = set()
    
    if csv_file_bytes:
        try:
            csv_file_bytes.seek(0)
            df_csv = pd.read_csv(csv_file_bytes)
            cols = list(df_csv.columns)
            
            col_cnc = next((c for c in cols if 'cnc' in str(c).lower()), None)
            col_hoja = next((c for c in cols if 'hoja' in str(c).lower()), None)
            col_esp = next((c for c in cols if 'esp' in str(c).lower()), None)
            col_cant_req = next((c for c in cols if 'cant' in str(c).lower() and 'chapa' in str(c).lower()), None)
            col_cant_cortada = next((c for c in cols if ('cortada' in str(c).lower() and 'peso' not in str(c).lower())), None)
            cols_maquinas = [c for c in cols if hasattr(c, 'startswith') and any(c.startswith(f"{i}.") for i in range(1, 10))]
            
            for _, r in df_csv.iterrows():
                k_cnc = str(r[col_cnc]).strip().upper().replace(" ", "") if col_cnc else ""
                val_hoja = str(r[col_hoja]).strip() if col_hoja else ""
                
                # Check si está cortada
                cortada = False
                try:
                    cant_req = float(r[col_cant_req]) if col_cant_req and pd.notna(r[col_cant_req]) else 1.0
                    cant_cortada = float(r[col_cant_cortada]) if col_cant_cortada and pd.notna(r[col_cant_cortada]) else 0.0
                    if cant_cortada >= cant_req and cant_req > 0:
                        cortada = True
                except (ValueError, TypeError):
                    pass
                
                if cortada:
                    if k_cnc and k_cnc != 'NAN':
                        cncs_cortados.add(k_cnc)
                    if val_hoja and val_hoja.lower() != 'nan':
                        hojas_cortadas.add(val_hoja)
                
                # Mapeo de datos para la UI
                val_esp = str(r[col_esp]).strip() if col_esp else "S/D"
                maq_asignada = "S/D"
                for maq in cols_maquinas:
                    valor = str(r[maq]).lower()
                    if 'checked' in valor or 'true' in valor or valor == '1':
                        maq_asignada = maq.split(".", 1)[-1].strip() if "." in maq else maq
                        break
                        
                entry = {
                    "hoja": val_hoja if val_hoja.lower() != 'nan' else "S/D",
                    "espesor": val_esp if val_esp.lower() != 'nan' else "S/D",
                    "maquina": maq_asignada
                }
                if k_cnc and k_cnc != 'NAN':
                    mapa_cnc_datos[k_cnc] = entry
                    # También indexar por número de hoja como fallback
                    # Extraer número entre paréntesis si existe: "SID-4(49)" -> "49"
                    num_match = re.search(r'\((\d+)\)', k_cnc)
                    if num_match:
                        mapa_cnc_datos[num_match.group(1)] = entry
                # Indexar directamente por val_hoja también (ej: "33", "49")
                if val_hoja and val_hoja.lower() not in ('nan', 's/d', ''):
                    mapa_cnc_datos[val_hoja.strip()] = entry
        except Exception as e:
            print("Error procesando CSV en recomendaciones:", e)

    faltantes_detectados = set()
    # Para cada conjunto: el set COMPLETO de piezas que le faltan (normalizadas a base)
    conjunto_todos_faltantes = defaultdict(set)
    
    for _, row in df_resultados.iterrows():
        estado = str(row["Estado"]).upper()
        conjunto_nombre = str(row["Pieza / Ensamblaje"])
        if "INCOMPLE" in estado or "PARCIAL" in estado:
            obs = str(row["Observación de Faltantes"])
            if obs and obs != "OK":
                partes = obs.split(",")
                for p in partes:
                    matches = re.findall(r'((?:W|WB|FL|ST|PC|PL)-\S+)', p)
                    for m in matches:
                        pieza = m.strip().upper()
                        if pieza:
                            # Normalizar siempre a la pieza BASE (sin sufijo -A, -B)
                            base_match = re.match(r'^([A-Z]+-\d+)', pieza)
                            base = base_match.group(1) if base_match else pieza
                            faltantes_detectados.add(base)
                            conjunto_todos_faltantes[conjunto_nombre].add(base)

    if not faltantes_detectados:
        return []
        
    hojas_a_cortar = defaultdict(set)
    
    for p_dict in pdf_dicts:
        for k_ref, inventario_chapa in p_dict.items():
            k_orig = str(k_ref).strip()
            if str(k_orig).startswith("_"): continue

            # IGNORAR SI YA ESTÁ CORTADA
            k_limpio = k_orig.upper().replace(" ", "")
            if k_limpio in cncs_cortados or k_orig in hojas_cortadas:
                continue

            for pieza_chapa in inventario_chapa.keys():
                pieza_chapa_str = str(pieza_chapa).upper()
                # Normalizar a base para comparar
                base_m = re.match(r'^([A-Z]+-\d+)', pieza_chapa_str)
                base_chapa = base_m.group(1) if base_m else pieza_chapa_str
                if base_chapa in faltantes_detectados:
                    hojas_a_cortar[k_orig].add(base_chapa)

    # Colapsar aliases duplicados (mismo set de piezas → conservar nombre más largo)
    hojas_unicas_dict = {}
    for ref_name, piezas_set in hojas_a_cortar.items():
        llave = frozenset(piezas_set)
        if llave not in hojas_unicas_dict:
            hojas_unicas_dict[llave] = ref_name
        else:
            if len(ref_name) > len(hojas_unicas_dict[llave]):
                hojas_unicas_dict[llave] = ref_name

    # --- Criterio real: ¿cuántos conjuntos quedan 100% ARMABLES si se corta esta hoja? ---
    # Un conjunto "completa" si TODOS sus faltantes están cubiertos por las piezas de la hoja.
    candidatas = []
    for k_froz, nombre_ref in hojas_unicas_dict.items():
        piezas_hoja = set(k_froz)  # bases que esta hoja provee

        completan = []   # conjuntos que quedan 100% armables
        ayudan = []      # conjuntos que mejoran pero no completan

        for conj, sus_faltantes in conjunto_todos_faltantes.items():
            if not sus_faltantes:
                continue
            if sus_faltantes.issubset(piezas_hoja):
                # Esta hoja cubre TODOS los faltantes → conjunto completado
                completan.append(conj)
            elif sus_faltantes & piezas_hoja:
                # Esta hoja cubre ALGUNOS faltantes → ayuda parcialmente
                ayudan.append(conj)

        if completan or ayudan:
            candidatas.append((nombre_ref, piezas_hoja, completan, ayudan))

    # Ordenar: primero por completan DESC, luego por ayudan DESC
    candidatas.sort(key=lambda x: (len(x[2]), len(x[3])), reverse=True)

    # Armar filas de resultado
    recomendaciones = []
    for k, v, completan, ayudan in candidatas[:10]:
        k_limpio = k.upper().replace(" ", "")
        datos = (
            mapa_cnc_datos.get(k_limpio)
            or mapa_cnc_datos.get(k.strip())
            or {"hoja": "S/D", "espesor": "S/D", "maquina": "S/D"}
        )

        completan_str = ", ".join(sorted(completan)) if completan else "-"
        ayudan_str = ", ".join(sorted(ayudan)) if ayudan else "-"

        recomendaciones.append({
            "Nº Hoja": datos["hoja"],
            "Máquina": datos["maquina"],
            "Espesor": datos["espesor"],
            "Chapa Lantek (Prog CNC / Ref)": k,
            "✅ Completa (cant.)": len(completan),
            "⚠️ Ayuda parcial (cant.)": len(ayudan),
            "Se Podría Armar (100%)": completan_str,
            "Ayuda a Completar": ayudan_str,
            "Piezas en la Plancha": ", ".join(sorted(list(v))[:15])
        })

    return recomendaciones


def resumen_chapas_nesting(csv_file) -> pd.DataFrame:
    """
    Lee el CSV de Airtable y construye un resumen de las chapas que componen
    el nesting, agrupado por espesor (de menor a mayor).

    Columnas esperadas en el CSV: Espesor (mm), LARGO, ANCHO, PESO, Cant. Chapas
    - PESO ya es el peso de UNA chapa en Airtable.
    - Peso Total = PESO × Cant. Chapas

    Retorna un DataFrame con:
        Espesor (mm) | Largo (mm) | Ancho (mm) | Cant. Chapas | Peso Unit. (Kg) | Peso Total (Kg)
    agrupado por filas idénticas (mismo espesor + largo + ancho) y ordenado por espesor asc.
    """
    try:
        if hasattr(csv_file, 'seek'):
            csv_file.seek(0)
        df = pd.read_csv(csv_file)
    except Exception:
        return pd.DataFrame()

    # Detectar columnas robustamente (case-insensitive)
    def find_col(df, *keywords):
        for col in df.columns:
            col_norm = str(col).lower().strip()
            if all(k in col_norm for k in keywords):
                return col
        return None

    col_esp   = find_col(df, 'esp')
    col_largo = find_col(df, 'largo')
    col_ancho = find_col(df, 'ancho')
    col_cant  = find_col(df, 'cant', 'chap') or find_col(df, 'cant')

    # Buscar columna PESO excluyendo "cortado" y "faltante"
    col_peso = next(
        (c for c in df.columns
         if 'peso' in str(c).lower()
         and 'cortado' not in str(c).lower()
         and 'faltante' not in str(c).lower()),
        None
    )

    if not all([col_esp, col_largo, col_ancho, col_peso]):
        return pd.DataFrame()  # Sin las columnas mínimas no se puede continuar

    # Conversión numérica
    df[col_esp]   = pd.to_numeric(df[col_esp],   errors='coerce')
    df[col_largo] = pd.to_numeric(df[col_largo], errors='coerce')
    df[col_ancho] = pd.to_numeric(df[col_ancho], errors='coerce')
    df[col_peso]  = pd.to_numeric(df[col_peso],  errors='coerce').fillna(0)

    if col_cant:
        df[col_cant] = pd.to_numeric(df[col_cant], errors='coerce').fillna(1)
    else:
        df['_cant_tmp'] = 1
        col_cant = '_cant_tmp'

    # Eliminar filas sin datos geométricos
    df = df.dropna(subset=[col_esp, col_largo, col_ancho])
    if df.empty:
        return pd.DataFrame()

    df['_peso_unit']  = df[col_peso]                    # peso unitario = columna PESO
    df['_peso_total'] = df[col_peso] * df[col_cant]
    df['_cant']       = df[col_cant]

    # Agrupar filas con misma geometría
    grupo = df.groupby([col_esp, col_largo, col_ancho], as_index=False).agg(
        cant_chapas=('_cant', 'sum'),
        peso_unit=('_peso_unit', 'first'),
        peso_total=('_peso_total', 'sum')
    )

    grupo = grupo.sort_values(col_esp).reset_index(drop=True)

    grupo.rename(columns={
        col_esp:       'Espesor (mm)',
        col_largo:     'Largo (mm)',
        col_ancho:     'Ancho (mm)',
        'cant_chapas': 'Cant. Chapas',
        'peso_unit':   'Peso Unit. (Kg)',
        'peso_total':  'Peso Total (Kg)',
    }, inplace=True)

    # Formateo visual
    grupo['Espesor (mm)']    = grupo['Espesor (mm)'].apply(lambda x: f"{x:g}")
    grupo['Largo (mm)']      = grupo['Largo (mm)'].apply(lambda x: f"{int(x):,}")
    grupo['Ancho (mm)']      = grupo['Ancho (mm)'].apply(lambda x: f"{int(x):,}")
    grupo['Cant. Chapas']    = grupo['Cant. Chapas'].apply(lambda x: int(x))
    grupo['Peso Unit. (Kg)'] = grupo['Peso Unit. (Kg)'].apply(lambda x: f"{x:,.1f}")
    grupo['Peso Total (Kg)'] = grupo['Peso Total (Kg)'].apply(lambda x: f"{x:,.1f}")

    return grupo

