import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import traceback

# ==========================================
# 1. CONFIGURACIÓN INICIAL DE LA APP
# ==========================================
st.set_page_config(page_title="Suite Contable", page_icon="💼", layout="wide")

# --- VARIABLES DE SESIÓN GLOBALES ---
if 'contador_usos' not in st.session_state: st.session_state.contador_usos = 0
if 'asiento_generado' not in st.session_state: st.session_state.asiento_generado = None
if 'fecha_asiento' not in st.session_state: st.session_state.fecha_asiento = ""
if 'proveedores_faltantes' not in st.session_state: st.session_state.proveedores_faltantes = []

LIMITE_GRATIS = 3

# --- ESTILOS NATIVOS (COMPATIBLES CON DARK MODE DE STREAMLIT) ---
st.markdown("""
    <style>
    h1 { color: #4DA8DA !important; font-family: 'Arial', sans-serif; font-weight: 800; text-align: center; text-transform: uppercase; letter-spacing: 1px; padding-bottom: 20px; }
    h2, h3, h4 { color: #4DA8DA !important; }
    .stFileUploader { border: 2px dashed #4DA8DA !important; border-radius: 8px; padding: 10px; }
    
    .counter-container { border-left: 5px solid #4DA8DA; border-radius: 4px; padding: 15px; text-align: center; margin-bottom: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .counter-number { font-weight: 800; font-size: 24px; color: #E53E3E; }

    /* Pestañas (Tamaño Grande) */
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 55px; white-space: pre-wrap; border-radius: 4px 4px 0px 0px; gap: 1px; padding-top: 10px; padding-bottom: 10px; font-weight: bold; font-size: 22px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNCIONES Y MOTORES (MÓDULO BANCOS)
# ==========================================
def limpiar_monto_ar(texto):
    if not texto or pd.isna(texto): return 0.0
    t = str(texto).replace('$', '').replace(' ', '').replace('"', '').replace('+', '')
    es_negativo = '-' in t
    t = t.replace('-', '')
    
    if len(t) >= 3 and t[-3] in ['.', ',']:
        t = t[:-3].replace('.', '').replace(',', '') + '.' + t[-2:]
    else:
        t = t.replace('.', '').replace(',', '.')
        
    try: return -float(t) if es_negativo else float(t)
    except: return 0.0

def aplicar_diccionario_final(df, df_dic_manual=None):
    reglas = {
        'DEP.EFVO.AUTOSERVICIO TICKET': 'CAJA',
        'COMISION CHEQUE PAGADO POR CLEARING': 'GASTOS Y COMISIONES BANCARIAS',
        'COMISION SERVICIO DE CUENTA': 'GASTOS Y COMISIONES BANCARIAS',
        'COM. GESTION TRANSF.FDOS ENTRE BCOS': 'GASTOS Y COMISIONES BANCARIAS',
        'COM. GESTION TRANSF.FDOS': 'GASTOS Y COMISIONES BANCARIAS',
        'COM. DEPOSITO DE CHEQUE': 'GASTOS Y COMISIONES BANCARIAS',
        'COMISION CHEQUE PAGADO POR': 'GASTOS Y COMISIONES BANCARIAS',
        'N/D MANTENIMIENTO MENSUAL': 'GASTOS Y COMISIONES BANCARIAS',
        'INTERESES SOBRE SALDOS': 'GASTOS Y COMISIONES BANCARIAS',
        'DEVOLUCION COMISIONES POR': 'GASTOS Y COMISIONES BANCARIAS',
        'INTERESES SOBRE SALDOS DEUDORES': 'INT. BANCARIOS',
        'IVA': 'IVA CREDITO FISCAL',
        'IMP. DEB. LEY 25.413': 'LEY 25.413',
        'IMP. DEB. LEY 25413': 'LEY 25.413',
        'IMP. CRE. LEY 25.413': 'LEY 25.413',
        'DEV.IMP.DEB.LEY 25.413': 'LEY 25.413',
        'N/D DBCR 25413 S/CR TASA GRAL': 'LEY 25.413',
        'N/D FV IMPDBCR 25413 S/DB TASA GRAL': 'LEY 25.413',
        'N/D FV IMPDBCR 25413 S/CR TASA GRA': 'LEY 25.413',
        'N/D DBCR 25413 S/DB TASA GRAL': 'LEY 25.413',
        'PERCEP. IVA': 'PERCEPCIONES IVA',
        'RETENCION IVA PERCEPCION': 'PERCEPCIONES IVA',
        'RESCATE FIMA': 'RESCATE FIMA',
        'ING. BRUTOS S/ CRED': 'SIRCREB',
        'REG.RECAU.SIRCREB': 'SIRCREB',
        'N/D IIBB SANTA FE SIRCREB': 'SIRCREB',
        'N/D FV IIBB SANTA FE SIRCREB': 'SIRCREB',
        'SUSCRIPCION FIMA': 'SUSCRIPCION FIMA',
        'NAVE - VENTA CON TARJETA': 'TARJETAS A COBRAR',
        'LIQ COMER PRISMA': 'TARJETAS A COBRAR',
        'NAVE PAGO': 'TARJETAS A COBRAR'
    }
    if df_dic_manual is not None:
        for _, r in df_dic_manual.iterrows():
            reglas[str(r.iloc[0]).upper().strip()] = str(r.iloc[1]).strip()
            
    dic_ord = {k: v for k, v in sorted(reglas.items(), key=lambda x: len(x[0]), reverse=True)}
    df['Imputación'] = "✨ A Clasificar"
    for idx, row in df.iterrows():
        concepto = str(row['Concepto']).upper()
        if "SALDO INICIAL" in concepto:
            df.at[idx, 'Imputación'] = ""
            continue
        for clave, cuenta in dic_ord.items():
            if clave in concepto:
                df.at[idx, 'Imputación'] = cuenta
                break
    return df

def motor_galicia_tradicional(pdf):
    texto_p1 = pdf.pages[0].extract_text() or ""
    texto_completo = "\n".join([p.extract_text() or "" for p in pdf.pages])
    lineas = texto_completo.split('\n')
    patron_f = re.compile(r'^\s*(\d{2}/\d{2}/\d{2,4})')
    patron_m = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d{2}')
    movs, monto_ini = [], 0.0
    
    match_ini = re.search(r'saldo[\s\n]*inicial[\s\n]*\$?[\s\n]*(-?\d{1,3}(?:\.\d{3})*,\d{2})', texto_completo.lower())
    if match_ini: 
        monto_ini = limpiar_monto_ar(match_ini.group(1))
    else:
        todos_montos = patron_m.findall(texto_p1)
        if todos_montos: monto_ini = limpiar_monto_ar(todos_montos[0])
            
    last_s = monto_ini
    for linea in lineas:
        match_f = patron_f.match(linea.strip())
        if match_f:
            f = match_f.group(1)
            ms = patron_m.findall(linea)
            if ms and "SALDO" not in linea.upper():
                val = limpiar_monto_ar(ms[-2] if len(ms) >= 2 else ms[0])
                sa = limpiar_monto_ar(ms[-1])
                
                dif = round(sa - last_s, 2)
                if abs(dif - val) > 0.01 and abs(dif + val) < 0.01:
                    val = -val 
                    
                conc = re.sub(r'\s+', ' ', linea.replace(f, "")).strip()
                for m in ms: conc = conc.replace(m, "")
                movs.append({'Fecha': f, 'Concepto': conc, 'Debitos': abs(val) if val < 0 else 0.0, 'Creditos': val if val > 0 else 0.0, 'Neto': val, 'Saldo': sa})
                last_s = sa
    
    df = pd.DataFrame(movs)
    if not df.empty:
        fila_ini = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': monto_ini}])
        df = pd.concat([fila_ini, df], ignore_index=True)
    return {"Resumen Galicia": df}

def motor_galicia_office(pdf):
    texto_layout = "\n".join([p.extract_text(layout=True) or "" for p in pdf.pages])
    lineas = texto_layout.split('\n')
    patron_fecha = re.compile(r'^\s*"?(\d{2}/\d{2}/\d{2,4})')
    patron_monto = re.compile(r'([+-]?\s?\$?\s?\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)')
    movs_crudos = []
    mov_actual = None
    
    for linea in lineas:
        if not linea.strip() or "office" in linea.lower(): continue
        match_f = patron_fecha.search(linea.strip())
        es_nueva_fila = match_f and match_f.start() < 15
        montos = patron_monto.findall(linea)
        if es_nueva_fila:
            if mov_actual: movs_crudos.append(mov_actual)
            mov_actual = {'Fecha': match_f.group(1), 'Concepto': re.sub(r'\s+', ' ', linea.replace(match_f.group(1), "")).strip().replace('"', ''), 'Montos_Raw': montos}
        elif mov_actual:
            if montos: mov_actual['Montos_Raw'].extend(montos)
            mov_actual['Concepto'] += " " + re.sub(r'\s+', ' ', linea).strip()
    if mov_actual: movs_crudos.append(mov_actual)
    
    movimientos = []
    for i, mov in enumerate(movs_crudos):
        montos = mov['Montos_Raw']
        if len(montos) >= 2: m_mov_str, m_saldo_str = montos[-2], montos[-1]
        elif len(montos) == 1: m_mov_str, m_saldo_str = montos[0], None
        else: continue
        val_mov = abs(limpiar_monto_ar(m_mov_str))
        val_saldo = limpiar_monto_ar(m_saldo_str) if m_saldo_str else None
        neto = 0.0
        if '-' in m_mov_str: neto = -val_mov
        elif '+' in m_mov_str: neto = val_mov
        else:
            if val_saldo is not None and i > 0 and movimientos[-1]['Saldo'] is not None:
                delta = val_saldo - movimientos[-1]['Saldo']
                neto = val_mov if abs(delta - val_mov) < 0.10 else -val_mov
            else: neto = val_mov if any(x in mov['Concepto'].upper() for x in ["DEP", "CREDITO", "RESCATE", "VENTA", "TICKET"]) else -val_mov
        movimientos.append({'Fecha': mov['Fecha'], 'Concepto': mov['Concepto'], 'Debitos': abs(neto) if neto < 0 else 0.0, 'Creditos': neto if neto > 0 else 0.0, 'Neto': neto, 'Saldo': val_saldo})
    
    df = pd.DataFrame(movimientos)
    if not df.empty:
        s_pdf_0 = df.iloc[0]['Saldo'] if df.iloc[0]['Saldo'] is not None else 0.0
        monto_inicial = s_pdf_0 - df.iloc[0]['Neto']
        df['Saldo'] = df['Saldo'].fillna(monto_inicial + df['Neto'].cumsum())
        fila_ini = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': monto_inicial}])
        df = pd.concat([fila_ini, df], ignore_index=True)
    return {"Resumen Galicia Office": df}

def motor_macro(pdf):
    cuentas = {}
    patron_nro = re.compile(r'(?:NRO\.|CUENTA|CTA)\s*[:]?\s*([\d-]+)', re.IGNORECASE)
    patron_f = re.compile(r'(\d{2}/\d{2}/\d{2,4})')
    patron_m = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d{2}')
    cta_actual = "Cuenta Principal"
    
    for page in pdf.pages:
        texto = page.extract_text(layout=True) or ""
        for linea in texto.split('\n'):
            linea_u = linea.upper()
            
            match_c = patron_nro.search(linea_u)
            if match_c:
                cta_actual = f"Cuenta {match_c.group(1)}"
                if cta_actual not in cuentas:
                    cuentas[cta_actual] = {"movs": [], "ini": 0.0, "last_saldo": 0.0, "ini_set": False}
                continue
                
            if cta_actual not in cuentas: cuentas[cta_actual] = {"movs": [], "ini": 0.0, "last_saldo": 0.0, "ini_set": False}
            
            if "SALDO ULTIMO EXTRACTO" in linea_u or "SALDO ANTERIOR" in linea_u:
                ms = patron_m.findall(linea)
                if ms:
                    val = limpiar_monto_ar(ms[-1])
                    cuentas[cta_actual]["ini"] = val
                    cuentas[cta_actual]["last_saldo"] = val
                    cuentas[cta_actual]["ini_set"] = True
                continue

            match_f = patron_f.search(linea)
            if match_f and not any(x in linea_u for x in ["TOTAL", "SALDO FINAL", "HOJA"]):
                ms = patron_m.findall(linea)
                if len(ms) >= 2:
                    saldo_pdf = limpiar_monto_ar(ms[-1])
                    neto = round(saldo_pdf - cuentas[cta_actual]["last_saldo"], 2)
                    if abs(neto) < 0.01: continue 
                    conc = re.sub(r'\s+', ' ', linea.replace(match_f.group(1), "")).strip()
                    for m in ms: conc = conc.replace(m, "")
                    cuentas[cta_actual]["movs"].append({'Fecha': match_f.group(1), 'Concepto': conc, 'Debitos': abs(neto) if neto < 0 else 0.0, 'Creditos': neto if neto > 0 else 0.0, 'Neto': neto, 'Saldo': saldo_pdf})
                    cuentas[cta_actual]["last_saldo"] = saldo_pdf
    
    final_dfs = {}
    for cta, data in cuentas.items():
        if data["movs"]:
            df = pd.DataFrame(data["movs"])
            fila_ini = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': data["ini"]}])
            final_dfs[cta] = pd.concat([fila_ini, df], ignore_index=True)
    return final_dfs

def motor_mercado_pago(pdf):
    texto_layout = "\n".join([p.extract_text(layout=True) or "" for p in pdf.pages])
    m_ini = re.search(r"Saldo inicial[\s:]*\$?\s*(-?[\d\.]+[,.]\d{2})", texto_layout, re.IGNORECASE)
    s_ini = limpiar_monto_ar(m_ini.group(1)) if m_ini else 0.0
    movs = []
    mov_actual = None
    for linea in texto_layout.split('\n'):
        if not linea.strip() or "DETALLE DE MOVIMIENTOS" in linea.upper() or "SALDO INICIAL" in linea.upper() or "ENTRADAS:" in linea.upper() or "SALIDAS:" in linea.upper():
            continue
        m_f = re.search(r'^(\d{2}-\d{2}-\d{4})', linea.strip())
        if m_f:
            if mov_actual:
                movs.append(mov_actual)
            fecha = m_f.group(1)
            resto = linea.replace(fecha, "", 1)
            montos = re.findall(r'-?\s?\$?\s*[\d\.]+[,.]\d{2}', resto)
            ids = re.findall(r'\b\d{10,}\b', resto)
            concepto = resto
            for m in montos: 
                concepto = concepto.replace(m, "")
            for i in ids: 
                concepto = concepto.replace(i, "")
            val_str = montos[-2] if len(montos) >= 2 else (montos[0] if montos else "0")
            mov_actual = {
                "Fecha": fecha.replace('-', '/'),
                "Concepto": concepto.strip(),
                "Valor_str": val_str
            }
        elif mov_actual:
            l_clean = re.sub(r'\b\d{10,}\b', '', linea).strip()
            l_clean = re.sub(r'-?\s?\$?\s*[\d\.]+[,.]\d{2}', '', l_clean).strip()
            if l_clean and "Mercado Libre" not in l_clean and "www." not in l_clean and not re.match(r'^\d+/\d+', l_clean):
                mov_actual["Concepto"] += " " + l_clean
    if mov_actual:
        movs.append(mov_actual)
    movimientos_finales = []
    for m in movs:
        val = limpiar_monto_ar(m["Valor_str"])
        conc = re.sub(r'\s+', ' ', m["Concepto"]).strip()
        movimientos_finales.append({
            "Fecha": m["Fecha"],
            "Concepto": conc,
            "Debitos": abs(val) if val < 0 else 0.0,
            "Creditos": val if val > 0 else 0.0,
            "Neto": val
        })
    df = pd.DataFrame(movimientos_finales)
    if not df.empty:
        df['Saldo'] = s_ini + df['Neto'].cumsum()
        fi = pd.DataFrame([{"Fecha": df.iloc[0]["Fecha"], "Concepto": "SALDO INICIAL", "Debitos": 0.0, "Creditos": 0.0, "Neto": 0.0, "Saldo": s_ini}])
        return {"Mercado Pago": pd.concat([fi, df], ignore_index=True)}
    return {}

def motor_icbc(pdf):
    texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
    m_ini = re.search(r"SALDO\s+ULTIMO\s+EXTRACTO.*?([\d\.]+,\d{2}-?)", texto.upper().replace('\n', ' '))
    
    def limpiar_icbc(t):
        neg = '-' in str(t)
        num = str(t).replace('-', '').replace('.', '').replace(',', '.')
        return -float(num) if neg else float(num)

    s_ini = limpiar_icbc(m_ini.group(1)) if m_ini else 0.0
    movs = []
    for line in texto.split('\n'):
        m_f = re.match(r'^(\d{2}-\d{2})\s+(.+)', line.strip())
        if m_f:
            f, resto = m_f.groups()
            ms = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}-?', resto)
            if ms:
                val = limpiar_icbc(ms[0])
                conc = resto.replace(ms[0], '').strip()
                conc = re.sub(r'\b\d{4}\b', '', conc).strip()
                movs.append({'Fecha': f.replace('-', '/'), 'Concepto': conc, 'Debitos': abs(val) if val < 0 else 0.0, 'Creditos': val if val > 0 else 0.0, 'Neto': val})
    
    df = pd.DataFrame(movs)
    if not df.empty:
        df['Saldo'] = s_ini + df['Neto'].cumsum()
        fi = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': s_ini}])
        return {"Resumen ICBC": pd.concat([fi, df], ignore_index=True)}
    return {}

# -----------------------------------------------------
# MOTOR CREDICOOP ARREGLADO DEFINITIVO
# -----------------------------------------------------
def motor_credicoop(pdf):
    movs = []
    s_ini = 0.0
    ini_set = False
    
    # Palabras a ignorar para limpiar la lectura
    skip_phrases = [
        "BANCO CREDICOOP", "CONTACTO TELEFONIC", "RESUMEN:", "DEBITO DIRECTO", 
        "CBU DE SU", "CONTINUA EN", "VIENE DE PAGINA", "PAGINA ", "CUENTA CORRIENTE", 
        "CALIDAD DE SERVICIOS", "WWW.BANCOCREDICOOP", "LIQUIDACION DE INTERESES", 
        "TNA", "TEA", "CFTEA", "PERCIBIDO DEL", "DENOMINACION", "IMPUESTO LEY", 
        "PERCIBIDO", "ALICUOTA", "F E C H A", "FECHA", "COMBTE", "DESCRIPCION", 
        "DEBITO", "CREDITO", "SALDO", "SALDO ANTERIOR", "TOTALES"
    ]
    
    for page in pdf.pages:
        words = page.extract_words()
        if not words: continue
        
        # Agrupación precisa de renglones
        lineas = {}
        for w in words:
            y = round(w['top'])
            matched_y = next((ky for ky in lineas.keys() if abs(ky - y) <= 3), None)
            if matched_y is not None:
                lineas[matched_y].append(w)
            else:
                lineas[y] = [w]
                
        for y in sorted(lineas.keys()):
            f_w = sorted(lineas[y], key=lambda x: x['x0'])
            txt = " ".join([w['text'] for w in f_w])
            txt_upper = txt.upper()
            
            # --- ARREGLO SALDO INICIAL ---
            # Quitamos los espacios para que un signo negativo separado no se pierda
            if "SALDO" in txt_upper and "ANTERIOR" in txt_upper and not ini_set:
                txt_junto = txt.replace(" ", "")
                ms = re.findall(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})', txt_junto)
                if ms:
                    s_ini = limpiar_monto_ar(ms[-1])
                    ini_set = True
                continue
                
            if any(sp in txt_upper for sp in skip_phrases):
                continue
            
            # --- ARREGLO RENGLONES PEGADOS ---
            # Buscamos la fecha explicitamente revisando las primeras palabras
            fecha = None
            for p in f_w:
                if re.match(r'^\d{2}/\d{2}/\d{2,4}$', p['text']) and p['x0'] < 100:
                    fecha = p['text']
                    break
                    
            # Si encontramos una fecha nueva, armamos un movimiento
            if fecha:
                deb, cre = 0.0, 0.0
                cp = []
                for p in f_w:
                    if p['text'] == fecha: continue
                    
                    if re.match(r'^-?\d{1,3}(?:\.\d{3})*,\d{2}$', p['text']) and p['x0'] > 300:
                        v = limpiar_monto_ar(p['text'])
                        # Calculamos el centro de la palabra para que las cifras grandes no invadan columnas
                        cx = (p['x0'] + p['x1']) / 2
                        if cx < 430: 
                            deb = abs(v)
                        elif cx < 510: 
                            cre = v
                    else:
                        if not re.match(r'^\d{5,8}$', p['text']): # Ignorar IDs de operación
                            cp.append(p['text'])
                            
                if deb > 0 or cre > 0 or cp:
                    movs.append({
                        "Fecha": fecha, 
                        "Concepto": " ".join(cp).strip(), 
                        "Debitos": deb, 
                        "Creditos": cre, 
                        "Neto": cre - deb
                    })
                    
            # Si NO hay fecha, es continuación del renglón de arriba
            elif movs:
                deb, cre = 0.0, 0.0
                cp = []
                for p in f_w:
                    if re.match(r'^-?\d{1,3}(?:\.\d{3})*,\d{2}$', p['text']) and p['x0'] > 300:
                        v = limpiar_monto_ar(p['text'])
                        cx = (p['x0'] + p['x1']) / 2
                        if cx < 430: 
                            deb = abs(v)
                        elif cx < 510: 
                            cre = v
                    else:
                        if not re.match(r'^\d{5,8}$', p['text']):
                            cp.append(p['text'])
                            
                if cp:
                    movs[-1]["Concepto"] += " " + " ".join(cp).strip()
                if deb > 0:
                    movs[-1]["Debitos"] += deb
                    movs[-1]["Neto"] -= deb
                if cre > 0:
                    movs[-1]["Creditos"] += cre
                    movs[-1]["Neto"] += cre
                    
    # Limpiamos basuras que hayan quedado con neto 0 y sin concepto claro
    movs = [m for m in movs if m["Neto"] != 0 or len(m["Concepto"]) > 5]
    
    df = pd.DataFrame(movs)
    if not df.empty:
        df['Saldo'] = s_ini + df['Neto'].cumsum()
        fi = pd.DataFrame([{"Fecha": df.iloc[0]["Fecha"], "Concepto": "SALDO INICIAL", "Debitos": 0.0, "Creditos": 0.0, "Neto": 0.0, "Saldo": s_ini}])
        return {"Resumen Credicoop": pd.concat([fi, df], ignore_index=True)}
    return {}

def motor_frances(pdf):
    texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
    m_ini = re.search(r"SALDO\s+ANTERIOR.*?([\d\.]+,\d{2})", texto.upper())
    s_ini = limpiar_monto_ar(m_ini.group(1)) if m_ini else 0.0
    movs, last_s = [], s_ini
    for line in texto.split('\n'):
        m_f = re.match(r'^(\d{2}/\d{2})\s+(.+)', line.strip())
        if m_f:
            fecha_str = m_f.group(1)
            resto_linea = m_f.group(2)
            if fecha_str == "00/00" or "SIN MOVIMIENTOS" in resto_linea.upper():
                continue
            ms = re.findall(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})', line)
            if ms:
                sa = limpiar_monto_ar(ms[-1])
                dif = round(sa - last_s, 2)
                if abs(dif) > 0.01:
                    movs.append({"Fecha": fecha_str, "Concepto": line.replace(fecha_str, "").replace(ms[-1], "").strip(), "Debitos": abs(dif) if dif < 0 else 0.0, "Creditos": dif if dif > 0 else 0.0, "Neto": dif, "Saldo": sa})
                    last_s = sa
    df = pd.DataFrame(movs)
    if not df.empty:
        fi = pd.DataFrame([{"Fecha": df.iloc[0]["Fecha"], "Concepto": "SALDO INICIAL", "Debitos": 0.0, "Creditos": 0.0, "Neto": 0.0, "Saldo": s_ini}])
        return {"Resumen BBVA": pd.concat([fi, df], ignore_index=True)}
    return {}

# ==========================================
# 3. FUNCIONES Y MOTORES (MÓDULO COMPRAS)
# ==========================================
def limpiar_memoria():
    st.session_state.asiento_generado = None
    st.session_state.fecha_asiento = ""
    st.session_state.proveedores_faltantes = []

# ==========================================
# 4. INTERFAZ GRÁFICA UNIFICADA
# ==========================================
st.markdown("<h1>💼 Suite Contable</h1>", unsafe_allow_html=True)

col_prog, _ = st.columns([1, 2])
with col_prog:
    st.write(f"📊 **Uso de cortesía:** {st.session_state.contador_usos} de {LIMITE_GRATIS}")
    st.progress(st.session_state.contador_usos / LIMITE_GRATIS)

if st.session_state.contador_usos >= LIMITE_GRATIS:
    st.markdown("---")
    _, col_centro, _ = st.columns([1, 2, 1])
    with col_centro:
        st.markdown("""
            <div class="donation-card">
                <h2>¡Límite Alcanzado! ❤️</h2>
                <p>Este bot es gratuito. Si te sirve, podés colaborar para mantenerlo online.</p>
                <hr style="border: 0.5px solid #333; margin: 20px 0;">
                <p><b>☕ DATOS PARA COLABORAR</b></p>
                <p style='color: #00FFCC; font-weight: bold; font-size: 20px;'>Alias: vg1990.mp</p>
                <p>Titular: Verónica</p>
            </div>
            <br>
        """, unsafe_allow_html=True)
        if st.button("🔄 Reiniciar App", use_container_width=True):
            st.session_state.contador_usos = 0
            limpiar_memoria()
            st.rerun()
    st.stop()

# --- PESTAÑAS ---
tab1, tab2 = st.tabs(["🏦 Bancos a Excel", "🛒 Compras a Asientos"])

with tab1:
    st.markdown("### 📄 Procesador de Extractos Bancarios")
    banco_sel = st.selectbox("Seleccionar Banco:", ["Macro", "Galicia", "Mercado Pago", "ICBC", "Credicoop", "BBVA Francés"])
    c1, c2 = st.columns(2)
    with c1: 
        archivo_pdf = st.file_uploader("Subir PDF del banco", type="pdf")
    with c2: 
        archivo_dic_bancos = st.file_uploader("Subir Diccionario (Opcional)", type="xlsx", key="dic_bancos")
        
    if archivo_pdf:
        try:
            with pdfplumber.open(archivo_pdf) as pdf:
                if banco_sel == "Galicia":
                    texto_p1 = (pdf.pages[0].extract_text() or "").upper()
                    dict_dfs = motor_galicia_office(pdf) if "OFFICE" in texto_p1 or '","' in texto_p1 else motor_galicia_tradicional(pdf)
                elif banco_sel == "Macro": dict_dfs = motor_macro(pdf)
                elif banco_sel == "ICBC": dict_dfs = motor_icbc(pdf)
                elif banco_sel == "Mercado Pago": dict_dfs = motor_mercado_pago(pdf)
                elif banco_sel == "Credicoop": dict_dfs = motor_credicoop(pdf)
                else: dict_dfs = motor_frances(pdf)
                
                if dict_dfs:
                    st.session_state.contador_usos += 1
                    df_manual_b = pd.read_excel(archivo_dic_bancos) if archivo_dic_bancos else None
                    
                    cta_sel = st.selectbox("Cuenta detectada:", list(dict_dfs.keys()))
                    df_final_b = aplicar_diccionario_final(dict_dfs[cta_sel], df_manual_b)
                    
                    cols_visibles_b = ['Fecha', 'Concepto', 'Debitos', 'Creditos', 'Saldo', 'Imputación']
                    df_mostrar_b = df_final_b[[c for c in cols_visibles_b if c in df_final_b.columns]]
                    
                    st.dataframe(df_mostrar_b.style.format({'Debitos': '{:,.2f}', 'Creditos': '{:,.2f}', 'Saldo': '{:,.2f}'}), use_container_width=True)
                    
                    output_b = io.BytesIO()
                    with pd.ExcelWriter(output_b) as writer:
                        for k, df_b in dict_dfs.items():
                            df_exp_b = aplicar_diccionario_final(df_b, df_manual_b)
                            df_exp_b[[c for c in cols_visibles_b if c in df_exp_b.columns]].to_excel(writer, index=False, sheet_name=str(k)[:31].replace("/", "-"))
                    st.download_button("🚀 DESCARGAR EXCEL", output_b.getvalue(), f"conciliacion_{banco_sel.lower()}.xlsx")
                else:
                    st.warning("No se encontraron movimientos. Verificá que seleccionaste el banco correcto.")
        except Exception as e:
            st.error(f"Error técnico en Bancos: {e}")

with tab2:
    st.markdown("### 🛒 Generador de Asientos de Compras")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1: 
        archivo_compras = st.file_uploader("1. Subir compras (.xlsx, .xls)", type=["xlsx", "xls"], key="compras", on_change=limpiar_memoria)
    with col_c2: 
        archivo_dicc_compras = st.file_uploader("2. Subir diccionario de compras", type=["xlsx"], key="diccionario_com", on_change=limpiar_memoria)

    if archivo_compras and archivo_dicc_compras:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔥 GENERAR ASIENTO AUTOMÁTICO", key="btn_compras"):
            with st.spinner('🚀 Procesando facturas y cuadrando IVA...'):
                try:
                    df = pd.read_excel(archivo_compras)
                    df_dicc = pd.read_excel(archivo_dicc_compras)
                    
                    df.columns = df.columns.astype(str).str.strip().str.upper()
                    df_dicc.columns = df_dicc.columns.astype(str).str.strip().str.upper()
                    
                    for v in ['PROVEEDOR', 'PROVEEDORES', 'RAZON SOCIAL', 'RAZÓN SOCIAL', 'NOMBRE', 'DENOMINACION', 'DENOMINACIÓN EMISOR']:
                        if v in df.columns: df.rename(columns={v: 'PROVEEDOR_KEY'}, inplace=True); break
                    for v in ['PROVEEDOR', 'PROVEEDORES', 'RAZON SOCIAL', 'RAZÓN SOCIAL', 'NOMBRE', 'DENOMINACION', 'DENOMINACIÓN EMISOR']:
                        if v in df_dicc.columns: df_dicc.rename(columns={v: 'PROVEEDOR_KEY'}, inplace=True); break
                    
                    falta_compras = 'PROVEEDOR_KEY' not in df.columns
                    falta_dicc = 'PROVEEDOR_KEY' not in df_dicc.columns
                    
                    if falta_compras and falta_dicc:
                        st.error("❌ No detectamos la columna de Proveedor en NINGUNO de los dos archivos. Asegurate de que la fila 1 tenga los títulos correctos.")
                        st.stop()
                    elif falta_compras:
                        st.error("❌ Error en el Paso 1: No encontramos la columna 'Proveedor' (o 'Razón Social') en tu archivo de COMPRAS.")
                        st.stop()
                    elif falta_dicc:
                        st.error("❌ Error en el Paso 2: No encontramos la columna 'Proveedor' (o 'Razón Social') en tu archivo de DICCIONARIO.")
                        st.stop()
                    
                    df = df.dropna(subset=['PROVEEDOR_KEY'])
                    df = df[df['PROVEEDOR_KEY'].astype(str).str.strip() != '']
                    df = df[df['PROVEEDOR_KEY'].astype(str).str.lower() != 'nan']
                    df = df[~df['PROVEEDOR_KEY'].astype(str).str.upper().str.contains('TOTAL', na=False)]
                    
                    for col in df.columns:
                        if 'FECHA' in col.upper():
                            df = df[~df[col].astype(str).str.upper().str.contains('TOTAL', na=False)]
                            break

                    df['P_MATCH'] = df['PROVEEDOR_KEY'].astype(str).str.upper().str.replace(r'[.,]', '', regex=True).str.replace(r'\s+', ' ', regex=True).str.strip()
                    df_dicc['P_MATCH'] = df_dicc['PROVEEDOR_KEY'].astype(str).str.upper().str.replace(r'[.,]', '', regex=True).str.replace(r'\s+', ' ', regex=True).str.strip()
                    
                    df_dicc['CUENTA'] = df_dicc['CUENTA'].astype(str).replace(['nan', 'NaN', 'None', ''], '⚡ Clasificar Proveedor')
                    
                    mapeo = dict(zip(df_dicc['P_MATCH'], df_dicc['CUENTA']))
                    df['Cuenta_Gasto'] = df['P_MATCH'].map(mapeo).fillna('⚡ Clasificar Proveedor')

                    faltantes_reales = df[df['Cuenta_Gasto'] == '⚡ Clasificar Proveedor']['PROVEEDOR_KEY'].unique().tolist()
                    st.session_state.proveedores_faltantes = [str(p) for p in faltantes_reales if str(p).strip() != '']

                    diccionario_columnas = {
                        'Exento': ['EXENTO', 'OPERACIONES EXENTAS', 'NO GRAVADO', 'CONCEPTOS NO GRAVADOS', 'IMP. OP. EXENTAS'], 
                        'Gravado': ['GRAVADO', 'NETO GRAVADO', 'IMPORTE NETO GRAVADO', 'IMP. NETO GRAVADO'], 
                        'IVA 10,50': ['IVA 10,50', 'IVA 10.50', 'IVA 10,50%', 'IVA 10.50%', 'IVA 10,5%', 'IVA 10.5%', 'IVA AL 10,50%'], 
                        'IVA 21%': ['IVA', 'IMPORTE IVA', 'IVA 21%', 'IVA 21', 'IVA AL 21%', 'IVA 21,00%', 'IVA 21.00%'], 
                        'IVA 27%': ['IVA 27%', 'IVA 27', 'IVA AL 27%'], 
                        'Retenciones IIBB': ['RETENCIONES IIBB', 'PERCEPCIONES IIBB', 'PERC. IIBB', 'PERCEPCIÓN IIBB', 'OTROS TRIBUTOS'], 
                        'Percepciones IVA': ['PERCEPCIONES IVA', 'PERC. IVA', 'PERCEPCIÓN IVA', 'PERCEPCION IVA', 'PERC IVA', 'PERCEPCIONES NACIONALES'],
                        'Importe Total': ['IMPORTE TOTAL', 'TOTAL', 'TOTAL FACTURADO']
                    }

                    for col_estandar, variantes in diccionario_columnas.items():
                        col_encontrada = False
                        for variante in variantes:
                            if variante in df.columns:
                                if df[variante].dtype == 'object':
                                    serie = df[variante].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                    df[col_estandar] = pd.to_numeric(serie, errors='coerce').fillna(0)
                                else:
                                    df[col_estandar] = pd.to_numeric(df[variante], errors='coerce').fillna(0)
                                col_encontrada = True; break
                        if not col_encontrada: df[col_estandar] = 0.0

                    df['Neto_Total'] = df['Exento'] + df['Gravado']
                    f_max = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce').max() if 'FECHA' in df.columns else None
                    f_str = f_max.strftime('%d/%m/%Y') if not pd.isnull(f_max) else "01/01/2026"

                    asiento = []
                    gastos = df.groupby('Cuenta_Gasto')['Neto_Total'].sum().reset_index()
                    for _, r in gastos.iterrows():
                        if r['Neto_Total'] != 0: asiento.append({"Fecha": f_str, "Cuenta": r['Cuenta_Gasto'], "Debe": r['Neto_Total'], "Haber": 0.0})
                    
                    lista_impuestos = [
                        ("IVA CF 10.5%", 'IVA 10,50'), 
                        ("IVA CF 21%", 'IVA 21%'), 
                        ("IVA CF 27%", 'IVA 27%'), 
                        ("Ret. IIBB", 'Retenciones IIBB'),
                        ("Percepciones IVA", 'Percepciones IVA')
                    ]
                    
                    for nom_cta, col_val in lista_impuestos:
                        if df[col_val].sum() != 0: asiento.append({"Fecha": f_str, "Cuenta": nom_cta, "Debe": df[col_val].sum(), "Haber": 0.0})
                    
                    asiento.append({"Fecha": f_str, "Cuenta": "Proveedores", "Debe": 0.0, "Haber": df['Importe Total'].sum()})

                    df_final_c = pd.DataFrame(asiento)
                    diferencia = round(round(df_final_c['Haber'].sum(), 2) - round(df_final_c['Debe'].sum(), 2), 2)
                    
                    if diferencia != 0:
                        asiento.append({"Fecha": f_str, "Cuenta": "⚡ Ajuste/Diferencia", "Debe": diferencia if diferencia > 0 else 0.0, "Haber": abs(diferencia) if diferencia < 0 else 0.0})
                        df_final_c = pd.DataFrame(asiento)

                    st.session_state.asiento_generado = df_final_c
                    st.session_state.fecha_asiento = f_str
                    st.session_state.contador_usos += 1
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error en el proceso: {e}")

    if st.session_state.asiento_generado is not None:
        st.markdown("---")
        if st.session_state.proveedores_faltantes:
            st.error(f"⚠️ **¡Atención! Hay {len(st.session_state.proveedores_faltantes)} proveedor(es) sin clasificar.**")
            st.info("💡 Copiá los nombres de esta lista y pegalos en la columna 'Proveedor' de tu Excel de Diccionario para la próxima vez:")
            df_faltantes = pd.DataFrame({"Proveedor (Copiar)": st.session_state.proveedores_faltantes, "Cuenta Contable (A completar)": ""})
            st.dataframe(df_faltantes, use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("""<div style="background-color: rgba(43, 108, 176, 0.1); border-left: 5px solid #2B6CB0; padding: 20px; border-radius: 10px; margin-bottom: 25px;"><h3 style="color: #2B6CB0; margin: 0;">✅ ¡Asiento generado!</h3></div>""", unsafe_allow_html=True)
        st.table(st.session_state.asiento_generado.style.format({"Debe": "${:,.2f}", "Haber": "${:,.2f}"}))
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            try:
                buf_c = io.BytesIO()
                with pd.ExcelWriter(buf_c, engine='openpyxl') as writer: st.session_state.asiento_generado.to_excel(writer, index=False)
                st.download_button("📥 DESCARGAR ASIENTO (EXCEL)", buf_c.getvalue(), f"asiento_compras_{st.session_state.fecha_asiento.replace('/','_')}.xlsx", key="down_compras")
            except Exception:
                csv_data = st.session_state.asiento_generado.to_csv(index=False, sep=';', decimal=',')
                st.download_button("📥 DESCARGAR ASIENTO (CSV)", csv_data.encode('utf-8-sig'), f"asiento_compras.csv", mime="text/csv", key="down_compras_csv")

st.markdown("<br><br><p style='text-align: center; opacity: 0.5;'>Hecho con amor ❤️ para automatizar contabilidad</p>", unsafe_allow_html=True)
