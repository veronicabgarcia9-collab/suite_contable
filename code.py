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

# --- ESTILOS MODO OSCURO EJECUTIVO ---
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; color: #E0E0E0; }
    
    /* Título General */
    h1 { color: #4DA8DA; font-family: 'Segoe UI', sans-serif; font-weight: 800; text-align: center; text-transform: uppercase; letter-spacing: 1px; padding-bottom: 20px; }
    
    /* Subtítulos y Labels */
    h2, h3, h4, label, .stMarkdown p { color: #4DA8DA !important; }
    
    /* Cajas de subida de archivos */
    [data-testid="stFileUploadDropzone"] {
        background-color: #1E1E1E !important;
        border: 2px dashed #4DA8DA !important;
        border-radius: 12px !important;
    }
    
    /* Contenedor de contador */
    .counter-container { 
        background-color: #1E1E1E; 
        border: 1px solid #333; 
        border-radius: 10px; 
        padding: 15px; 
        text-align: center; 
        margin-bottom: 25px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3); 
        color: #E0E0E0; 
    }
    .counter-number { color: #00FFCC; font-weight: 800; font-size: 24px; }

    /* Botones Principales */
    div.stButton > button:first-child { 
        background-color: #4DA8DA; 
        color: #0E1117; 
        border: none; 
        border-radius: 8px; 
        font-weight: bold; 
        width: 100%; 
        transition: 0.3s; 
        height: 3rem;
    }
    div.stButton > button:first-child:hover { background-color: #388BB8; transform: translateY(-2px); }
    
    /* Botón de Descarga */
    .stDownloadButton > button { 
        background-color: #00FFCC !important; 
        color: #0E1117 !important; 
        border: none !important; 
        border-radius: 8px !important; 
        width: 100%; 
        font-weight: bold; 
        transition: 0.3s;
    }
    .stDownloadButton > button:hover { background-color: #00CCA3 !important; opacity: 0.9; }
    
    /* Card de Donación */
    .donation-card { 
        background-color: #1E1E1E; 
        border: 1px solid #333; 
        border-radius: 12px; 
        padding: 25px; 
        text-align: center; 
        box-shadow: 0 10px 20px rgba(0,0,0,0.4); 
        color: #E0E0E0; 
    }
    
    /* Pestañas (TAMAÑO 22px y Colores Dark) */
    .stTabs [data-baseweb="tab-list"] { gap: 24px; background-color: transparent; }
    .stTabs [data-baseweb="tab"] { 
        height: 60px; 
        white-space: pre-wrap; 
        background-color: #1E1E1E; 
        border-radius: 8px 8px 0px 0px; 
        padding: 10px 20px; 
        font-weight: bold; 
        font-size: 22px !important; 
        color: #888; 
    }
    .stTabs [aria-selected="true"] { 
        color: #4DA8DA !important; 
        border-bottom: 4px solid #4DA8DA !important; 
        background-color: #262730;
    }

    /* Dataframes y tablas */
    .stDataFrame, .stTable { background-color: #1E1E1E; border-radius: 10px; }
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
    texto_completo = "\n".join([p.extract_text() or "" for p in pdf.pages])
    lineas = texto_completo.split('\n')
    patron_f = re.compile(r'^\s*(\d{2}/\d{2}/\d{2,4})')
    patron_m = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d{2}')
    movs, monto_ini = [], 0.0
    match_ini = re.search(r'saldo[\s\n]*inicial[\s\n]*\$?[\s\n]*(-?\d{1,3}(?:\.\d{3})*,\d{2})', texto_completo.lower())
    if match_ini: monto_ini = limpiar_monto_ar(match_ini.group(1))
    else:
        p1 = pdf.pages[0].extract_text() or ""
        tm = patron_m.findall(p1)
        if tm: monto_ini = limpiar_monto_ar(tm[0])
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
                if abs(dif - val) > 0.01 and abs(dif + val) < 0.01: val = -val 
                conc = re.sub(r'\s+', ' ', linea.replace(f, "")).strip()
                for m in ms: conc = conc.replace(m, "")
                movs.append({'Fecha': f, 'Concepto': conc, 'Debitos': abs(val) if val < 0 else 0.0, 'Creditos': val if val > 0 else 0.0, 'Neto': val, 'Saldo': sa})
                last_s = sa
    df = pd.DataFrame(movs)
    if not df.empty:
        fi = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': monto_ini}])
        df = pd.concat([fi, df], ignore_index=True)
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
        montos = patron_monto.findall(linea)
        if match_f and match_f.start() < 15:
            if mov_actual: movs_crudos.append(mov_actual)
            mov_actual = {'Fecha': match_f.group(1), 'Concepto': re.sub(r'\s+', ' ', linea.replace(match_f.group(1), "")).strip().replace('"', ''), 'Montos_Raw': montos}
        elif mov_actual:
            if montos: mov_actual['Montos_Raw'].extend(montos)
            mov_actual['Concepto'] += " " + re.sub(r'\s+', ' ', linea).strip()
    if mov_actual: movs_crudos.append(mov_actual)
    movs_fin = []
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
            if val_saldo is not None and i > 0 and movs_fin[-1]['Saldo'] is not None:
                delta = val_saldo - movs_fin[-1]['Saldo']
                neto = val_mov if abs(delta - val_mov) < 0.10 else -val_mov
            else: neto = val_mov if any(x in mov['Concepto'].upper() for x in ["DEP", "CREDITO", "RESCATE", "VENTA", "TICKET"]) else -val_mov
        movs_fin.append({'Fecha': mov['Fecha'], 'Concepto': mov['Concepto'], 'Debitos': abs(neto) if neto < 0 else 0.0, 'Creditos': neto if neto > 0 else 0.0, 'Neto': neto, 'Saldo': val_saldo})
    df = pd.DataFrame(movs_fin)
    if not df.empty:
        s0 = df.iloc[0]['Saldo'] if df.iloc[0]['Saldo'] is not None else 0.0
        ini = s0 - df.iloc[0]['Neto']
        df['Saldo'] = df['Saldo'].fillna(ini + df['Neto'].cumsum())
        fi = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': ini}])
        df = pd.concat([fi, df], ignore_index=True)
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
                if cta_actual not in cuentas: cuentas[cta_actual] = {"movs": [], "ini": 0.0, "last_saldo": 0.0}
                continue
            if cta_actual not in cuentas: cuentas[cta_actual] = {"movs": [], "ini": 0.0, "last_saldo": 0.0}
            if "SALDO ULTIMO EXTRACTO" in linea_u or "SALDO ANTERIOR" in linea_u:
                ms = patron_m.findall(linea)
                if ms:
                    val = limpiar_monto_ar(ms[-1])
                    cuentas[cta_actual]["ini"] = val
                    cuentas[cta_actual]["last_saldo"] = val
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
            fi = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': data["ini"]}])
            final_dfs[cta] = pd.concat([fi, df], ignore_index=True)
    return final_dfs

def motor_mercado_pago(pdf):
    texto_layout = "\n".join([p.extract_text(layout=True) or "" for p in pdf.pages])
    m_ini = re.search(r"Saldo inicial[\s:]*\$?\s*(-?[\d\.]+[,.]\d{2})", texto_layout, re.IGNORECASE)
    s_ini = limpiar_monto_ar(m_ini.group(1)) if m_ini else 0.0
    movs, mov_actual = [], None
    for linea in texto_layout.split('\n'):
        if not linea.strip() or any(x in linea.upper() for x in ["DETALLE", "SALDO INICIAL", "ENTRADAS", "SALIDAS"]): continue
        m_f = re.search(r'^(\d{2}-\d{2}-\d{4})', linea.strip())
        if m_f:
            if mov_actual: movs.append(mov_actual)
            fecha = m_f.group(1)
            resto = linea.replace(fecha, "", 1)
            montos = re.findall(r'-?\s?\$?\s*[\d\.]+[,.]\d{2}', resto)
            ids = re.findall(r'\b\d{10,}\b', resto)
            concepto = resto
            for m in montos: concepto = concepto.replace(m, "")
            for i in ids: concepto = concepto.replace(i, "")
            val_str = montos[-2] if len(montos) >= 2 else (montos[0] if montos else "0")
            mov_actual = {"Fecha": fecha.replace('-', '/'), "Concepto": concepto.strip(), "Valor_str": val_str}
        elif mov_actual:
            l_clean = re.sub(r'\b\d{10,}\b', '', linea).strip()
            l_clean = re.sub(r'-?\s?\$?\s*[\d\.]+[,.]\d{2}', '', l_clean).strip()
            if l_clean and not any(x in l_clean for x in ["Mercado Libre", "www."]) and not re.match(r'^\d+/\d+', l_clean):
                mov_actual["Concepto"] += " " + l_clean
    if mov_actual: movs.append(mov_actual)
    movs_fin = []
    for m in movs:
        val = limpiar_monto_ar(m["Valor_str"])
        conc = re.sub(r'\s+', ' ', m["Concepto"]).strip()
        movs_fin.append({"Fecha": m["Fecha"], "Concepto": conc, "Debitos": abs(val) if val < 0 else 0.0, "Creditos": val if val > 0 else 0.0, "Neto": val})
    df = pd.DataFrame(movs_fin)
    if not df.empty:
        df['Saldo'] = s_ini + df['Neto'].cumsum()
        fi = pd.DataFrame([{"Fecha": df.iloc[0]["Fecha"], "Concepto": "SALDO INICIAL", "Debitos": 0.0, "Creditos": 0.0, "Neto": 0.0, "Saldo": s_ini}])
        return {"Mercado Pago": pd.concat([fi, df], ignore_index=True)}
    return {}

def motor_icbc(pdf):
    texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
    m_ini = re.search(r"SALDO\s+ULTIMO\s+EXTRACTO.*?([\d\.]+,\d{2}-?)", texto.upper().replace('\n', ' '))
    def lim(t):
        n = '-' in str(t); v = str(t).replace('-', '').replace('.', '').replace(',', '.')
        return -float(v) if n else float(v)
    s_ini = lim(m_ini.group(1)) if m_ini else 0.0
    movs = []
    for line in texto.split('\n'):
        m_f = re.match(r'^(\d{2}-\d{2})\s+(.+)', line.strip())
        if m_f:
            f, resto = m_f.groups()
            ms = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}-?', resto)
            if ms:
                val = lim(ms[0]); conc = re.sub(r'\b\d{4}\b', '', resto.replace(ms[0], '').strip()).strip()
                movs.append({'Fecha': f.replace('-', '/'), 'Concepto': conc, 'Debitos': abs(val) if val < 0 else 0.0, 'Creditos': val if val > 0 else 0.0, 'Neto': val})
    df = pd.DataFrame(movs)
    if not df.empty:
        df['Saldo'] = s_ini + df['Neto'].cumsum()
        fi = pd.DataFrame([{'Fecha': df.iloc[0]['Fecha'], 'Concepto': 'SALDO INICIAL', 'Debitos': 0.0, 'Creditos': 0.0, 'Neto': 0.0, 'Saldo': s_ini}])
        return {"Resumen ICBC": pd.concat([fi, df], ignore_index=True)}
    return {}

def motor_credicoop(pdf):
    movs, s_ini, ini_set = [], 0.0, False
    for page in pdf.pages:
        words = page.extract_words()
        lineas = {}
        for w in words: lineas.setdefault(round(w['top'], 1), []).append(w)
        for y in sorted(lineas.keys()):
            f_w = sorted(lineas[y], key=lambda x: x['x0'])
            txt = " ".join([w['text'] for w in f_w])
            if "SALDO ANTERIOR" in txt.upper() and not ini_set:
                ms = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', txt)
                if ms: s_ini = limpiar_monto_ar(ms[-1]); ini_set = True
            m_f = re.match(r'^(\d{2}/\d{2}/\d{2})', txt)
            if m_f:
                fecha, deb, cre, cp = m_f.group(1), 0.0, 0.0, []
                for p in f_w:
                    if re.search(r'\d,\d{2}', p['text']):
                        v = limpiar_monto_ar(p['text'])
                        if p['x0'] < 445: deb = abs(v)
                        else: cre = v
                    elif p['text'] not in fecha: cp.append(p['text'])
                if deb > 0 or cre > 0: movs.append({"Fecha": fecha, "Concepto": " ".join(cp), "Debitos": deb, "Creditos": cre, "Neto": cre-deb})
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
            f_s = m_f.group(1); r_l = m_f.group(2)
            if f_s == "00/00" or "SIN MOVIMIENTOS" in r_l.upper(): continue
            ms = re.findall(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})', line)
            if ms:
                sa = limpiar_monto_ar(ms[-1]); dif = round(sa - last_s, 2)
                if abs(dif) > 0.01:
                    movs.append({"Fecha": f_s, "Concepto": line.replace(f_s, "").replace(ms[-1], "").strip(), "Debitos": abs(dif) if dif < 0 else 0.0, "Creditos": dif if dif > 0 else 0.0, "Neto": dif, "Saldo": sa})
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
        if st.button("🔄 Reiniciar App"):
            st.session_state.contador_usos = 0; limpiar_memoria(); st.rerun()
    st.stop()

# --- PESTAÑAS (TAMAÑO 22px) ---
tab1, tab2 = st.tabs(["🏦 Bancos a Excel", "🛒 Compras a Asientos"])

with tab1:
    st.markdown("### 📄 Procesador de Extractos Bancarios")
    banco_sel = st.selectbox("Seleccionar Banco:", ["Macro", "Galicia", "Mercado Pago", "ICBC", "Credicoop", "BBVA Francés"])
    c1, c2 = st.columns(2)
    with c1: archivo_pdf = st.file_uploader("Subir PDF del banco", type="pdf")
    with c2: archivo_dic_bancos = st.file_uploader("Subir Diccionario (Opcional)", type="xlsx", key="db")
    if archivo_pdf:
        try:
            with pdfplumber.open(archivo_pdf) as pdf:
                if banco_sel == "Galicia":
                    t1 = (pdf.pages[0].extract_text() or "").upper()
                    d_dfs = motor_galicia_office(pdf) if "OFFICE" in t1 or '","' in t1 else motor_galicia_tradicional(pdf)
                elif banco_sel == "Macro": d_dfs = motor_macro(pdf)
                elif banco_sel == "ICBC": d_dfs = motor_icbc(pdf)
                elif banco_sel == "Mercado Pago": d_dfs = motor_mercado_pago(pdf)
                elif banco_sel == "Credicoop": d_dfs = motor_credicoop(pdf)
                else: d_dfs = motor_frances(pdf)
                if d_dfs:
                    st.session_state.contador_usos += 1
                    df_man_b = pd.read_excel(archivo_dic_bancos) if archivo_dic_bancos else None
                    c_sel = st.selectbox("Cuenta detectada:", list(d_dfs.keys()))
                    df_fin_b = aplicar_diccionario_final(d_dfs[c_sel], df_man_b)
                    cols_b = ['Fecha', 'Concepto', 'Debitos', 'Creditos', 'Saldo', 'Imputación']
                    df_m_b = df_fin_b[[c for c in cols_b if c in df_fin_b.columns]]
                    st.dataframe(df_m_b.style.format({'Debitos': '{:,.2f}', 'Creditos': '{:,.2f}', 'Saldo': '{:,.2f}'}), use_container_width=True)
                    out_b = io.BytesIO()
                    with pd.ExcelWriter(out_b) as writer:
                        for k, df_b in d_dfs.items():
                            df_ex_b = aplicar_diccionario_final(df_b, df_man_b)
                            df_ex_b[[c for c in cols_b if c in df_ex_b.columns]].to_excel(writer, index=False, sheet_name=str(k)[:31].replace("/", "-"))
                    st.download_button("🚀 DESCARGAR EXCEL", out_b.getvalue(), f"conciliacion_{banco_sel.lower()}.xlsx")
        except Exception as e: st.error(f"Error: {e}")

with tab2:
    st.markdown("### 🛒 Generador de Asientos de Compras")
    cc1, cc2 = st.columns(2)
    with cc1: arc_comp = st.file_uploader("1. Subir compras", type=["xlsx", "xls"], key="c", on_change=limpiar_memoria)
    with cc2: arc_dic_c = st.file_uploader("2. Subir diccionario", type=["xlsx"], key="dc", on_change=limpiar_memoria)
    if arc_comp and arc_dic_c:
        if st.button("🔥 GENERAR ASIENTO"):
            try:
                df = pd.read_excel(arc_comp); df_d = pd.read_excel(arc_dic_c)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                df_d.columns = df_d.columns.astype(str).str.strip().str.upper()
                for v in ['PROVEEDOR', 'RAZON SOCIAL', 'NOMBRE', 'DENOMINACION']:
                    if v in df.columns: df.rename(columns={v: 'PROVEEDOR_KEY'}, inplace=True); break
                for v in ['PROVEEDOR', 'RAZON SOCIAL', 'NOMBRE', 'DENOMINACION']:
                    if v in df_d.columns: df_d.rename(columns={v: 'PROVEEDOR_KEY'}, inplace=True); break
                df = df.dropna(subset=['PROVEEDOR_KEY'])
                df['P_MATCH'] = df['PROVEEDOR_KEY'].astype(str).str.upper().str.replace(r'[.,]', '', regex=True).str.replace(r'\s+', ' ', regex=True).str.strip()
                df_d['P_MATCH'] = df_d['PROVEEDOR_KEY'].astype(str).str.upper().str.replace(r'[.,]', '', regex=True).str.replace(r'\s+', ' ', regex=True).str.strip()
                mapeo = dict(zip(df_d['P_MATCH'], df_d['CUENTA']))
                df['Cuenta_Gasto'] = df['P_MATCH'].map(mapeo).fillna('⚡ Clasificar')
                st.session_state.proveedores_faltantes = df[df['Cuenta_Gasto'] == '⚡ Clasificar']['PROVEEDOR_KEY'].unique().tolist()
                dict_cols = {'Exento': ['EXENTO', 'NO GRAVADO'], 'Gravado': ['GRAVADO', 'NETO GRAVADO'], 'IVA 10,5': ['IVA 10,5', 'IVA 10.5'], 'IVA 21': ['IVA 21', 'IVA'], 'IVA 27': ['IVA 27'], 'IIBB': ['RETENCIONES IIBB', 'OTROS TRIBUTOS'], 'Perc_IVA': ['PERCEPCIONES IVA'], 'Total': ['IMPORTE TOTAL', 'TOTAL']}
                for c_est, vars in dict_cols.items():
                    fnd = False
                    for v in vars:
                        if v in df.columns:
                            s = df[v].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                            df[c_est] = pd.to_numeric(s, errors='coerce').fillna(0); fnd = True; break
                    if not fnd: df[c_est] = 0.0
                df['Neto_T'] = df['Exento'] + df['Gravado']
                f_s = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce').max().strftime('%d/%m/%Y')
                asnt = []
                gst = df.groupby('Cuenta_Gasto')['Neto_T'].sum().reset_index()
                for _, r in gst.iterrows(): 
                    if r['Neto_T'] != 0: asnt.append({"Fecha": f_s, "Cuenta": r['Cuenta_Gasto'], "Debe": r['Neto_T'], "Haber": 0.0})
                for n, c in [("IVA 10.5", 'IVA 10,5'), ("IVA 21", 'IVA 21'), ("IVA 27", 'IVA 27'), ("IIBB", 'IIBB'), ("Perc IVA", 'Perc_IVA')]:
                    if df[c].sum() != 0: asnt.append({"Fecha": f_s, "Cuenta": n, "Debe": df[c].sum(), "Haber": 0.0})
                asnt.append({"Fecha": f_s, "Cuenta": "Proveedores", "Debe": 0.0, "Haber": df['Total'].sum()})
                df_f_c = pd.DataFrame(asnt)
                dif = round(df_f_c['Haber'].sum() - df_f_c['Debe'].sum(), 2)
                if dif != 0: asnt.append({"Fecha": f_s, "Cuenta": "⚡ Ajuste", "Debe": dif if dif > 0 else 0.0, "Haber": abs(dif) if dif < 0 else 0.0})
                st.session_state.asiento_generado = pd.DataFrame(asnt); st.session_state.fecha_asiento = f_s; st.session_state.contador_usos += 1; st.rerun()
            except Exception as e: st.error(f"Error: {e}")

    if st.session_state.asiento_generado is not None:
        if st.session_state.proveedores_faltantes:
            st.error("⚠️ Proveedores sin clasificar detectados.")
            st.dataframe(pd.DataFrame({"Proveedor": st.session_state.proveedores_faltantes}), use_container_width=True)
        st.table(st.session_state.asiento_generado.style.format({"Debe": "${:,.2f}", "Haber": "${:,.2f}"}))
        buf_c = io.BytesIO()
        with pd.ExcelWriter(buf_c) as writer: st.session_state.asiento_generado.to_excel(writer, index=False)
        st.download_button("📥 DESCARGAR EXCEL", buf_c.getvalue(), f"asiento_{st.session_state.fecha_asiento.replace('/','_')}.xlsx")

st.markdown("<br><br><p style='text-align: center; opacity: 0.5;'>Hecho con amor ❤️ para contadores</p>", unsafe_allow_html=True)
