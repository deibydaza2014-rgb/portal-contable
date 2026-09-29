import pickle
import shutil
import json
import base64
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import io
import zipfile
import re
import datetime
import os
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Sistema ERP y Auditoria Contable DIAN", layout="wide", page_icon="🏢")

# Estilos visuales profesionales
st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    .stButton>button { background-color: #0070ba; color: white; border-radius: 6px; font-weight: 600; }
    .card-box { background: white; padding: 22px; border-radius: 10px; border: 1px solid #e2e8f0; box-shadow: 0 2px 5px rgba(0,0,0,0.04); margin-bottom: 18px; }
    .audit-card { background: #ffffff; padding: 22px; border-left: 5px solid #0070ba; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
    .badge-active { background-color: #dcfce7; color: #15803d; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .badge-next { background-color: #f1f5f9; color: #64748b; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .tag-propio { background-color: #dcfce7; color: #15803d; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .tag-comp { background-color: #eff6ff; color: #1d4ed8; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-family: monospace; }
    .badge-tax { background-color: #e0f2fe; color: #0369a1; padding: 3px 8px; border-radius: 6px; font-size: 12px; font-weight: 600; margin-right: 4px; }
</style>
""", unsafe_allow_html=True)

# 1. ESTADOS DE SESION
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "empresa_activa" not in st.session_state:
    st.session_state["empresa_activa"] = None
if "proceso_activo" not in st.session_state:
    st.session_state["proceso_activo"] = None

# PANTALLA 1: LOGIN
if not st.session_state["autenticado"]:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.markdown("<h2 style='text-align: center; color: #0f172a;'>Portal ERP y Auditoria Contable</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #64748b;'>Plataforma unificada para gestion contable y DIAN (Siigo / World Office)</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario o Correo Electronico", value="deibydaza2014@gmail.com")
            password = st.text_input("Contrasena", type="password", value="123456")
            submit = st.form_submit_button("Ingresar al Ecosistema")
            
            if submit:
                if usuario and password:
                    st.session_state["autenticado"] = True
                    st.rerun()
                else:
                    st.error("Ingresa usuario y contrasena.")
    st.stop()

# PANTALLA 2: SELECTOR DE EMPRESAS CON RESPONSABILIDADES RUT
EMPRESAS_DISPONIBLES = [
    {
        "nombre": "INDUMAQ ER SAS",
        "nit": "901.346.412-5",
        "actividad": "Comercio y Reparacion de Maquinaria / Importaciones",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "ACTIVA",
        "badge": "badge-active"
    },
    {
        "nombre": "ASMINCOL S.A.S.",
        "nit": "900.467.519-1",
        "actividad": "Servicios Mineros y Construccion",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "CONSTRUDISENO CT SAS",
        "nit": "900.524.356-8",
        "actividad": "Construccion y Obras Civiles",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "SCG TRANSPORTES",
        "nit": "901.700.731-8",
        "actividad": "Transporte de Carga y Logistica",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "OPJ SAS",
        "nit": "901.425.101-3",
        "actividad": "Servicios Generales y Operaciones",
        "regimen": "Responsable de IVA",
        "es_gran_contribuyente": False,
        "es_autorretenedor": False,
        "responsabilidades_rut": ["O-48"],
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    }
]

if not st.session_state["empresa_activa"]:
    st.title("Seleccion de Empresa")
    st.caption("Selecciona la entidad sobre la cual vas a trabajar:")
    
    c1, c2 = st.columns(2)
    for i, emp in enumerate(EMPRESAS_DISPONIBLES):
        col = c1 if i % 2 == 0 else c2
        with col:
            st.markdown(f"""
            <div class="card-box">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h3 style="margin:0; color:#0f172a;">🏢 {emp['nombre']}</h3>
                    <span class="{emp['badge']}">{emp['estado']}</span>
                </div>
                <p style="margin:8px 0 2px 0; color:#475569;"><b>NIT:</b> {emp['nit']} - <b>Regimen:</b> {emp['regimen']}</p>
                <p style="margin:2px 0 14px 0; color:#64748b; font-size:14px;">{emp['actividad']}</p>
            </div>
            """, unsafe_allow_html=True)
            if emp["estado"] == "ACTIVA":
                if st.button(f"Ingresar a {emp['nombre']}", key=f"btn_emp_{i}"):
                    st.session_state["empresa_activa"] = emp
                    st.session_state["proceso_activo"] = None
                    st.rerun()
            else:
                st.button(f"Pendiente Datos / Parametrizacion ({emp['estado']})", key=f"btn_emp_{i}", disabled=True)
    st.stop()

empresa = st.session_state["empresa_activa"]

# PANTALLA 3: MENU DE PROCESOS OPERATIVOS
PROCESOS_SISTEMA = [
    {
        "id": "facturacion",
        "icono": "Facturas",
        "titulo": "Facturas de Compra, Venta y Devoluciones",
        "desc": "Carga de reportes DIAN/Token, desbloqueo y renombrado de facturas por comprobante, auditoria contable y plantilla Siigo.",
        "estado": "ACTIVO",
        "badge": "badge-active"
    },
    {
        "id": "nomina",
        "icono": "Nomina",
        "titulo": "Gestion Laboral y Nomina Electronica",
        "desc": "Calculo de liquidacion de nomina, provisiones de prestaciones sociales, seguridad social y emision de soportes electronicos DIAN.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "id": "conciliacion",
        "icono": "Bancos",
        "titulo": "Tesoreria y Conciliacion Bancaria",
        "desc": "Cruce automatizado de extractos bancarios (Bancolombia, Davivienda) contra libros auxiliares y control de partidas conciliatorias.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "id": "notas",
        "icono": "Ajustes",
        "titulo": "Notas de Contabilidad y Cierre Fiscal",
        "desc": "Comprobantes de ajuste, amortizaciones de intangibles (NIC 38), depreciaciones y conciliacion fiscal NIIF vs. DIAN.",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    }
]

if not st.session_state["proceso_activo"]:
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.subheader(f"{empresa['nombre']} - Panel de Procesos")
        st.caption(f"NIT: {empresa['nit']} - Selecciona el modulo de trabajo que deseas ejecutar:")
    with col_t2:
        if st.button("Cambiar de Empresa"):
            st.session_state["empresa_activa"] = None
            st.session_state["proceso_activo"] = None
            st.rerun()

    st.markdown("---")
    
    cp1, cp2 = st.columns(2)
    for idx, proc in enumerate(PROCESOS_SISTEMA):
        col_p = cp1 if idx % 2 == 0 else cp2
        with col_p:
            st.markdown(f"""
            <div class="card-box">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h3 style="margin:0; color:#0f172a;">{proc['titulo']}</h3>
                    <span class="{proc['badge']}">{proc['estado']}</span>
                </div>
                <p style="margin:12px 0 16px 0; color:#475569; font-size:14px; line-height:1.5;">{proc['desc']}</p>
            </div>
            """, unsafe_allow_html=True)
            if proc["estado"] == "ACTIVO":
                if st.button(f"Abrir Modulo de Facturacion", key=f"btn_proc_{idx}"):
                    st.session_state["proceso_activo"] = proc["id"]
                    st.rerun()
            else:
                st.button(f"Modulo en Construccion ({proc['estado']})", key=f"btn_proc_{idx}", disabled=True)
    st.stop()

# PANTALLA 4: FACTURACION, AUDITORIA Y SIIGO
col_nav1, col_nav2 = st.columns(2)
with col_nav1:
    st.subheader(f"{empresa['nombre']} - Facturacion, Auditoria y Siigo")
    st.caption(f"NIT: {empresa['nit']} - Modulo Activo: Facturas de Compra, Venta y Devoluciones")
with col_nav2:
    if st.button("Volver a Procesos"):
        st.session_state["proceso_activo"] = None
        st.rerun()

st.markdown("---")

# ==============================================================================
# GESTOR DE HISTORIAL DE TRABAJOS Y AUDITORÍAS PASADAS (PERSISTENCIA TOTAL)
# ==============================================================================
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "almacenamiento_contable")
os.makedirs(DATA_DIR, exist_ok=True)

def get_empresa_trabajos_dir(empresa_dict):
    nit_clean = re.sub(r"\D", "", str(empresa_dict.get("nit", "empresa")))
    d = os.path.join(DATA_DIR, nit_clean, "trabajos")
    os.makedirs(d, exist_ok=True)
    return d

def guardar_trabajo_en_historial(empresa_dict, df_procesado, excel_bytes=None, excel_nombre="Reporte.xlsx", dict_pdfs_renombrados=None, dict_pdfs_originales=None, zip_bytes=None, consecutivo_ini=680, job_id=None):
    try:
        base_dir = get_empresa_trabajos_dir(empresa_dict)
        if not job_id:
            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            job_id = f"job_{now_str}"
            
        job_dir = os.path.join(base_dir, job_id)
        os.makedirs(job_dir, exist_ok=True)
        
        dir_renom = os.path.join(job_dir, "pdfs_renombrados")
        os.makedirs(dir_renom, exist_ok=True)
        dir_orig = os.path.join(job_dir, "pdfs_originales")
        os.makedirs(dir_orig, exist_ok=True)
        
        # 1. Guardar DataFrame procesado
        if df_procesado is not None:
            with open(os.path.join(job_dir, "df_procesado.pkl"), "wb") as f:
                pickle.dump(df_procesado, f)
                
        # 2. Guardar archivo Excel original subido
        if excel_bytes:
            with open(os.path.join(job_dir, "excel_original.xlsx"), "wb") as f_ex:
                f_ex.write(excel_bytes)
                
        # 3. Guardar PDFs renombrados / separados
        if dict_pdfs_renombrados:
            for fname, bdata in dict_pdfs_renombrados.items():
                with open(os.path.join(dir_renom, fname), "wb") as pf:
                    pf.write(bdata)
                    
        # 4. Guardar PDFs originales subidos (unificados o separados)
        if dict_pdfs_originales:
            for fname, bdata in dict_pdfs_originales.items():
                with open(os.path.join(dir_orig, fname), "wb") as pf:
                    pf.write(bdata)
                    
        # 5. Guardar paquete ZIP de facturas si existe
        if zip_bytes:
            with open(os.path.join(job_dir, "paquete_facturas.zip"), "wb") as zf:
                zf.write(zip_bytes)
                
        now_disp = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
        # Verificar archivos físicos guardados en disco
        n_orig_disco = len(os.listdir(dir_orig)) if os.path.exists(dir_orig) else 0
        n_renom_disco = len(os.listdir(dir_renom)) if os.path.exists(dir_renom) else 0
        tiene_excel = os.path.exists(os.path.join(job_dir, "excel_original.xlsx")) or (excel_bytes is not None)
        
        meta = {
            "id": job_id,
            "nombre_trabajo": f"Trabajo del {now_disp} ({excel_nombre})",
            "archivo_excel": excel_nombre,
            "tiene_excel": tiene_excel,
            "fecha": now_disp,
            "total_facturas": len(df_procesado) if df_procesado is not None else 0,
            "consecutivo_inicial": consecutivo_ini,
            "total_valor": float(df_procesado["Total"].sum()) if (df_procesado is not None and "Total" in df_procesado) else 0.0,
            "total_pdfs_renombrados": n_renom_disco,
            "total_pdfs_originales": n_orig_disco
        }
        with open(os.path.join(job_dir, "meta.json"), "w", encoding="utf-8") as mf:
            json.dump(meta, mf, ensure_ascii=False, indent=2)
            
        # Actualizar índice general de la empresa para carga inmediata
        idx_file = os.path.join(base_dir, "historial_index.json")
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                index_data = json.load(f_idx)
        except Exception:
            index_data = []
        index_data = [x for x in index_data if x.get("id") != job_id]
        index_data.append(meta)
        index_data.sort(key=lambda x: x.get("id", ""), reverse=True)
        try:
            with open(idx_file, "w", encoding="utf-8") as f_idx:
                json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
        except Exception:
            pass
            
        return job_id
    except Exception:
        return None

def listar_trabajos_historial(empresa_dict):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    lista = []
    idx_file = os.path.join(base_dir, "historial_index.json")
    if os.path.exists(idx_file):
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                lista = json.load(f_idx)
        except Exception:
            lista = []
    if not lista and os.path.exists(base_dir):
        for jid in os.listdir(base_dir):
            jdir = os.path.join(base_dir, jid)
            meta_file = os.path.join(jdir, "meta.json")
            if os.path.isdir(jdir) and os.path.exists(meta_file):
                try:
                    with open(meta_file, "r", encoding="utf-8") as mf:
                        meta = json.load(mf)
                        lista.append(meta)
                except Exception:
                    pass
    lista.sort(key=lambda x: x.get("id", ""), reverse=True)
    return lista

def cargar_trabajo_historial(empresa_dict, job_id):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    jdir = os.path.join(base_dir, job_id)
    df = None
    dict_renom = {}
    dict_orig = {}
    zip_bytes = None
    excel_bytes = None
    excel_nombre = "Reporte.xlsx"
    
    meta_file = os.path.join(jdir, "meta.json")
    if os.path.exists(meta_file):
        try:
            with open(meta_file, "r", encoding="utf-8") as mf:
                m_data = json.load(mf)
                excel_nombre = m_data.get("archivo_excel", "Reporte.xlsx")
        except:
            pass

    pkl_file = os.path.join(jdir, "df_procesado.pkl")
    if os.path.exists(pkl_file):
        with open(pkl_file, "rb") as f:
            df = pickle.load(f)
            
    ex_file = os.path.join(jdir, "excel_original.xlsx")
    if os.path.exists(ex_file):
        with open(ex_file, "rb") as f_ex:
            excel_bytes = f_ex.read()
            
    # PDFs renombrados
    pdf_dir_r = os.path.join(jdir, "pdfs_renombrados")
    if not os.path.exists(pdf_dir_r):
        pdf_dir_r = os.path.join(jdir, "pdfs")
    if os.path.exists(pdf_dir_r):
        for fn in os.listdir(pdf_dir_r):
            if fn.endswith(".pdf"):
                with open(os.path.join(pdf_dir_r, fn), "rb") as pf:
                    dict_renom[fn] = pf.read()
                    
    # PDFs originales
    pdf_dir_o = os.path.join(jdir, "pdfs_originales")
    if os.path.exists(pdf_dir_o):
        for fn in os.listdir(pdf_dir_o):
            if fn.endswith(".pdf"):
                with open(os.path.join(pdf_dir_o, fn), "rb") as pf:
                    dict_orig[fn] = pf.read()
                    
    zip_file = os.path.join(jdir, "paquete_facturas.zip")
    if os.path.exists(zip_file):
        with open(zip_file, "rb") as zf:
            zip_bytes = zf.read()
    elif dict_renom or dict_orig:
        try:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_temp:
                fuente = dict_renom if dict_renom else dict_orig
                for fn, bdata in fuente.items():
                    zf_temp.writestr(fn, bdata)
            buf.seek(0)
            zip_bytes = buf.getvalue()
        except Exception:
            pass
            
    return df, dict_renom, dict_orig, excel_bytes, excel_nombre, zip_bytes

def eliminar_trabajo_historial(empresa_dict, job_id):
    base_dir = get_empresa_trabajos_dir(empresa_dict)
    jdir = os.path.join(base_dir, job_id)
    if os.path.exists(jdir):
        shutil.rmtree(jdir, ignore_errors=True)
    idx_file = os.path.join(base_dir, "historial_index.json")
    if os.path.exists(idx_file):
        try:
            with open(idx_file, "r", encoding="utf-8") as f_idx:
                index_data = json.load(f_idx)
            index_data = [x for x in index_data if x.get("id") != job_id]
            with open(idx_file, "w", encoding="utf-8") as f_idx:
                json.dump(index_data, f_idx, ensure_ascii=False, indent=2)
        except Exception:
            pass

# ==============================================================================
# MOTOR DE BÚSQUEDA Y EXTRACCIÓN INTELIGENTE DE PDF (UNIFICADO O SEPARADO)
# ==============================================================================
@st.cache_data(show_spinner=False)
def cache_extraer_textos_pdf(fbytes):
    """Extrae el texto de todas las páginas de un PDF y lo almacena en caché en RAM para búsqueda ultrarrápida."""
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        return [(p.extract_text() or "") for p in reader.pages]
    except Exception:
        return []

def buscar_y_extraer_pdf(fac_sel, dict_renombrados=None, dict_originales=None):
    """
    Busca y extrae los bytes del PDF de la factura seleccionada con coincidencia inteligente:
    1. Si ya se separó/renombró, busca coincidencia en dict_renombrados.
    2. Si se subieron PDFs individuales, busca por número de factura o NIT en el nombre.
    3. Si se subió un PDF unificado (o archivo consolidado como Todas_Licadas.pdf),
       escanea cada página identificando el folio (con o sin ceros a la izquierda),
       prefijo y el NIT del emisor (con o sin dígito de verificación) y extrae las páginas exactas.
    Retorna: (pdf_bytes, descripcion_origen, lista_paginas)
    """
    folio_clean = str(fac_sel.get("Folio", "")).replace("-", "").strip().upper()
    folio_sc = folio_clean.lstrip("0")
    pref_clean = str(fac_sel.get("Prefijo", "")).replace("-", "").strip().upper()
    fac_clean = str(fac_sel.get("Factura", "")).replace("-", "").strip().upper()
    
    nit_digits = re.sub(r"\D", "", str(fac_sel.get("NIT Emisor", "")))
    nit_base = nit_digits[:-1] if len(nit_digits) >= 10 else nit_digits
    soporte_nom = str(fac_sel.get("Soporte PDF Renombrado", ""))

    # 1. Búsqueda en renombrados
    if dict_renombrados:
        if soporte_nom in dict_renombrados:
            return dict_renombrados[soporte_nom], f"PDF Renombrado ({soporte_nom})", None
        for k, v in dict_renombrados.items():
            k_clean = k.replace("-", "").replace(" ", "").upper()
            if folio_clean and len(folio_clean) >= 3 and folio_clean in k_clean:
                return v, f"PDF Renombrado ({k})", None

    # 2. Búsqueda en originales por nombre de archivo
    if dict_originales:
        for fname, fbytes in dict_originales.items():
            fn_clean = fname.replace("-", "").replace(" ", "").upper()
            if (fac_clean and len(fac_clean) >= 4 and fac_clean in fn_clean) or (folio_clean and len(folio_clean) >= 3 and folio_clean in fn_clean):
                return fbytes, f"PDF Original Individual ({fname})", None

        # 3. Escaneo ultrarrápido página por página con caché en RAM (PDFs unificados)
        for fname, fbytes in dict_originales.items():
            try:
                paginas_txt = cache_extraer_textos_pdf(fbytes)
                pags_coincidentes = []
                for p_idx, txt in enumerate(paginas_txt):
                    txt_clean = txt.replace("-", "").replace(" ", "").upper()
                    
                    # Criterio de coincidencia robusta
                    match = False
                    
                    # A. Prefijo + Folio juntos
                    fac_full = (pref_clean + folio_clean) if pref_clean else folio_clean
                    fac_full_sc = (pref_clean + folio_sc) if pref_clean else folio_sc
                    if len(fac_full) >= 4 and fac_full in txt_clean:
                        match = True
                    elif len(fac_full_sc) >= 3 and fac_full_sc in txt_clean:
                        match = True
                    else:
                        # B. NIT del emisor + Número de Factura
                        tiene_nit = (nit_digits and len(nit_digits) >= 6 and nit_digits in txt_clean) or (nit_base and len(nit_base) >= 6 and nit_base in txt_clean)
                        if tiene_nit:
                            if len(folio_clean) >= 2 and folio_clean in txt_clean:
                                match = True
                            elif len(folio_sc) >= 2 and folio_sc in txt_clean:
                                match = True
                                
                    if match:
                        pags_coincidentes.append(p_idx)
                
                if pags_coincidentes:
                    reader = PdfReader(io.BytesIO(fbytes))
                    writer = PdfWriter()
                    for p in pags_coincidentes:
                        writer.add_page(reader.pages[p])
                    out = io.BytesIO()
                    writer.write(out)
                    out.seek(0)
                    str_p = ", ".join([str(p+1) for p in pags_coincidentes])
                    return out.getvalue(), f"Extraído de '{fname}' (Págs {str_p})", [p+1 for p in pags_coincidentes]
            except Exception:
                pass

    return None, None, None

# PANEL DE HISTORIAL DE TRABAJOS Y AUDITORÍAS PASADAS
st.markdown("### 🗂️ Historial de Trabajos y Auditorías Pasadas")
with st.expander("📂 Consultar y Cargar Trabajos Pasados de esta Empresa", expanded=False):
    trabajos_guardados = listar_trabajos_historial(empresa)
    if trabajos_guardados:
        st.caption("Selecciona cualquier trabajo realizado previamente para auditar facturas, ver PDFs o exportar:")
        for tb in trabajos_guardados:
            c_h1, c_h2, c_h3 = st.columns([3, 1, 1])
            with c_h1:
                n_renom = tb.get("total_pdfs_renombrados", tb.get("total_pdfs", 0))
                n_orig = tb.get("total_pdfs_originales", 0)
                info_pdf_str = f"{n_renom} PDFs procesados" if n_renom > 0 else (f"{n_orig} PDFs subidos" if n_orig > 0 else "Sin PDFs")
                st.markdown(f"📄 **{tb['nombre_trabajo']}** — {tb['total_facturas']} facturas (${tb['total_valor']:,.2f}) — **{info_pdf_str}**")
            with c_h2:
                if st.button("📂 Cargar este Trabajo", key=f"btn_h_load_{tb['id']}"):
                    df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa, tb["id"])
                    if df_g is not None:
                        st.session_state["df_procesado"] = df_g
                    if pdfs_r_g:
                        st.session_state["dict_pdfs"] = pdfs_r_g
                    if pdfs_o_g:
                        st.session_state["raw_uploaded_pdfs"] = pdfs_o_g
                    if ex_b_g:
                        st.session_state["excel_bytes"] = ex_b_g
                        st.session_state["excel_nombre"] = ex_n_g
                    if zip_g:
                        st.session_state["zip_pdfs"] = zip_g
                        st.session_state["total_zip_pdfs"] = len(pdfs_r_g) if pdfs_r_g else 0
                    st.session_state["job_actual_id"] = tb["id"]
                    st.success(f"¡Trabajo '{tb['nombre_trabajo']}' cargado! Puedes ir a Auditoría o Siigo.")
                    st.rerun()
            with c_h3:
                if st.button("🗑️", key=f"btn_h_del_{tb['id']}"):
                    eliminar_trabajo_historial(empresa, tb["id"])
                    st.rerun()
    else:
        st.info("💡 Aún no tienes trabajos guardados para esta empresa. Cada vez que subas un archivo Excel o proceses PDFs, se guardará aquí como un trabajo pasado para que nunca tengas que empezar desde cero.")

st.markdown("---")

tab_compras, tab_auditoria, tab_siigo = st.tabs([
    "1. Cargar Documentos, Desbloquear y Renombrar PDFs",
    "2. Auditoria y Trazabilidad Fiscal",
    "3. Exportar Planilla Oficial a Siigo"
])

CODIGOS_IMPUESTO_SIIGO = {
    '24081001': 1,   # IVA 19% compras bienes
    '24081003': 2,   # IVA 5% compras bienes
    '24080601': 1,   # IVA 19% ventas
    '24080602': 2,   # IVA 5% ventas
    '23651501': 3,   # Retefuente 11%
    '23652001': 4,   # Retefuente 10%
    '23652501': 5,   # Retefuente 6%
    '23652503': 6,   # Retefuente 4% (Servicios / Agenciamiento)
    '23654001': 7,   # Retefuente 2.5% (Compras generales / Repuestos)
    '23654004': 18,  # Retefuente 3.5% (Alojamiento / Hoteles)
    '23653502': 19,  # Retefuente 7%
    '23657001': 20,  # Retefuente 2%
    '23652505': 21,  # Retefuente 1%
    '23654006': 23,  # Retefuente Combustibles 0.1%
    '23680501': 8,   # ReteICA 11.04
    '23680503': 9,   # ReteICA 13.8
    '23680505': 10,  # ReteICA 9.66
    '23680507': 11,  # ReteICA 8
    '23680509': 12,  # ReteICA 7
    '23680511': 13,  # ReteICA 6.9
    '23680513': 14,  # ReteICA 4.14
    '23670101': 15,  # ReteIVA 15%
    '24950103': 16   # Impoconsumo 8%
}

AGENTES_ADUANEROS = ["DHL", "ADUANA", "EURO SHIPPING", "PORTUARIA", "ALMACENADORA", "CARGO", "TRADE GLOBAL", "TERMINAL", "BUENAVENTURA"]

def escanear_regimen_texto_pdf(texto):
    """
    Escaneo inteligente de responsabilidades fiscales del EMISOR en la factura electrónica:
    1. Descarta el pie de página de proveedores tecnológicos (Siigo, Facturatech, Carvajal, Cadena, etc.)
       para no confundir las resoluciones de autorretenedor del software con las del proveedor real.
    2. Aísla el bloque de datos del Emisor antes de los datos del Cliente/Adquirente.
    3. Descarta frases negativas como 'No somos grandes contribuyentes ni autorretenedores'.
    4. Detecta afirmativamente O-13 (Gran Contribuyente), O-15 (Autorretenedor), O-47 (Régimen Simple), O-48 y O-49.
    """
    if not texto:
        return ""
    
    txt = texto.upper()
    
    # 1. Cortar pie de página de proveedores tecnológicos de software
    for corte in ["PROVEEDOR TECNOLÓGICO", "PROVEEDOR TECNOLOGICO", "IMPRESO POR SOFTWARE", "DESARROLLADO POR", "SOFTWARE SIIGO", "TECNOLOGÍA TRANSACCIONAL", "THE FACTORY HKA", "DISPAPELES", "FACTURATECH", "CADENA S.A."]:
        pos_c = txt.find(corte)
        if pos_c != -1:
            txt = txt[:pos_c]
            
    # 2. Aislar el bloque del Emisor antes de los datos del Cliente/Adquirente
    pos_cli = -1
    for k_cli in ["DATOS DEL CLIENTE", "DATOS DEL ADQUIRENTE", "CLIENTE:", "SEÑOR(ES):", "SEÑORES:", "ADQUIRENTE:", "FACTURADO A:"]:
        p = txt.find(k_cli)
        if p != -1 and (pos_cli == -1 or p < pos_cli):
            pos_cli = p
            
    txt_emisor = txt[:pos_cli] if pos_cli != -1 else txt
    
    # 3. Analizar Gran Contribuyente (O-13)
    es_gc = False
    if "O-13" in txt_emisor or "0-13" in txt_emisor:
        es_gc = True
    elif "GRAN CONTRIBUYENTE" in txt_emisor or "GRANDES CONTRIBUYENTES" in txt_emisor:
        if not re.search(r"(?:NO\s+(?:SOMOS\s+)?|NI\s+)GRANDES?\s+CONTRIBUYENTES?", txt_emisor):
            es_gc = True
            
    # 4. Analizar Autorretenedor (O-15)
    es_autorr = False
    if "O-15" in txt_emisor or "0-15" in txt_emisor:
        es_autorr = True
    elif "AUTORRETENEDOR" in txt_emisor or "AUTORETENEDOR" in txt_emisor:
        if not re.search(r"(?:NO\s+(?:SOMOS\s+)?|NI\s+)AUTO[R]?RETENEDOR(?:ES)?", txt_emisor):
            es_autorr = True
            
    # 5. Analizar Régimen Simple de Tributación (O-47 / RST)
    es_rst = False
    if "O-47" in txt_emisor or "0-47" in txt_emisor:
        es_rst = True
    elif "REGIMEN SIMPLE" in txt_emisor or "RÉGIMEN SIMPLE" in txt_emisor or "SIMPLE DE TRIBUTACION" in txt_emisor or "RST" in txt_emisor:
        es_rst = True

    codigos = []
    if es_gc: codigos.append("O-13")
    if es_autorr: codigos.append("O-15")
    if es_rst: codigos.append("O-47")
    if not codigos: codigos.append("O-48")
    
    return ";".join(codigos)

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, valor_iva, tipo_doc, resp_emisor="", empresa_compradora=None):
    nombre = str(nombre_emisor).upper()
    resp = str(resp_emisor).upper()
    if empresa_compradora is None:
        empresa_compradora = {"es_gran_contribuyente": False, "es_autorretenedor": False}
        
    tipo_str = str(tipo_doc).upper()
    if "CREDITO" in tipo_str or "CRÉDITO" in tipo_str or "NC" in tipo_str or "DEVOLUCION" in tipo_str:
        t_comp = 17
        op = "Devolucion Compra"
    elif "NOTA" in tipo_str or "AJUSTE" in tipo_str or "CONTABILIDAD" in tipo_str:
        t_comp = 14
        op = "Nota Contabilidad"
    else:
        t_comp = 10
        op = "Compra"
    
    # 1. Agentes Aduaneros e Importaciones
    if any(k in nombre for k in AGENTES_ADUANEROS):
        rfte_t = round(valor_base * 0.04, 2) if valor_base >= 210000 or "EURO SHIPPING" in nombre or "DHL" in nombre else 0.0
        rica_t = round(valor_base * 0.00966, 2) if valor_base >= 210000 and "BUENAVENTURA" not in nombre and "PORTUARIA" not in nombre else 0.0
        cta_p, cta_c, desc = "146505", "22050501", f"Importacion / Transito - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23652503", "24081501", "23680505"
        cat, razon_b = "Importacion (1465)", "Honorarios Agenciamiento vs Terceros"
        
    # 2. Repuestos / Mantenimiento
    elif any(k in nombre for k in ["FERROMENDEZ", "TORNILLOLOCO", "CAUCHOS", "ASIMFER", "MAFLEXCOL", "EMPRECOL", "BATTS ZONE", "MECANIZAR", "HIDRAHULICAS", "ELECTRICOS", "ILUMINACION", "BAMACOLGROUP"]):
        rfte_t = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica_t = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "23654001", "24081001", "23680501"
        if valor_base >= 500000:
            cta_p, cta_c, desc = "14350101", "22050501", f"Repuestos / Inventario - {nombre_emisor[:25]}"
            cat, razon_b = "Inventario", "Compra repuestos mayores a 500k"
        else:
            cta_p, cta_c, desc = "61800101", "23359501", f"Mantenimiento menor - {nombre_emisor[:25]}"
            cat, razon_b = "Costo Mantenimiento", "Repuestos menores a 500k"
            rfte_t, rica_t = 0.0, 0.0
            
    # 3. Hoteles / Viajes
    elif any(k in nombre for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK"]):
        rfte_t = round(valor_base * 0.035, 2) if valor_base >= 210000 else 0.0
        rica_t = round(valor_base * 0.00966, 2) if valor_base >= 210000 else 0.0
        cta_p, cta_c, desc = "51550501", "23359501", f"Alojamiento / Viaje - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23652501", "24081501", "23680505"
        cat, razon_b = "Gasto Viaje", "Hospedaje de personal"
        
    # 4. Software Siigo
    elif "SIIGO" in nombre:
        cta_p, cta_c, desc = "51352001", "23359501", f"Software Siigo - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081501", ""
        cat, razon_b = "Software", "Software y tecnologia"
        
    # 5. Papelería
    elif any(k in nombre for k in ["PANAMERICANA", "LIBRERIA"]):
        cta_p, cta_c, desc = "51953001", "23359501", f"Papeleria - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
        cat, razon_b = "Gastos Papeleria", "Utiles de oficina"
        
    # 6. Compras y Gastos Generales
    elif valor_base >= 500000:
        rfte_t = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica_t = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        cta_p, cta_c, desc = "14350101", "22050501", f"Compra mercancias - {nombre_emisor[:25]}"
        cta_rfte_b, cta_iva, cta_rica_b = "23654001", "24081001", "23680501"
        cat, razon_b = "Mercancia", "Compra general > 500k"
    else:
        cta_p, cta_c, desc = "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}"
        rfte_t, rica_t = 0.0, 0.0
        cta_rfte_b, cta_iva, cta_rica_b = "", "24081001", ""
        cat, razon_b = "Gasto General", "Compra menor general"

    # CRUCE DE RESPONSABILIDADES FISCALES DIAN
    es_emisor_gc = "O-13" in resp or "GRAN CONTRIBUYENTE" in resp
    es_emisor_autorr = "O-15" in resp or "AUTORRETENEDOR" in resp or "SIIGO" in nombre
    es_emisor_rst = "O-47" in resp or "SIMPLE" in resp
    es_comprador_gc = empresa_compradora.get("es_gran_contribuyente", False)

    # A. ReteFuente Renta
    if es_emisor_autorr:
        rfte = 0.0
        cta_rfte = ""
        razon_rfte = "Proveedor Autorretenedor (O-15): No aplica ReteFuente"
    elif es_emisor_rst:
        rfte = 0.0
        cta_rfte = ""
        razon_rfte = "Proveedor Régimen Simple (O-47): No aplica ReteFuente (Art. 911 E.T.)"
    else:
        rfte = rfte_t
        cta_rfte = cta_rfte_b if rfte > 0 else ""
        razon_rfte = f"ReteFuente ordinaria (${rfte:,.2f})" if rfte > 0 else "Base no supera tope"

    # B. ReteIVA (Art. 437-2 E.T.)
    if es_emisor_gc and not es_comprador_gc:
        reteiva = 0.0
        razon_reteiva = "Emisor Gran Contribuyente (O-13): Comprador ordinario no practica ReteIVA"
    elif es_emisor_rst:
        reteiva = 0.0
        razon_reteiva = "Emisor Régimen Simple: Exento de ReteIVA"
    elif es_comprador_gc and valor_iva > 0:
        reteiva = round(valor_iva * 0.15, 2)
        razon_reteiva = "Comprador Gran Contribuyente practica 15% de ReteIVA"
    else:
        reteiva = 0.0
        razon_reteiva = "No aplica ReteIVA"

    # C. ReteICA
    if es_emisor_gc or es_emisor_autorr or es_emisor_rst:
        rica = 0.0
        cta_rica = ""
        razon_rica = "Emisor GC / Autorretenedor: Exento de ReteICA territorial"
    else:
        rica = rica_t
        cta_rica = cta_rica_b if rica > 0 else ""
        razon_rica = f"ReteICA liquidado (${rica:,.2f})" if rica > 0 else "Base no supera tope"

    razon_total = f"{razon_b} | {razon_rfte} | {razon_reteiva}"
    audit_dict = {
        "es_gc": es_emisor_gc,
        "es_autorr": es_emisor_autorr,
        "es_rst": es_emisor_rst,
        "razon_rfte": razon_rfte,
        "razon_reteiva": razon_reteiva,
        "razon_reteica": razon_rica,
        "razon_rica": razon_rica
    }
    return t_comp, op, cta_p, cta_c, desc, rfte, rica, reteiva, cta_rfte, cta_iva, cta_rica, cat, razon_total, audit_dict

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el reporte de facturas, y los PDFs (unificados o separados) para desbloquear, guardar y renombrar automáticamente por comprobante.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel de la DIAN (ej. prueba.xlsx)", type=["xlsx", "xls"])
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF (unificadas o separadas)", type=["pdf"], accept_multiple_files=True)
        
    st.markdown("##### 🔢 3. Consecutivos Iniciales por Tipo de Comprobante Siigo:")
    st.caption("Cada tipo de comprobante lleva su propia numeración independiente. Digita en qué número vas en cada uno:")
    
    c_con1, c_con2, c_con3, c_con4 = st.columns(4)
    with c_con1:
        cons_ini_fac = st.number_input("Facturas de Compra (Tipo 10):", min_value=1, max_value=999999, value=680, step=1, help="Consecutivo inicial para facturas de compra normales.")
    with c_con2:
        cons_ini_nc = st.number_input("Notas Crédito (Tipo 17):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas crédito / devoluciones.")
    with c_con3:
        cons_ini_nota = st.number_input("Notas de Contabilidad (Tipo 14):", min_value=1, max_value=999999, value=1, step=1, help="Consecutivo inicial para notas contables o ajustes.")
    with c_con4:
        clave_pdf_extra = st.text_input("Clave PDF (opcional):", type="password", help="Si algún archivo PDF tiene contraseña específica, ingrésala aquí.")
        
    # Almacenar PDFs originales subidos en session_state y en el historial de inmediato (solo una vez por subida)
    if archivos_pdfs:
        pdfs_sig = f"{len(archivos_pdfs)}_" + "_".join(f"{p.name}_{p.size}" for p in archivos_pdfs[:5])
        if st.session_state.get("_ultimo_pdfs_proc_sig") != pdfs_sig:
            if "raw_uploaded_pdfs" not in st.session_state:
                st.session_state["raw_uploaded_pdfs"] = {}
            for p in archivos_pdfs:
                st.session_state["raw_uploaded_pdfs"][p.name] = p.getvalue()
                
            st.session_state["job_actual_id"] = guardar_trabajo_en_historial(
                empresa,
                st.session_state.get("df_procesado"),
                excel_bytes=st.session_state.get("excel_bytes"),
                excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                zip_bytes=st.session_state.get("zip_pdfs"),
                consecutivo_ini=cons_ini_fac,
                job_id=st.session_state.get("job_actual_id")
            )
            st.session_state["_ultimo_pdfs_proc_sig"] = pdfs_sig
            st.success(f"💾 **{len(archivos_pdfs)} archivo(s) PDF guardados** en el Historial de {empresa['nombre']}.")

    if archivo_excel is not None:
        excel_sig = f"{archivo_excel.name}_{archivo_excel.size}_{cons_ini_fac}_{cons_ini_nc}_{cons_ini_nota}"
        if st.session_state.get("_ultimo_excel_proc_sig") != excel_sig or "df_procesado" not in st.session_state:
            st.session_state["excel_bytes"] = archivo_excel.getvalue()
            st.session_state["excel_nombre"] = archivo_excel.name
            st.session_state["_ultimo_excel_proc_sig"] = excel_sig
            
            df_dian = pd.read_excel(io.BytesIO(st.session_state["excel_bytes"]))

            # Identificar columna de fecha y ordenar cronológicamente de Enero a la fecha actual
            col_fecha = None
            for col_cand in ["Fecha Emisión", "Fecha Emision", "Fecha", "Fecha de Emisión"]:
                if col_cand in df_dian.columns:
                    col_fecha = col_cand
                    break
            if col_fecha is None:
                col_fecha = df_dian.columns[0]

            df_dian["_fecha_dt"] = pd.to_datetime(df_dian[col_fecha], dayfirst=True, errors="coerce")
            df_dian = df_dian.sort_values(by="_fecha_dt", ascending=True).reset_index(drop=True)

            st.success(f"Reporte procesado y ordenado cronológicamente (Enero a la fecha): **{len(df_dian)} facturas identificadas**.")

            contadores_consecutivos = {
                10: int(cons_ini_fac),
                17: int(cons_ini_nc),
                14: int(cons_ini_nota),
                16: int(cons_ini_fac)
            }
            filas = []
            for idx, r in df_dian.iterrows():
                tipo_doc = r.get("Tipo de documento") or r.get("Tipo documento") or "Factura electrónica"
                folio = str(r.get("Folio") or r.get("Factura") or r.get("Factura Num") or f"Doc_{idx+1}").strip()
                if folio.endswith(".0"):
                    folio = folio[:-2]
                prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
                if prefijo == "nan" or prefijo == "None":
                    prefijo = ""

                fecha_val = r.get(col_fecha)
                if pd.notna(fecha_val):
                    try:
                        fecha_dt = pd.to_datetime(fecha_val, dayfirst=True)
                        fecha_str = fecha_dt.strftime("%d/%m/%Y")
                    except:
                        fecha_str = str(fecha_val).split()[0]
                else:
                    fecha_str = "01/01/2026"

                nit_e = str(r.get("NIT Emisor") or r.get("NIT") or "").strip()
                if nit_e.endswith(".0"):
                    nit_e = nit_e[:-2]
                nom_e = str(r.get("Nombre Emisor") or r.get("Proveedor (Emisor)") or r.get("Proveedor") or "Proveedor").strip()

                try: iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
                except: iva = 0.0
                try: tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
                except: tot = 0.0
                base = round(tot - iva, 2)

                col_resp = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["régimen", "regimen", "responsabilidad", "obligacion"])), None)
                resp_e = str(r.get(col_resp, "")).strip() if col_resp and pd.notna(r.get(col_resp)) else ""

                t_comp, op, cta_p, cta_c, desc, rfte, rica, riva, cta_rfte, cta_iva, cta_rica, cat, razon, audit_dict = clasificar_factura(
                    nit_e, nom_e, base, iva, tipo_doc, resp_e, empresa
                )
                consecutivo = contadores_consecutivos.get(t_comp, 1)
                contadores_consecutivos[t_comp] += 1

                nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
                nombre_pdf_esperado = f"Comp_{t_comp}-{consecutivo}_{prefijo}{folio}_{nom_limpio_prov}.pdf"

                filas.append({
                    "N°": idx + 1,
                    "Tipo Comp": t_comp,
                    "Consecutivo": consecutivo,
                    "Comprobante Siigo": f"Comp {t_comp}-{consecutivo}",
                    "Fecha": fecha_str,
                    "Prefijo": prefijo,
                    "Folio": folio,
                    "Factura": f"{prefijo}-{folio}" if prefijo else folio,
                    "Proveedor": nom_e,
                    "NIT Emisor": nit_e,
                    "Régimen Fiscal Emisor": resp_e if resp_e else "O-48 (Estándar)",
                    "Descripcion": desc,
                    "Operacion": op,
                    "Cta Principal": cta_p,
                    "Categoría": cat,
                    "Base": base,
                    "IVA": iva,
                    "ReteFuente": rfte,
                    "ReteICA": rica,
                    "ReteIVA": riva,
                    "Neto a Pagar": round(tot - rfte - rica - riva, 2),
                    "Audit Info": audit_dict,
                    "Cta Contrapartida": cta_c,
                    "Cta IVA": cta_iva,
                    "Cta ReteFuente": cta_rfte,
                    "Cta ReteICA": cta_rica,
                    "Total": tot,
                    "Razón Contable": razon,
                    "Soporte PDF Renombrado": nombre_pdf_esperado
                })

            df_proc = pd.DataFrame(filas)
            st.session_state["df_procesado"] = df_proc

            # Guardar automáticamente en historial de trabajos (Excel + PDFs + Registros)
            st.session_state["job_actual_id"] = guardar_trabajo_en_historial(
                empresa, df_proc,
                excel_bytes=st.session_state.get("excel_bytes"),
                excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                zip_bytes=st.session_state.get("zip_pdfs"),
                consecutivo_ini=cons_ini_fac,
                job_id=st.session_state.get("job_actual_id")
            )
            st.success(f"💾 **Reporte Excel '{archivo_excel.name}' guardado** exitosamente en el Historial de {empresa['nombre']}.")

    if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
        df_proc = st.session_state["df_procesado"]
        
        c_job1, c_job2, c_job3 = st.columns([2.5, 1, 1])
        with c_job1:
            n_orig_mem = len(st.session_state.get("raw_uploaded_pdfs", {}))
            n_renom_mem = len(st.session_state.get("dict_pdfs", {}))
            nom_ex = st.session_state.get("excel_nombre", "Reporte.xlsx")
            st.info(f"📂 **Trabajo Activo:** `{nom_ex}` ({len(df_proc)} facturas) | 📑 **{n_orig_mem} PDFs guardados** | 📄 **{n_renom_mem} procesados**")
        with c_job2:
            if "excel_bytes" in st.session_state and st.session_state["excel_bytes"]:
                st.download_button(
                    label="📥 Descargar Excel",
                    data=st.session_state["excel_bytes"],
                    file_name=st.session_state.get("excel_nombre", "Reporte_Guardado.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_excel_tab_compras",
                    use_container_width=True
                )
        with c_job3:
            if "zip_pdfs" in st.session_state and st.session_state["zip_pdfs"]:
                st.download_button(
                    label="📦 Descargar ZIP Facturas",
                    data=st.session_state["zip_pdfs"],
                    file_name=f"Facturas_{empresa['nombre'].replace(' ', '_')}.zip",
                    mime="application/zip",
                    key="btn_dl_zip_tab_compras",
                    use_container_width=True
                )

    
        c_aud_btn1, c_aud_btn2 = st.columns(2)
        with c_aud_btn1:
            st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronológico Enero - Actual):")
        with c_aud_btn2:
            if st.button("⚡ Auditar y Escanear Régimen Fiscal de los PDFs (Ultrarrápido)", key="btn_escanear_regimen_todos"):
                dict_renom_s = st.session_state.get("dict_pdfs", {})
                dict_orig_s = st.session_state.get("raw_uploaded_pdfs", {})
                total_modificados = 0
                if dict_renom_s or dict_orig_s:
                    with st.spinner("⚡ Extrayendo texto en memoria y cruzando regímenes fiscales..."):
                        # 1. Pre-cargar textos en caché de todos los PDFs una sola vez
                        mapa_textos = []
                        for fn, fb in dict_renom_s.items():
                            mapa_textos.append((fn, [txt.replace("-", "").replace(" ", "").upper() for txt in cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))
                        for fn, fb in dict_orig_s.items():
                            mapa_textos.append((fn, [txt.replace("-", "").replace(" ", "").upper() for txt in cache_extraer_textos_pdf(fb)], cache_extraer_textos_pdf(fb)))
                        
                        # 2. Cruzar en milisegundos cada factura directamente en memoria RAM
                        for r_idx, r_mat in df_proc.iterrows():
                            fol = str(r_mat.get("Folio", "")).replace("-", "").strip().upper()
                            fol_sc = fol.lstrip("0")
                            pref = str(r_mat.get("Prefijo", "")).replace("-", "").strip().upper()
                            fac_full = (pref + fol) if pref else fol
                            fac_full_sc = (pref + fol_sc) if pref else fol_sc
                            nit_d = re.sub(r"\D", "", str(r_mat.get("NIT Emisor", "")))
                            nit_b = nit_d[:-1] if len(nit_d) >= 10 else nit_d
                            
                            texto_factura_encontrado = ""
                            
                            # Buscar en textos pre-extraídos
                            for _, list_clean, list_raw in mapa_textos:
                                for p_idx, t_clean in enumerate(list_clean):
                                    m = False
                                    if len(fac_full) >= 4 and fac_full in t_clean:
                                        m = True
                                    elif len(fac_full_sc) >= 3 and fac_full_sc in t_clean:
                                        m = True
                                    elif (nit_d and len(nit_d) >= 6 and nit_d in t_clean) or (nit_b and len(nit_b) >= 6 and nit_b in t_clean):
                                        if len(fol) >= 2 and fol in t_clean:
                                            m = True
                                        elif len(fol_sc) >= 2 and fol_sc in t_clean:
                                            m = True
                                    if m:
                                        texto_factura_encontrado += list_raw[p_idx] + "\n"
                                if texto_factura_encontrado:
                                    break
                            
                            if texto_factura_encontrado:
                                reg_f = escanear_regimen_texto_pdf(texto_factura_encontrado)
                                if reg_f and reg_f != r_mat.get("Régimen Fiscal Emisor"):
                                    df_proc.at[r_idx, "Régimen Fiscal Emisor"] = reg_f
                                    t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                                        df_proc.at[r_idx, "NIT Emisor"], df_proc.at[r_idx, "Proveedor"],
                                        df_proc.at[r_idx, "Base"], df_proc.at[r_idx, "IVA"],
                                        df_proc.at[r_idx, "Operacion"], reg_f, empresa
                                    )
                                    df_proc.at[r_idx, "ReteFuente"] = rfte_n
                                    df_proc.at[r_idx, "ReteICA"] = rica_n
                                    df_proc.at[r_idx, "ReteIVA"] = riva_n
                                    df_proc.at[r_idx, "Cta ReteFuente"] = c_rf_n
                                    df_proc.at[r_idx, "Cta ReteICA"] = c_ri_n
                                    df_proc.at[r_idx, "Razón Contable"] = razon_n
                                    df_proc.at[r_idx, "Audit Info"] = audit_n
                                    df_proc.at[r_idx, "Neto a Pagar"] = round(df_proc.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                    total_modificados += 1

                    st.session_state["df_procesado"] = df_proc
                    guardar_trabajo_en_historial(
                        empresa, df_proc,
                        excel_bytes=st.session_state.get("excel_bytes"),
                        excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                        dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                        dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                        zip_bytes=st.session_state.get("zip_pdfs"),
                        job_id=st.session_state.get("job_actual_id")
                    )
                    st.success(f"⚡ Auditoría completada en segundos: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su régimen oficial del PDF.")
                    st.rerun()
                else:
                    st.warning("Primero sube los PDFs en el campo 2 para poder escanearlos.")

        st.dataframe(df_proc[["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "Régimen Fiscal Emisor", "Base", "IVA", "ReteFuente", "ReteICA", "ReteIVA", "Neto a Pagar", "Soporte PDF Renombrado"]], use_container_width=True)

        st.markdown("---")
        st.markdown("#### 📑 Procesamiento, Separación y Renombrado de PDFs Multi-Página:")
        st.caption("Identifica y une facturas completas de 2 o más hojas a partir de un PDF consolidado, asignando a cada una su comprobante y consecutivo oficial.")

        col_cfg1, col_cfg2 = st.columns(2)
        with col_cfg1:
            modo_sep = st.radio(
                "Configuración de reconocimiento de facturas:",
                [
                    "🔍 Detección Inteligente (Reconoce facturas completas de 2 o más hojas por 'Página 1 de 2', CUFE y Número)",
                    "📄 Agrupación fija de 2 páginas por factura (cada factura tiene 2 hojas)",
                    "📄 Agrupación fija de 3 páginas por factura (cada factura tiene 3 hojas)",
                    "📑 Separar 1 página por factura (facturas de 1 sola hoja)",
                    "📂 Mantener archivos individuales (sin separar páginas)"
                ]
            )
        with col_cfg2:
            st.info("💡 **Garantía de Factura Completa:** El motor inteligente detecta dónde empieza cada factura (Prefijo, Folio, NIT y marcadores de paginación). Todas las páginas de una misma factura se unen en un solo archivo PDF completo nombrado `Comp_10-XXX_Factura_Proveedor.pdf`.")

        if st.button("🔓 Desbloquear, Separar y Renombrar PDFs ahora"):
            buffer_zip = io.BytesIO()
            total_generados = 0

            if "dict_pdfs" not in st.session_state:
                st.session_state["dict_pdfs"] = {}

            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                df_ref = st.session_state.get("df_procesado", pd.DataFrame())

                # Lista de PDFs a procesar: los recién subidos o los almacenados en la sesión
                pdfs_a_procesar = archivos_pdfs if archivos_pdfs else [
                    io.BytesIO(b_bytes) for b_bytes in st.session_state.get("raw_uploaded_pdfs", {}).values()
                ]

                for idx_pdf, pdf_item in enumerate(pdfs_a_procesar):
                    try:
                        pdf_name = getattr(pdf_item, "name", f"Documento_{idx_pdf+1}.pdf")
                        reader = PdfReader(pdf_item)
                        desencriptado = False
                        if reader.is_encrypted:
                            claves_probar = []
                            if "clave_pdf_extra" in locals() and clave_pdf_extra:
                                claves_probar.append(clave_pdf_extra.strip())

                            # NITs empresa compradora
                            nit_rec = re.sub(r"\D", "", str(empresa.get("nit", "")))
                            if nit_rec:
                                if len(nit_rec) > 9:
                                    claves_probar.extend([nit_rec, nit_rec[:-1], f"{nit_rec[:-1]}-{nit_rec[-1]}"])
                                else:
                                    claves_probar.extend([nit_rec, f"{nit_rec}5", f"{nit_rec}-5"])
                            claves_probar.extend(["901346412", "9013464125", "901346412-5", ""])

                            # NITs emisores en reporte
                            if not df_ref.empty:
                                for n_e in df_ref["NIT Emisor"].dropna().unique():
                                    n_c = re.sub(r"\D", "", str(n_e))
                                    if n_c and len(n_c) >= 7:
                                        claves_probar.append(n_c)

                            for pwd in claves_probar:
                                try:
                                    res_dec = reader.decrypt(pwd)
                                    if res_dec in (1, 2) or not reader.is_encrypted:
                                        desencriptado = True
                                        break
                                except:
                                    pass

                            if reader.is_encrypted and not desencriptado:
                                st.warning(f"⚠️ No se pudo desencriptar automáticamente {pdf_name}. Ingrese la contraseña en el campo correspondiente.")
                                continue

                        num_pags = len(reader.pages)

                        # Caso 1: Archivo de una sola página o modo sin separación
                        if num_pags == 1 or "Mantener archivos individuales" in modo_sep:
                            writer = PdfWriter()
                            texto_pdf = ""
                            for page in reader.pages:
                                writer.add_page(page)
                                try: texto_pdf += page.extract_text() + "\n"
                                except: pass

                            nombre_final = f"Comprobante_{idx_pdf+1}_{pdf_name}"
                            if not df_ref.empty:
                                for _, r_mat in df_ref.iterrows():
                                    folio_m = str(r_mat["Folio"]).strip()
                                    nit_m = str(r_mat["NIT Emisor"]).strip()
                                    if (folio_m and folio_m in texto_pdf) or (nit_m and nit_m in texto_pdf):
                                        nombre_final = r_mat["Soporte PDF Renombrado"]
                                        break

                            # Escaneo de régimen fiscal en el texto del PDF
                            regimen_escaneado = escanear_regimen_texto_pdf(texto_pdf)
                            if regimen_escaneado and not df_ref.empty:
                                for r_idx, r_mat in df_ref.iterrows():
                                    fol_m = str(r_mat["Folio"]).strip()
                                    nit_m = str(r_mat["NIT Emisor"]).strip()
                                    if (fol_m and fol_m in texto_pdf) or (nit_m and nit_m in texto_pdf):
                                        df_ref.at[r_idx, "Régimen Fiscal Emisor"] = regimen_escaneado
                                        _, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _, razon_n, audit_n = clasificar_factura(
                                            df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                                            df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                                            df_ref.at[r_idx, "Operacion"], regimen_escaneado, empresa
                                        )
                                        df_ref.at[r_idx, "ReteFuente"] = rfte_n
                                        df_ref.at[r_idx, "ReteICA"] = rica_n
                                        df_ref.at[r_idx, "ReteIVA"] = riva_n
                                        df_ref.at[r_idx, "Cta ReteFuente"] = cta_rf_n
                                        df_ref.at[r_idx, "Cta ReteICA"] = cta_ri_n
                                        df_ref.at[r_idx, "Razón Contable"] = razon_n
                                        df_ref.at[r_idx, "Audit Info"] = audit_n
                                        df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                        break

                            pdf_bytes = io.BytesIO()
                            writer.write(pdf_bytes)
                            pdf_bytes.seek(0)
                            b_data = pdf_bytes.getvalue()
                            zf.writestr(nombre_final, b_data)
                            st.session_state["dict_pdfs"][nombre_final] = b_data
                            total_generados += 1

                        # Caso 2: Agrupación fija de 2 páginas
                        elif "Agrupación fija de 2 páginas" in modo_sep:
                            for p_i in range(0, num_pags, 2):
                                writer = PdfWriter()
                                writer.add_page(reader.pages[p_i])
                                if p_i + 1 < num_pags:
                                    writer.add_page(reader.pages[p_i + 1])

                                fac_idx = p_i // 2
                                if not df_ref.empty and fac_idx < len(df_ref):
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"]
                                else:
                                    nombre_final = f"Factura_Pags_{p_i+1}-{p_i+2}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 3: Agrupación fija de 3 páginas
                        elif "Agrupación fija de 3 páginas" in modo_sep:
                            for p_i in range(0, num_pags, 3):
                                writer = PdfWriter()
                                for off in range(3):
                                    if p_i + off < num_pags:
                                        writer.add_page(reader.pages[p_i + off])

                                fac_idx = p_i // 3
                                if not df_ref.empty and fac_idx < len(df_ref):
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"]
                                else:
                                    nombre_final = f"Factura_Pags_{p_i+1}-{p_i+3}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 4: 1 página por factura
                        elif "Separar 1 página" in modo_sep:
                            for p_i in range(num_pags):
                                writer = PdfWriter()
                                writer.add_page(reader.pages[p_i])
                                if not df_ref.empty and p_i < len(df_ref):
                                    nombre_final = df_ref.iloc[p_i]["Soporte PDF Renombrado"]
                                else:
                                    nombre_final = f"Factura_Pag_{p_i+1}_{pdf_name}"

                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                pdf_bytes.seek(0)
                                b_data = pdf_bytes.getvalue()
                                zf.writestr(nombre_final, b_data)
                                st.session_state["dict_pdfs"][nombre_final] = b_data
                                total_generados += 1

                        # Caso 5: Detección inteligente por CUFE, Folio o 'Página 1 de N'
                        else:
                            patron_pag = re.compile(r'(?:P[ÁAáa]G(?:INA|\.)?|HOJA|PAGE)\s*(\d+)\s*(?:DE|\/|OF)\s*(\d+)', re.IGNORECASE)
                            curr_writer = None
                            curr_inv_row = None
                            paginas_del_comprobante = 0
                            facturas_generadas = []

                            def buscar_coincidencia_factura(texto_pagina):
                                t_clean = texto_pagina.replace("-", "").replace(" ", "").upper()
                                for _, r_cand in df_ref.iterrows():
                                    fol_cand = str(r_cand["Folio"]).replace("-", "").strip().upper()
                                    pref_cand = str(r_cand["Prefijo"]).replace("-", "").strip().upper()
                                    fac_cand = str(r_cand["Factura"]).replace("-", "").strip().upper()
                                    nit_cand = str(r_cand["NIT Emisor"]).replace("-", "").strip().upper()

                                    if pref_cand and fol_cand and (pref_cand + fol_cand) in t_clean:
                                        return r_cand
                                    if fac_cand and len(fac_cand) >= 4 and fac_cand in t_clean:
                                        return r_cand
                                    if nit_cand and fol_cand and len(fol_cand) >= 3 and (nit_cand in t_clean and fol_cand in t_clean):
                                        return r_cand
                                    if fol_cand and len(fol_cand) >= 4 and fol_cand in t_clean:
                                        return r_cand
                                return None

                            for p_idx in range(num_pags):
                                try: txt_p = reader.pages[p_idx].extract_text() or ""
                                except: txt_p = ""

                                # Escaneo automático del Régimen Fiscal en el texto de la página
                                reg_detectado = escanear_regimen_texto_pdf(txt_p)

                                m_pag = patron_pag.search(txt_p)
                                cur_p, tot_p = (int(m_pag.group(1)), int(m_pag.group(2))) if m_pag else (None, None)

                                inv_encontrada = buscar_coincidencia_factura(txt_p)

                                es_inicio_nueva_factura = False
                                if cur_p == 1:
                                    es_inicio_nueva_factura = True
                                elif cur_p is not None and cur_p > 1:
                                    es_inicio_nueva_factura = False
                                elif inv_encontrada is not None:
                                    if curr_inv_row is None or inv_encontrada["Comprobante Siigo"] != curr_inv_row["Comprobante Siigo"]:
                                        es_inicio_nueva_factura = True
                                elif curr_inv_row is None:
                                    es_inicio_nueva_factura = True

                                if es_inicio_nueva_factura:
                                    if curr_writer is not None and curr_inv_row is not None:
                                        out_c = io.BytesIO()
                                        curr_writer.write(out_c)
                                        c_bytes = out_c.getvalue()
                                        doc_nombre = curr_inv_row["Soporte PDF Renombrado"]
                                        zf.writestr(doc_nombre, c_bytes)
                                        st.session_state["dict_pdfs"][doc_nombre] = c_bytes
                                        total_generados += 1
                                        facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} págs)")

                                    curr_inv_row = inv_encontrada if inv_encontrada is not None else (df_ref.iloc[len(facturas_generadas)] if len(facturas_generadas) < len(df_ref) else None)
                                    if curr_inv_row is not None and reg_detectado and not df_ref.empty:
                                        r_idx = curr_inv_row.name
                                        df_ref.at[r_idx, "Régimen Fiscal Emisor"] = reg_detectado
                                        _, _, _, _, _, rfte_n, rica_n, riva_n, cta_rf_n, _, cta_ri_n, _, razon_n, audit_n = clasificar_factura(
                                            df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                                            df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                                            df_ref.at[r_idx, "Operacion"], reg_detectado, empresa
                                        )
                                        df_ref.at[r_idx, "ReteFuente"] = rfte_n
                                        df_ref.at[r_idx, "ReteICA"] = rica_n
                                        df_ref.at[r_idx, "ReteIVA"] = riva_n
                                        df_ref.at[r_idx, "Cta ReteFuente"] = cta_rf_n
                                        df_ref.at[r_idx, "Cta ReteICA"] = cta_ri_n
                                        df_ref.at[r_idx, "Razón Contable"] = razon_n
                                        df_ref.at[r_idx, "Audit Info"] = audit_n
                                        df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                                    curr_writer = PdfWriter()
                                    curr_writer.add_page(reader.pages[p_idx])
                                    paginas_del_comprobante = 1
                                else:
                                    if curr_writer is None:
                                        curr_inv_row = df_ref.iloc[0] if len(df_ref) > 0 else None
                                        curr_writer = PdfWriter()
                                    curr_writer.add_page(reader.pages[p_idx])
                                    paginas_del_comprobante += 1

                            if curr_writer is not None and curr_inv_row is not None:
                                out_c = io.BytesIO()
                                curr_writer.write(out_c)
                                c_bytes = out_c.getvalue()
                                doc_nombre = curr_inv_row["Soporte PDF Renombrado"]
                                zf.writestr(doc_nombre, c_bytes)
                                st.session_state["dict_pdfs"][doc_nombre] = c_bytes
                                total_generados += 1
                                facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} págs)")

                    except Exception as e:
                        st.error(f"Error procesando {pdf_name}: {e}")

            buffer_zip.seek(0)
            st.session_state["zip_pdfs"] = buffer_zip.getvalue()
            st.session_state["total_zip_pdfs"] = total_generados
            st.session_state["df_procesado"] = df_ref

            # Guardar trabajo actualizado con PDFs procesados
            guardar_trabajo_en_historial(
                empresa, df_ref,
                excel_bytes=st.session_state.get("excel_bytes"),
                excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                zip_bytes=st.session_state["zip_pdfs"],
                consecutivo_ini=cons_ini_fac,
                job_id=st.session_state.get("job_actual_id")
            )
            st.success(f"¡Procesamiento exitoso! Se separaron y renombraron **{total_generados} facturas completas** vinculadas a sus comprobantes y se actualizaron retenciones según el régimen extraído del PDF.")

        if "zip_pdfs" in st.session_state:
            st.download_button(
                label=f"📦 Descargar Archivos PDF Renombrados ({st.session_state['total_zip_pdfs']} facturas completas en ZIP)",
                data=st.session_state["zip_pdfs"],
                file_name=f"Facturas_Renombradas_Comprobantes_{empresa['nombre'].replace(' ', '_')}.zip",
                mime="application/zip",
                use_container_width=True
            )



    else:
        st.info("💡 Sube el reporte Excel de la DIAN arriba o carga un trabajo guardado desde el **Historial de Trabajos** para comenzar.")

with tab_auditoria:
    st.markdown("### Modulo de Auditoria Contable y Trazabilidad")
    st.caption("Inspeccion de cuentas, deducciones y separacion de gastos por cuenta de terceros.")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        opciones_fac = [f"[{r['Comprobante Siigo']}] {r['Fecha']} - {r['Factura']} - {r['Proveedor']} (${r['Total']:,.0f})" for _, r in df_p.iterrows()]
        seleccion = st.selectbox("Selecciona una factura para auditar:", opciones_fac)
        
        comp_sel = seleccion.split("]")[0].replace("[", "")
        fac_sel = df_p[df_p["Comprobante Siigo"] == comp_sel].iloc[0]
        es_aduanero = any(k in fac_sel["Proveedor"].upper() for k in AGENTES_ADUANEROS)
        
        audit_info = fac_sel.get("Audit Info", {})
        if not isinstance(audit_info, dict):
            audit_info = {}
            
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            st.markdown(f"""
            <div class="audit-card">
                <h4>Detalle del Comprobante: {fac_sel['Comprobante Siigo']}</h4>
                <p><b>Soporte PDF Vinculado:</b> <span class="tag-comp">{fac_sel['Soporte PDF Renombrado']}</span></p>
                <p><b>Factura:</b> {fac_sel['Factura']} - <b>Proveedor:</b> {fac_sel['Proveedor']} (NIT: {fac_sel['NIT Emisor']})</p>
                <p><b>Fecha de Emisión:</b> {fac_sel['Fecha']} - <b>Total:</b> ${fac_sel['Total']:,.2f}</p>
                <p><b>Régimen Fiscal Emisor:</b> <span class="badge-active">{fac_sel.get('Régimen Fiscal Emisor', 'O-48')}</span></p>
                <div style="margin: 8px 0;">
                    <span class="badge-tax">{'✅ Gran Contribuyente (O-13)' if audit_info.get('es_gc') else 'Común (No GC)'}</span>
                    <span class="badge-tax">{'✅ Autorretenedor (O-15)' if audit_info.get('es_autorr') else 'No Autorretenedor'}</span>
                    <span class="badge-tax">{'⚠️ Régimen Simple (O-47)' if audit_info.get('es_rst') else 'Régimen Ordinario'}</span>
                </div>
                <hr>
                <h5>Trazabilidad de la Contabilización y Validación DIAN:</h5>
                <ul>
                    <li><b>Cuenta Asignada:</b> <span class="tag-propio">{fac_sel['Cta Principal']}</span> - {fac_sel['Categoría']}</li>
                    <li><b>Cruce DIAN:</b> {fac_sel['Razón Contable']}</li>
                    <li><b>Base Gravable:</b> ${fac_sel['Base']:,.2f} - <b>IVA:</b> ${fac_sel['IVA']:,.2f}</li>
                    <li><b>Retención en la Fuente:</b> ${fac_sel['ReteFuente']:,.2f} ({audit_info.get('razon_rfte', 'N/A')})</li>
                    <li><b>ReteICA:</b> ${fac_sel.get('ReteICA', 0.0):,.2f} ({audit_info.get('razon_reteica', audit_info.get('razon_rica', 'N/A'))})</li>
                    <li><b>ReteIVA:</b> ${fac_sel.get('ReteIVA', 0.0):,.2f} ({audit_info.get('razon_reteiva', 'N/A')})</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("⚙️ Corregir / Ajustar Régimen de este Proveedor manualmente:", expanded=False):
                st.caption("Si la factura física o el RUT indica un régimen diferente, cámbialo aquí para recalcular el asiento al instante:")
                c_mod1, c_mod2 = st.columns(2)
                with c_mod1:
                    opciones_reg_man = [
                        "O-48 (Responsable de IVA / Común - Retención Ordinaria)",
                        "O-15 (Autorretenedor de Renta - Sin ReteFuente)",
                        "O-13 (Gran Contribuyente - Sin ReteIVA)",
                        "O-13;O-15 (Gran Contribuyente y Autorretenedor)",
                        "O-47 (Régimen Simple - Exento ReteFuente y ReteIVA)",
                        "O-49 (No Responsable de IVA)"
                    ]
                    idx_actual = 0
                    reg_actual = str(fac_sel.get("Régimen Fiscal Emisor", ""))
                    if "O-13" in reg_actual and "O-15" in reg_actual: idx_actual = 3
                    elif "O-15" in reg_actual: idx_actual = 1
                    elif "O-13" in reg_actual: idx_actual = 2
                    elif "O-47" in reg_actual: idx_actual = 4
                    elif "O-49" in reg_actual: idx_actual = 5
                    
                    nuevo_reg_sel = st.selectbox("Selecciona Régimen Fiscal:", opciones_reg_man, index=idx_actual, key=f"sel_reg_{fac_sel['Comprobante Siigo']}")
                with c_mod2:
                    st.write("")
                    st.write("")
                    if st.button("💾 Aplicar Cambio", key=f"btn_aplica_reg_{fac_sel['Comprobante Siigo']}"):
                        cod_reg = nuevo_reg_sel.split()[0]
                        r_idx = df_p[df_p["Comprobante Siigo"] == comp_sel].index[0]
                        df_p.at[r_idx, "Régimen Fiscal Emisor"] = cod_reg
                        t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                            df_p.at[r_idx, "NIT Emisor"], df_p.at[r_idx, "Proveedor"],
                            df_p.at[r_idx, "Base"], df_p.at[r_idx, "IVA"],
                            df_p.at[r_idx, "Operacion"], cod_reg, empresa
                        )
                        df_p.at[r_idx, "ReteFuente"] = rfte_n
                        df_p.at[r_idx, "ReteICA"] = rica_n
                        df_p.at[r_idx, "ReteIVA"] = riva_n
                        df_p.at[r_idx, "Cta ReteFuente"] = c_rf_n
                        df_p.at[r_idx, "Cta ReteICA"] = c_ri_n
                        df_p.at[r_idx, "Razón Contable"] = razon_n
                        df_p.at[r_idx, "Audit Info"] = audit_n
                        df_p.at[r_idx, "Neto a Pagar"] = round(df_p.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                        st.session_state["df_procesado"] = df_p
                        guardar_trabajo_en_historial(
                            empresa, df_p,
                            excel_bytes=st.session_state.get("excel_bytes"),
                            excel_nombre=st.session_state.get("excel_nombre", "Reporte.xlsx"),
                            dict_pdfs_renombrados=st.session_state.get("dict_pdfs", {}),
                            dict_pdfs_originales=st.session_state.get("raw_uploaded_pdfs", {}),
                            zip_bytes=st.session_state.get("zip_pdfs"),
                            job_id=st.session_state.get("job_actual_id")
                        )
                        st.success(f"¡Régimen de {fac_sel['Proveedor']} actualizado a {cod_reg} y asiento recalculado!")
                        st.rerun()
            
            if es_aduanero:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("""
                <div style="background:#fef3c7; padding:15px; border-radius:8px; border-left:5px solid #d97706;">
                    <h5 style="color:#92400e; margin-top:0;">Alerta de Importacion / Agenciamiento Aduanero</h5>
                    <p style="color:#78350f; font-size:14px; margin-bottom:5px;">
                    En este documento intervienen <b>gastos por cuenta de terceros</b> y <b>honorarios propios del agente</b>:
                    </p>
                    <ul style="color:#78350f; font-size:14px;">
                        <li><b>Ingresos Propios del Agente (Honorarios / Comision):</b> Gravados con IVA 19%, sujetos a ReteFuente de Servicios (4%) u Honorarios (11%).</li>
                        <li><b>Pagos por Cuenta de Terceros (Tributos / Fletes / Bodegaje):</b> Imputables directamente al costo de importacion (Cuenta 146505). No llevan IVA del agente ni retencion en la fuente sobre el intermediario.</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
                
        with col_a2:
            st.markdown("#### Resumen Financiero")
            st.metric("Base Gravable", f"${fac_sel['Base']:,.2f}")
            st.metric("IVA Liquidado", f"${fac_sel['IVA']:,.2f}")
            st.metric("ReteFuente", f"-${fac_sel['ReteFuente']:,.2f}")
            st.metric("ReteICA", f"-${fac_sel.get('ReteICA', 0.0):,.2f}")
            reteiva_val = fac_sel.get("ReteIVA", 0.0)
            st.metric("ReteIVA", f"-${reteiva_val:,.2f}")
            neto_cxp = round(fac_sel["Total"] - fac_sel["ReteFuente"] - fac_sel.get("ReteICA", 0.0) - reteiva_val, 2)
            st.metric("Total Neto CxP (Cta 22 / 23)", f"${neto_cxp:,.2f}")

        # VISTA PREVIA DE LA FACTURA EN UN CUADRO (SOPORTE PARA UNIFICADOS O SEPARADOS)
        st.markdown("---")
        st.markdown("### 🔍 Vista Previa del Documento Soporte")
        st.caption("Visualiza el PDF de esta factura directamente en pantalla, tanto si subiste archivos individuales como un PDF consolidado/unificado.")

        dict_renom = st.session_state.get("dict_pdfs", {})
        dict_orig = st.session_state.get("raw_uploaded_pdfs", {})
        
        pdf_bytes_encontrado, origen_desc, pags_encontradas = buscar_y_extraer_pdf(
            fac_sel,
            dict_renombrados=dict_renom,
            dict_originales=dict_orig
        )

        if pdf_bytes_encontrado:
            b64_pdf = base64.b64encode(pdf_bytes_encontrado).decode('utf-8')
            
            col_doc1, col_doc2 = st.columns(2)
            with col_doc1:
                st.markdown(f"""
                <div style="background:#0070ba; color:white; padding:8px 14px; border-radius:6px 6px 0 0; font-weight:600; font-size:14px;">
                    📄 {origen_desc} — Factura {fac_sel['Factura']} ({fac_sel['Proveedor']})
                </div>
                """, unsafe_allow_html=True)
            with col_doc2:
                st.download_button(
                    label="📥 Descargar este PDF",
                    data=pdf_bytes_encontrado,
                    file_name=fac_sel["Soporte PDF Renombrado"],
                    mime="application/pdf",
                    use_container_width=True
                )
            
            # Visor robusto que combina PDF.js interactivo con respaldo en iframe nativo
            html_visor = f"""
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="utf-8">
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
              <style>
                body {{
                  margin: 0;
                  padding: 12px;
                  background: #334155;
                  display: flex;
                  flex-direction: column;
                  align-items: center;
                  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                }}
                .page-box {{
                  margin-bottom: 16px;
                  box-shadow: 0 4px 12px rgba(0,0,0,0.35);
                  border-radius: 4px;
                  background: white;
                  overflow: hidden;
                }}
                canvas {{
                  display: block;
                  max-width: 100%;
                  height: auto;
                }}
                #status {{
                  color: #e2e8f0;
                  padding: 10px;
                  font-size: 13px;
                  text-align: center;
                }}
                iframe {{
                  border: none;
                  width: 100%;
                  height: 540px;
                  background: white;
                  border-radius: 4px;
                }}
              </style>
            </head>
            <body>
              <div id="status">Cargando vista previa de la factura...</div>
              <div id="viewer-container"></div>
              <iframe id="fallback-frame" style="display:none;" src="data:application/pdf;base64,{b64_pdf}#toolbar=1&navpanes=0"></iframe>
              <script>
                function showFallback() {{
                  document.getElementById('status').style.display = 'none';
                  document.getElementById('viewer-container').style.display = 'none';
                  document.getElementById('fallback-frame').style.display = 'block';
                }}
                try {{
                  if (typeof pdfjsLib === 'undefined' && window['pdfjs-dist/build/pdf']) {{
                    window.pdfjsLib = window['pdfjs-dist/build/pdf'];
                  }}
                  if (typeof pdfjsLib !== 'undefined') {{
                    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                    const rawPdf = atob("{b64_pdf}");
                    const loadingTask = pdfjsLib.getDocument({{data: rawPdf}});
                    loadingTask.promise.then(function(pdf) {{
                      document.getElementById('status').style.display = 'none';
                      const container = document.getElementById('viewer-container');
                      for (let pNum = 1; pNum <= pdf.numPages; pNum++) {{
                        pdf.getPage(pNum).then(function(page) {{
                          const scale = 1.35;
                          const viewport = page.getViewport({{scale: scale}});
                          const pageBox = document.createElement('div');
                          pageBox.className = 'page-box';
                          const canvas = document.createElement('canvas');
                          const ctx = canvas.getContext('2d');
                          canvas.height = viewport.height;
                          canvas.width = viewport.width;
                          pageBox.appendChild(canvas);
                          container.appendChild(pageBox);
                          page.render({{canvasContext: ctx, viewport: viewport}});
                        }});
                      }}
                    }}).catch(function(err) {{
                      showFallback();
                    }});
                  }} else {{
                    showFallback();
                  }}
                }} catch (e) {{
                  showFallback();
                }}
              </script>
            </body>
            </html>
            """
            components.html(html_visor, height=580, scrolling=True)

        elif dict_orig or dict_renom:
            # Hay PDFs subidos pero no se identificó automáticamente esta factura
            st.warning("⚠️ No se identificó automáticamente el número de esta factura dentro del PDF. Puedes seleccionar manualmente cualquier PDF subido para visualizarlo:")
            todos_los_pdfs = {**dict_orig, **dict_renom}
            pdf_elegido = st.selectbox("Selecciona un archivo PDF cargado:", list(todos_los_pdfs.keys()))
            if pdf_elegido:
                b64_m = base64.b64encode(todos_los_pdfs[pdf_elegido]).decode('utf-8')
                st.download_button(
                    label=f"📥 Descargar {pdf_elegido}",
                    data=todos_los_pdfs[pdf_elegido],
                    file_name=pdf_elegido,
                    mime="application/pdf"
                )
                html_v_man = f"""
                <iframe src="data:application/pdf;base64,{b64_m}#toolbar=1" width="100%" height="540px" style="border:1px solid #cbd5e1; border-radius:6px;"></iframe>
                """
                components.html(html_v_man, height=560, scrolling=True)
        else:
            st.markdown(f"""
            <div style="border: 1px solid #cbd5e1; border-radius: 8px; padding: 18px; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px;">
                <div style="border-bottom: 2px solid #0070ba; padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center;">
                    <h4 style="margin:0; color:#0f172a;">📄 Cuadro de Factura Electrónica</h4>
                    <span style="background:#eff6ff; color:#1d4ed8; padding:3px 10px; border-radius:4px; font-weight:bold; font-family:monospace;">{fac_sel['Factura']}</span>
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 14px; color: #334155;">
                    <div><b>Proveedor (Emisor):</b> {fac_sel['Proveedor']}<br><b>NIT Emisor:</b> {fac_sel['NIT Emisor']}</div>
                    <div><b>Empresa Compradora:</b> {empresa['nombre']}<br><b>NIT Receptor:</b> {empresa['nit']}</div>
                    <div><b>Fecha de Emisión:</b> {fac_sel['Fecha']}</div>
                    <div><b>Valor Total Facturado:</b> ${fac_sel['Total']:,.2f}</div>
                </div>
                <p style="margin-top: 14px; margin-bottom: 0; font-size: 13px; color: #64748b;">
                    <i>💡 Nota: Sube los archivos PDF (unificados o separados) en la Pestaña 1 para ver el documento digitalizado en este visor.</i>
                </p>
            </div>
            """, unsafe_allow_html=True)

        # TABLA DE CONTABILIZACION (ASIENTO CONTABLE)
        st.markdown("### 📋 Asiento Contable del Comprobante (Cómo se Contabilizó)")
        st.caption("Detalle de partida doble con imputación de cuentas, débitos, créditos y sumas iguales:")
        
        asiento_filas = []
        es_nc = "Devolucion" in str(fac_sel["Operacion"])
        
        asiento_filas.append({
            "Código Cuenta": fac_sel["Cta Principal"],
            "Descripción de la Cuenta": f"{fac_sel['Categoría']} - {fac_sel['Proveedor'][:25]}",
            "Tercero / NIT": fac_sel["NIT Emisor"],
            "Débito ($)": 0.0 if es_nc else fac_sel["Base"],
            "Crédito ($)": fac_sel["Base"] if es_nc else 0.0
        })
        
        if fac_sel["IVA"] > 0:
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta IVA"],
                "Descripción de la Cuenta": f"IVA Descontable (Base: ${fac_sel['Base']:,.0f})",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": 0.0 if es_nc else fac_sel["IVA"],
                "Crédito ($)": fac_sel["IVA"] if es_nc else 0.0
            })
            
        if fac_sel["ReteFuente"] > 0 and fac_sel["Cta ReteFuente"]:
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta ReteFuente"],
                "Descripción de la Cuenta": f"ReteFuente Practicada ({fac_sel['Categoría']})",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": fac_sel["ReteFuente"] if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else fac_sel["ReteFuente"]
            })
            
        if fac_sel.get("ReteICA", 0.0) > 0 and fac_sel.get("Cta ReteICA"):
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta ReteICA"],
                "Descripción de la Cuenta": "Retención ICA Practicada",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": fac_sel["ReteICA"] if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else fac_sel["ReteICA"]
            })
            
        if fac_sel.get("ReteIVA", 0.0) > 0:
            asiento_filas.append({
                "Código Cuenta": "23670101",
                "Descripción de la Cuenta": "Retención de IVA Practicada (15%)",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": fac_sel["ReteIVA"] if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else fac_sel["ReteIVA"]
            })
            
        neto_cxp = round(fac_sel["Total"] - fac_sel["ReteFuente"] - fac_sel.get("ReteICA", 0.0) - fac_sel.get("ReteIVA", 0.0), 2)
        asiento_filas.append({
            "Código Cuenta": fac_sel["Cta Contrapartida"],
            "Descripción de la Cuenta": f"Proveedores Nacionales - Fac {fac_sel['Factura']}",
            "Tercero / NIT": fac_sel["NIT Emisor"],
            "Débito ($)": neto_cxp if es_nc else 0.0,
            "Crédito ($)": 0.0 if es_nc else neto_cxp
        })
        
        df_asiento = pd.DataFrame(asiento_filas)
        st.dataframe(
            df_asiento.style.format({"Débito ($)": "${:,.2f}", "Crédito ($)": "${:,.2f}"}),
            use_container_width=True,
            hide_index=True
        )
        
        sum_deb = df_asiento["Débito ($)"].sum()
        sum_cred = df_asiento["Crédito ($)"].sum()
        
        c_as1, c_as2, c_as3 = st.columns(3)
        with c_as1:
            st.metric("Total Débito", f"${sum_deb:,.2f}")
        with c_as2:
            st.metric("Total Crédito", f"${sum_cred:,.2f}")
        with c_as3:
            st.metric("Diferencia / Cuadre", f"${abs(sum_deb - sum_cred):,.2f}")
        st.success("✅ **Comprobante Verificado:** Partida doble cuadrada con sumas iguales.")
    else:
        st.info("Carga el archivo Excel en la Pestana 1 para habilitar la auditoria.")

with tab_siigo:
    st.markdown("### Descargar Planilla Oficial Siigo Nube (3 Hojas)")
    st.caption("Planilla oficial formulada con 'matriz_captura', 'interfaz_siigo' y 'Parametrización'.")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        
        wb = openpyxl.Workbook()
        
        # 1. Hoja matriz_captura
        ws_matriz = wb.active
        ws_matriz.title = "matriz_captura"
        headers_matriz = [
            "Tipo Comp", "Consecutivo", "Fecha (DD/MM/AAAA)", "NIT", "Prefijo", "Factura Num",
            "Descripcion", "Operacion", "Cta Principal", "Valor Base", "IVA", "ReteFuente",
            "ReteICA", "ReteIVA", "Cta Contrapartida"
        ]
        ws_matriz.append(headers_matriz)
        
        # 2. Hoja interfaz_siigo
        ws_interfaz = wb.create_sheet(title="interfaz_siigo")
        headers_interfaz = [
            "Tipo de comprobante", "Consecutivo comprobante", "Fecha de elaboración", "Sigla moneda",
            "Tasa de cambio", "Código cuenta contable", "Identificación tercero", "Sucursal",
            "Código producto", "Código de bodega", "Acción", "Cantidad producto", "Prefijo",
            "Consecutivo", "No. cuota", "Fecha vencimiento", "Código impuesto", "Código grupo activo fijo",
            "Código activo fijo", "Descripción", "Código centro/subcentro de costos", "Débito",
            "Crédito", "Observaciones", "Base gravable libro compras/ventas", "Base exenta libro compras/ventas", "Mes de cierre"
        ]
        ws_interfaz.append(headers_interfaz)
        
        fila_r = 2
        for _, item in df_p.iterrows():
            t_comp = item["Tipo Comp"]
            cons = item["Consecutivo"]
            f_str = item["Fecha"]
            nit = item["NIT Emisor"]
            pref = item["Prefijo"]
            fac_num = item["Folio"]
            desc = item["Descripcion"]
            op = item["Operacion"]
            cta_p = item["Cta Principal"]
            base = item["Base"]
            iva = item["IVA"]
            rfte = item["ReteFuente"]
            rica = item.get("ReteICA", 0.0)
            riva = item.get("ReteIVA", 0.0)
            cta_c = item["Cta Contrapartida"]
            cta_iva = item.get("Cta IVA", "24081001")
            cta_rfte = item.get("Cta ReteFuente", "23654001")
            cta_rica = item.get("Cta ReteICA", "23680501")
            
            ws_matriz.append([
                t_comp, cons, f_str, nit, pref, fac_num,
                desc, op, cta_p, base, iva, rfte, rica, riva, cta_c
            ])
            
            r = fila_r
            # Línea 1: Base Imponible
            ws_interfaz.append([
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                f"=matriz_captura!I{r}", f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "", "", "", "",
                f"=matriz_captura!G{r}", "",
                f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r})',
                f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r}, 0)',
                "", f"=matriz_captura!J{r}", 0.0, ""
            ])
            
            # Línea 2: IVA
            if iva > 0:
                cod_imp_iva = CODIGOS_IMPUESTO_SIIGO.get(str(cta_iva).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_iva, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_iva, "", "", f'="IVA Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!K{r})',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!K{r}, 0)',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])
                
            # Línea 3: ReteFuente
            if rfte > 0 and cta_rfte:
                cod_imp_rfte = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rfte).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_rfte, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_rfte, "", "", f'="ReteFuente Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!L{r}, 0)',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!L{r})',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])
                
            # Línea 4: ReteICA
            if rica > 0 and cta_rica:
                cod_imp_rica = CODIGOS_IMPUESTO_SIIGO.get(str(cta_rica).strip(), "")
                ws_interfaz.append([
                    f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                    cta_rica, f"=matriz_captura!D{r}", 0, "", "", "", "", "", "", "", "",
                    cod_imp_rica, "", "", f'="ReteICA Base: " & matriz_captura!J{r}', "",
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!M{r}, 0)',
                    f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!M{r})',
                    "", f"=matriz_captura!J{r}", 0.0, ""
                ])
                
            # Línea 5: Cuenta por Pagar (Contrapartida)
            formula_deb = f'=IF(matriz_captura!H{r}="Devolucion Compra", matriz_captura!J{r} + matriz_captura!K{r} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r}, 0)'
            formula_cred = f'=IF(matriz_captura!H{r}="Devolucion Compra", 0, matriz_captura!J{r} + matriz_captura!K{r} - matriz_captura!L{r} - matriz_captura!M{r} - matriz_captura!N{r})'
            ws_interfaz.append([
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", f"=matriz_captura!C{r}", "COP", 1,
                f"=matriz_captura!O{r}", f"=matriz_captura!D{r}", 0, "", "", "", "",
                f"=matriz_captura!A{r}", f"=matriz_captura!B{r}", 1, f"=matriz_captura!C{r}",
                "", "", "", f'="Fac " & matriz_captura!E{r} & "-" & matriz_captura!F{r}', "",
                formula_deb, formula_cred, "", 0.0, 0.0, ""
            ])
            fila_r += 1
            
        # 3. Hoja Parametrización
        ws_params = wb.create_sheet(title="Parametrización")
        ws_params.append(["Tipo Comprobante", "", "Impuesto", "Siigo_ID", "Tarifa", "Venta", "Compra", "Dev_Venta", "Dev_Compra"])
        comprobantes = [
            "1 - Ajustes contables", "2 - Depreciación", "3 - Costeo", "4 - Diferidos", "5 - Legalización de viaticos",
            "6 - Legalización de Caja menores", "7 - Obligaciones financieras", "8 - Nómina", "998 - Cierre año",
            "999 - Saldos iniciales", "992 - Comprobante de nómina", "993 - Comprobante de nómina provisión y seguridad social",
            "994 - Comprobante liquidación de contrato", "995 - Comprobante liquidación de primas", "996 - Comprobante liquidación de cesantías",
            "997 - Comprobante desembolso nómina", "9 - Recibo de Caja", "10 - Factura de Compra", "11 - Comprobante de Egreso",
            "12 - Factura de Venta", "13 - Nota Credito", "14 - Nota de Contabilidad", "15 - Documento soporte electronico",
            "777 - Ajustes contables de cartera", "9901 - Traslado de dinero", "16 - FACTURAS DE COMPRA (2)", "17 - NOTA CREDITO DE COMPRA"
        ]
        impuestos = [
            ("IVA 19%", 1, 0.19, "24080601", "24081001", "24082001", "24081002"),
            ("IVA Servicios 19%", "", 0.19, "24080601", "24081501", "24082001", "24082001"),
            ("IVA 5%", 2, 0.05, "24080602", "24081003", "24082002", "24081004"),
            ("Retefuente 11%", 3, 0.11, "13551509", "23651501", "13551510", "23651502"),
            ("Retefuente 10%", 4, 0.1, "13551507", "23652001", "13551508", "23652002"),
            ("Retefuente 6%", 5, 0.06, "13551505", "23652501", "13551506", "23652502"),
            ("Retefuente 4%", 6, 0.04, "13551503", "23652503", "13551504", "23652504"),
            ("Retefuente 3.5%", 18, 0.035, "13551513", "23654004", "13551514", "23654005"),
            ("Retefuente 2.5%", 7, 0.025, "13551501", "23654001", "13551502", "23654002"),
            ("Retefuente Combustibles 0.1%", 23, 0.001, "13551599", "23654006", "13551599", "23654006"),
            ("ReteICA 11.04", 8, 0.01104, "13551801", "23680501", "13551802", "23680502"),
            ("ReteICA 13.8", 9, 0.0138, "13551803", "23680503", "13551804", "23680504"),
            ("ReteICA 9.66", 10, 0.00966, "13551805", "23680505", "13551806", "23680506"),
            ("ReteICA 8", 11, 0.008, "13551807", "23680507", "13551808", "23680508"),
            ("ReteICA 7", 12, 0.007, "13551809", "23680509", "13551810", "23680510"),
            ("ReteICA 6.9", 13, 0.0069, "13551811", "23680511", "13551812", "23680512"),
            ("ReteICA 4.14", 14, 0.00414, "13551813", "23680513", "13551814", "23680514"),
            ("ReteIVA 15%", 15, 0.15, "13551701", "23670101", "13551702", "23670102")
        ]
        max_len = max(len(comprobantes), len(impuestos))
        for idx_row in range(max_len):
            c_val = comprobantes[idx_row] if idx_row < len(comprobantes) else ""
            if idx_row < len(impuestos):
                imp = list(impuestos[idx_row])
                ws_params.append([c_val, ""] + imp)
            else:
                ws_params.append([c_val, "", "", "", "", "", "", "", ""])
                
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label=f"📥 1. Descargar Planilla Oficial Siigo Nube (3 Hojas Formuladas) - {empresa['nombre']}",
            data=output,
            file_name=f"Plantilla_Siigo_{empresa['nombre'].replace(' ', '_')}_3_Hojas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Planilla generada con las 3 hojas oficiales: 'matriz_captura', 'interfaz_siigo' (con códigos de impuesto) y 'Parametrización'.")

        st.markdown("---")
        st.markdown("### 📊 Libro Auxiliar Contable: Facturas Una a Una y Consolidado")
        st.write("Genera el reporte administrativo con el detalle factura a factura y las hojas de consolidado por proveedor y cuentas contables:")

        out_aux = io.BytesIO()
        with pd.ExcelWriter(out_aux, engine='openpyxl') as writer_aux:
            cols_detalle = ["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "NIT Emisor", "Cta Principal", "Categoría", "Base", "IVA", "ReteFuente", "ReteICA", "Total", "Cta Contrapartida", "Razón Contable", "Soporte PDF Renombrado"]
            df_det_export = df_p[cols_detalle].copy()
            
            totales_dict = {
                "Comprobante Siigo": "TOTALES CONSOLIDADOS",
                "Fecha": "-", "Factura": f"{len(df_det_export)} Docs", "Proveedor": "-", "NIT Emisor": "-",
                "Cta Principal": "-", "Categoría": "-",
                "Base": df_det_export["Base"].sum(),
                "IVA": df_det_export["IVA"].sum(),
                "ReteFuente": df_det_export["ReteFuente"].sum(),
                "ReteICA": df_det_export["ReteICA"].sum(),
                "Total": df_det_export["Total"].sum(),
                "Cta Contrapartida": "-", "Razón Contable": "-", "Soporte PDF Renombrado": "-"
            }
            df_det_export = pd.concat([df_det_export, pd.DataFrame([totales_dict])], ignore_index=True)
            df_det_export.to_excel(writer_aux, sheet_name="Facturas_Una_a_Una", index=False)
            
            df_cons_prov = df_p.groupby(["NIT Emisor", "Proveedor"]).agg({
                "Comprobante Siigo": "count",
                "Base": "sum",
                "IVA": "sum",
                "ReteFuente": "sum",
                "ReteICA": "sum",
                "Total": "sum"
            }).reset_index().rename(columns={"Comprobante Siigo": "Cant Facturas"})
            df_cons_prov["Neto CxP"] = df_cons_prov["Total"] - df_cons_prov["ReteFuente"] - df_cons_prov["ReteICA"]
            df_cons_prov = df_cons_prov.sort_values(by="Total", ascending=False).reset_index(drop=True)
            df_cons_prov.to_excel(writer_aux, sheet_name="Consolidado_Proveedores", index=False)
            
            df_cons_cta = df_p.groupby(["Cta Principal", "Categoría"]).agg({
                "Comprobante Siigo": "count",
                "Base": "sum",
                "IVA": "sum",
                "ReteFuente": "sum",
                "ReteICA": "sum",
                "Total": "sum"
            }).reset_index().rename(columns={"Comprobante Siigo": "Cant Facturas"})
            df_cons_cta.to_excel(writer_aux, sheet_name="Consolidado_Cuentas_PUC", index=False)

        out_aux.seek(0)
        st.download_button(
            label=f"📥 2. Descargar Libro de Facturas (Una a Una + Consolidado) - {empresa['nombre']}",
            data=out_aux,
            file_name=f"Libro_Facturas_Detallado_Y_Consolidado_{empresa['nombre'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Primero procesa los documentos en la Pestana 1 para habilitar la descarga.")

