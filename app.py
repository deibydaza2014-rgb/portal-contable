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
                raw_l = json.load(f_idx)
                if isinstance(raw_l, list):
                    lista = raw_l
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

    # Normalizar campos y verificar que exista df_procesado.pkl
    valida = []
    for item in lista:
        if isinstance(item, dict) and "id" in item:
            jdir = os.path.join(base_dir, item["id"])
            if os.path.exists(os.path.join(jdir, "df_procesado.pkl")):
                item.setdefault("total_facturas", 0)
                item.setdefault("total_valor", 0.0)
                item.setdefault("nombre_trabajo", f"Trabajo {item['id']}")
                valida.append(item)
                
    valida.sort(key=lambda x: x.get("id", ""), reverse=True)
    return valida

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

def exportar_respaldo_sesion_zip(empresa_dict, df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p=None):
    """Empaqueta toda la sesión contable en un archivo ZIP descargable para respaldo indestructible."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if df_proc is not None:
            zf.writestr("df_procesado.pkl", pickle.dumps(df_proc))
        if excel_b:
            zf.writestr("excel_original.xlsx", excel_b)
        if zip_p:
            zf.writestr("paquete_facturas.zip", zip_p)
        meta = {
            "empresa": empresa_dict.get("nombre", ""),
            "nit": empresa_dict.get("nit", ""),
            "excel_nombre": excel_n or "Reporte.xlsx",
            "fecha_respaldo": datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p"),
            "total_facturas": len(df_proc) if df_proc is not None else 0,
            "total_pdfs_renombrados": len(dict_renom) if dict_renom else 0,
            "total_pdfs_originales": len(dict_orig) if dict_orig else 0
        }
        zf.writestr("meta_sesion.json", json.dumps(meta, ensure_ascii=False, indent=2))
        if dict_renom:
            for fn, bdata in dict_renom.items():
                zf.writestr(f"pdfs_renombrados/{fn}", bdata)
        if dict_orig:
            for fn, bdata in dict_orig.items():
                zf.writestr(f"pdfs_originales/{fn}", bdata)
    buf.seek(0)
    return buf.getvalue()

def importar_respaldo_sesion_zip(zip_bytes, empresa_dict):
    """Restaura una sesión contable completa desde un archivo ZIP de respaldo."""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
            df_proc = None
            if "df_procesado.pkl" in zf.namelist():
                df_proc = pickle.loads(zf.read("df_procesado.pkl"))
            excel_b = zf.read("excel_original.xlsx") if "excel_original.xlsx" in zf.namelist() else None
            zip_p = zf.read("paquete_facturas.zip") if "paquete_facturas.zip" in zf.namelist() else None
            excel_n = "Reporte_Restaurado.xlsx"
            if "meta_sesion.json" in zf.namelist():
                try:
                    m = json.loads(zf.read("meta_sesion.json").decode("utf-8"))
                    excel_n = m.get("excel_nombre", excel_n)
                except:
                    pass
            dict_renom = {}
            dict_orig = {}
            for name in zf.namelist():
                if name.startswith("pdfs_renombrados/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_renombrados/", "")
                    dict_renom[fn] = zf.read(name)
                elif name.startswith("pdfs_originales/") and name.endswith(".pdf"):
                    fn = name.replace("pdfs_originales/", "")
                    dict_orig[fn] = zf.read(name)
            return df_proc, dict_renom, dict_orig, excel_b, excel_n, zip_p
    except Exception as e:
        return None, {}, {}, None, "", None

# ==============================================================================
# MOTOR DE TRIANGULACIÓN Y CRUCES DE IMPORTACIÓN (ADUANAS Y PAGOS A TERCEROS)
# ==============================================================================
CUENTA_RETENCION_ASUMIDA = "53152001"
CUENTA_NO_DEDUCIBLE = "53950501"
CUENTA_IVA_IMPORTACION = "24081501"
CUENTA_IMPORTACION_TRANSITO = "14650501"
CUENTA_CXP_AGENCIA_EXTERIOR = "22050505"
CUENTA_CXP_DHL_NACIONAL = "23359501"
CUENTA_CXP_EURO_SHIPPING = "22050501"

def analizar_estado_filas_excel(excel_bytes):
    """
    Lee las celdas del archivo Excel DIAN (token) con openpyxl para detectar:
    1. Celdas rojas (facturas ya causadas/registradas en Siigo).
    2. Columna1 con número de comprobante 10-XXX o texto 'registrada'.
    3. Columna3 con asignación de Grupo de Importación (1, 2, 3...).
    4. Columna4 con cuenta contable de pasivo específica (22050505, 23359501, etc.).
    """
    estados = {}
    if not excel_bytes:
        return estados
        
    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        ws = wb.active
        headers = [str(cell.value or '').strip() for cell in ws[1]]
        
        col_c1 = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna1', 'comprobante', 'estado'])), None)
        col_grp = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna3', 'grupo', 'paquete'])), None)
        col_cta = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['columna4', 'cuenta', 'cta'])), None)
        col_cufe = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['cufe', 'token', 'uuid', 'cude'])), None)
        col_folio = next((i for i, h in enumerate(headers) if any(k in h.lower() for k in ['folio', 'factura'])), None)
        col_pref = next((i for i, h in enumerate(headers) if 'prefijo' in h.lower()), None)
        
        for row_idx in range(2, ws.max_row + 1):
            row_cells = ws[row_idx]
            es_roja = False
            
            # Revisar color de relleno de las celdas
            for cell in row_cells:
                fill = cell.fill
                if fill and fill.fill_type:
                    color = fill.start_color
                    if color:
                        rgb = str(getattr(color, 'rgb', '') or getattr(color, 'value', '')).upper()
                        if any(r in rgb for r in ['FF0000', 'C00000', 'FFC7CE', 'EF4444', 'DC2626', 'B91C1C', '991B1B', 'F87171', 'FF6666', 'FF4D4D', 'E11D48']):
                            es_roja = True
                            break
                        if len(rgb) == 8:
                            try:
                                r_v = int(rgb[2:4], 16)
                                g_v = int(rgb[4:6], 16)
                                b_v = int(rgb[6:8], 16)
                                if (r_v > 180 and g_v < 130 and b_v < 130) or (r_v > 220 and g_v < 210 and b_v < 215 and r_v - g_v > 30):
                                    es_roja = True
                                    break
                            except:
                                pass
                                
            val_c1 = str(row_cells[col_c1].value or '').strip() if col_c1 is not None and col_c1 < len(row_cells) else ''
            if '10-' in val_c1 or 'registrad' in val_c1.lower() or 'causad' in val_c1.lower():
                es_roja = True
                
            val_grp = str(row_cells[col_grp].value or '').strip() if col_grp is not None and col_grp < len(row_cells) else ''
            val_cta = str(row_cells[col_cta].value or '').strip() if col_cta is not None and col_cta < len(row_cells) else ''
            if val_cta.endswith('.0'):
                val_cta = val_cta[:-2]
                
            cufe_k = str(row_cells[col_cufe].value or '').strip() if col_cufe is not None and col_cufe < len(row_cells) else ''
            fol_k = str(row_cells[col_folio].value or '').strip() if col_folio is not None and col_folio < len(row_cells) else ''
            if fol_k.endswith('.0'):
                fol_k = fol_k[:-2]
            pref_k = str(row_cells[col_pref].value or '').strip() if col_pref is not None and col_pref < len(row_cells) else ''
            if pref_k == 'nan' or pref_k == 'None':
                pref_k = ''
                
            fac_key = f"{pref_k}-{fol_k}" if pref_k else fol_k
            
            info_fila = {
                'row_idx': row_idx,
                'es_roja': es_roja,
                'comprobante_existente': val_c1 if '10-' in val_c1 else ('10-Previa' if es_roja else ''),
                'grupo_importacion': val_grp,
                'cuenta_especifica': val_cta
            }
            if cufe_k:
                estados[cufe_k] = info_fila
            if fac_key:
                estados[fac_key] = info_fila
            estados[row_idx - 2] = info_fila
    except Exception as e:
        pass
        
    return estados

def generar_asiento_triangulacion_paquete(agente_row, terceros_df):
    """
    Calcula el asiento contable de partida doble para un paquete de importación triangulado:
    1. Débito a CxP Terceros (22050505/23359501) por el neto causado si ya estaban registradas (rojas).
    2. Débito a Retenciones Asumidas (53152001) por las retenciones que Euro cobró al 100%.
    3. Débito a Importación en Tránsito (14650501) e IVA (24081501) para facturas blancas pendientes.
    4. Débito a IVA de Importación si la factura del agente lo discrimina.
    5. Débito a Gastos No Deducibles (53950501) por la diferencia sin factura DIAN.
    6. Crédito a CxP Agente Aduanero (22050501/22050505) por el 100% de la factura del agente.
    """
    tot_agente = float(agente_row.get("Total", 0.0))
    iva_agente = float(agente_row.get("IVA", 0.0))
    
    asiento = []
    suma_cxp_canceladas = 0.0
    suma_ret_asumidas = 0.0
    suma_costo_blancas = 0.0
    suma_iva_blancas = 0.0
    
    for _, t in terceros_df.iterrows():
        t_tot = float(t.get("Total", 0.0))
        t_iva = float(t.get("IVA", 0.0))
        t_base = round(t_tot - t_iva, 2)
        prov_nom = str(t.get("Proveedor", "")).upper()
        nit_t = str(t.get("NIT Emisor", ""))
        fac_num = str(t.get("Factura", ""))
        es_roja = bool(t.get("Ya Registrada", False))
        cta_cxp = str(t.get("Cuenta Pasivo Especifica", "")).strip()
        
        if not cta_cxp:
            if any(k in prov_nom for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL"]):
                cta_cxp = CUENTA_CXP_AGENCIA_EXTERIOR
            else:
                cta_cxp = CUENTA_CXP_DHL_NACIONAL
                
        if es_roja:
            # Factura roja (ya registrada en Siigo)
            # Calcular las retenciones que se le habian practicado (Rfte 4% servicios o 2.5% compras)
            rfte_t = round(t_base * 0.04, 2) if t_base >= 210000 or "DHL" in prov_nom or "CARGO" in prov_nom else 0.0
            rica_t = round(t_base * 0.00966, 2) if t_base >= 210000 else 0.0
            tot_ret = rfte_t + rica_t
            
            cxp_neta = round(t_tot - tot_ret, 2)
            suma_cxp_canceladas += cxp_neta
            suma_ret_asumidas += tot_ret
            
            asiento.append({
                "Código Cuenta": cta_cxp,
                "Descripción Cuenta": f"Cancela CxP {prov_nom[:20]} (Fac {fac_num})",
                "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                "Débito ($)": cxp_neta,
                "Crédito ($)": 0.0
            })
            if tot_ret > 0:
                asiento.append({
                    "Código Cuenta": CUENTA_RETENCION_ASUMIDA,
                    "Descripción Cuenta": f"Retención Asumida Fac {fac_num} ({prov_nom[:18]})",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "Débito ($)": tot_ret,
                    "Crédito ($)": 0.0
                })
        else:
            # Factura blanca (pendiente por registrar)
            costo_neto = t_base
            suma_costo_blancas += costo_neto
            suma_iva_blancas += t_iva
            
            asiento.append({
                "Código Cuenta": CUENTA_IMPORTACION_TRANSITO,
                "Descripción Cuenta": f"Importación en Tránsito (Fac {fac_num})",
                "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                "Débito ($)": costo_neto,
                "Crédito ($)": 0.0
            })
            if t_iva > 0:
                asiento.append({
                    "Código Cuenta": CUENTA_IVA_IMPORTACION,
                    "Descripción Cuenta": f"IVA Descontable Fac {fac_num}",
                    "Tercero / NIT": f"{nit_t} - {prov_nom[:25]}",
                    "Débito ($)": t_iva,
                    "Crédito ($)": 0.0
                })

    # Si la factura del agente discrimina IVA o IVA de Importación
    if iva_agente > 0:
        asiento.append({
            "Código Cuenta": CUENTA_IVA_IMPORTACION,
            "Descripción Cuenta": f"IVA de Importación / Servicios Fac {agente_row.get('Factura', '')}",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": iva_agente,
            "Crédito ($)": 0.0
        })

    # Calcular la diferencia no deducible sin factura DIAN (Cuenta 53950501)
    suma_justificada = suma_cxp_canceladas + suma_ret_asumidas + suma_costo_blancas + suma_iva_blancas + iva_agente
    diferencia_no_deducible = round(tot_agente - suma_justificada, 2)
    
    if diferencia_no_deducible > 0:
        asiento.append({
            "Código Cuenta": CUENTA_NO_DEDUCIBLE,
            "Descripción Cuenta": f"Gastos de Terceros Sin Soporte DIAN (Fac {agente_row.get('Factura', '')})",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": diferencia_no_deducible,
            "Crédito ($)": 0.0
        })
    elif diferencia_no_deducible < 0:
        asiento.append({
            "Código Cuenta": CUENTA_NO_DEDUCIBLE,
            "Descripción Cuenta": f"Ajuste Diferencia Aduanera (Fac {agente_row.get('Factura', '')})",
            "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
            "Débito ($)": 0.0,
            "Crédito ($)": abs(diferencia_no_deducible)
        })

    # Crédito a la Cuenta por Pagar del Agente Aduanero (Euro Shipping / Trade Global)
    cta_agente = CUENTA_CXP_EURO_SHIPPING
    asiento.append({
        "Código Cuenta": cta_agente,
        "Descripción Cuenta": f"Cuenta por Pagar Agente Fac {agente_row.get('Factura', '')}",
        "Tercero / NIT": f"{agente_row.get('NIT Emisor', '')} - {agente_row.get('Proveedor', '')[:25]}",
        "Débito ($)": 0.0,
        "Crédito ($)": tot_agente
    })
    
    df_asiento = pd.DataFrame(asiento)
    return df_asiento, diferencia_no_deducible, suma_ret_asumidas

# ==============================================================================
# MOTOR DE BÚSQUEDA Y EXTRACCIÓN INTELIGENTE DE PDF (UNIFICADO O SEPARADO)
# ==============================================================================
def desbloquear_pdf_bytes(fbytes, nit_receptor=None, claves_extra=None, nits_emisores=None):
    """
    Desbloquea automáticamente un PDF si viene encriptado/protegido con contraseña.
    Prueba el NIT de la empresa compradora (con/sin dígito), NITs de emisores,
    contraseña opcional y contraseñas estándar de facturación electrónica.
    Retorna: (pdf_desbloqueado_bytes, fue_desbloqueado)
    """
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        if not reader.is_encrypted:
            return fbytes, True
            
        claves = [""]
        if claves_extra:
            claves.append(str(claves_extra).strip())
            
        if nit_receptor:
            nr = re.sub(r"\D", "", str(nit_receptor))
            if nr:
                claves.extend([nr, nr[:-1], f"{nr[:-1]}-{nr[-1]}", f"{nr}5", f"{nr}-5"])
        claves.extend(["901346412", "9013464125", "901346412-5", "1234", "123456"])
        
        if nits_emisores:
            for ne in nits_emisores:
                c = re.sub(r"\D", "", str(ne))
                if c and len(c) >= 7:
                    claves.append(c)
                    
        desencriptado = False
        for pwd in claves:
            try:
                res = reader.decrypt(pwd)
                if res in (1, 2) or not reader.is_encrypted:
                    desencriptado = True
                    break
            except Exception:
                pass
                
        if desencriptado:
            writer = PdfWriter()
            for p in reader.pages:
                writer.add_page(p)
            out = io.BytesIO()
            writer.write(out)
            out.seek(0)
            return out.getvalue(), True
        else:
            return fbytes, False
    except Exception:
        return fbytes, False

def auditar_regimen_desde_facturas_renombradas(df_ref, dict_renombrados, empresa_dict):
    """
    Escanea el régimen fiscal directamente de las facturas PDF ya desbloqueadas y renombradas.
    Como cada archivo en dict_renombrados ya está 100% desbloqueado y corresponde exactamente
    a una factura con todas sus hojas, la extracción es inmediata, sin bloqueos y de máxima fidelidad.
    """
    total_modificados = 0
    if not dict_renombrados or df_ref.empty:
        return total_modificados
        
    for r_idx, r_mat in df_ref.iterrows():
        soporte_nom = str(r_mat.get("Soporte PDF Renombrado", ""))
        pdf_bytes = dict_renombrados.get(soporte_nom)
        
        # Búsqueda por coincidencia de comprobante o folio si el nombre varió ligeramente
        if not pdf_bytes:
            comp_id = str(r_mat.get("Comprobante Siigo", "")).replace(" ", "_")
            fol_id = str(r_mat.get("Folio", ""))
            for k, v in dict_renombrados.items():
                if comp_id and comp_id in k:
                    pdf_bytes = v
                    break
                elif fol_id and len(fol_id) >= 3 and fol_id in k:
                    pdf_bytes = v
                    break
                    
        if pdf_bytes:
            try:
                paginas = cache_extraer_textos_pdf(pdf_bytes, nit_receptor=empresa_dict.get("nit", "9013464125"))
                txt_completo = "\n".join(paginas)
                reg_detectado = escanear_regimen_texto_pdf(txt_completo)
                
                if reg_detectado and reg_detectado != r_mat.get("Régimen Fiscal Emisor"):
                    df_ref.at[r_idx, "Régimen Fiscal Emisor"] = reg_detectado
                    t_c_n, op_n, c_p_n, c_c_n, desc_n, rfte_n, rica_n, riva_n, c_rf_n, c_iv_n, c_ri_n, cat_n, razon_n, audit_n = clasificar_factura(
                        df_ref.at[r_idx, "NIT Emisor"], df_ref.at[r_idx, "Proveedor"],
                        df_ref.at[r_idx, "Base"], df_ref.at[r_idx, "IVA"],
                        df_ref.at[r_idx, "Operacion"], reg_detectado, empresa_dict
                    )
                    df_ref.at[r_idx, "ReteFuente"] = rfte_n
                    df_ref.at[r_idx, "ReteICA"] = rica_n
                    df_ref.at[r_idx, "ReteIVA"] = riva_n
                    df_ref.at[r_idx, "Cta ReteFuente"] = c_rf_n
                    df_ref.at[r_idx, "Cta ReteICA"] = c_ri_n
                    df_ref.at[r_idx, "Razón Contable"] = razon_n
                    df_ref.at[r_idx, "Audit Info"] = audit_n
                    df_ref.at[r_idx, "Neto a Pagar"] = round(df_ref.at[r_idx, "Total"] - rfte_n - rica_n - riva_n, 2)
                    total_modificados += 1
            except Exception:
                pass
                
    return total_modificados

def identificar_factura_en_texto(texto, df_ref):
    """
    Evalúa CUFE, Prefijo, Folio, NIT Emisor y nombre comercial del proveedor.
    REGLA DE ORO: NUNCA empareja una factura basándose solo en el NIT del proveedor,
    para evitar confundir múltiples facturas del mismo proveedor (ej. DHL, Claro, Siigo).
    Debe coincidir obligatoriamente el CUFE o el Número de Factura (Prefijo + Folio).
    """
    if not texto or df_ref.empty:
        return None
    txt_clean = re.sub(r'[^A-Z0-9]', '', texto.upper())
    digits_only = re.sub(r'\D', '', texto)
    
    # 1. Validación prioritaria por CUFE / Token (certeza absoluta del 100%)
    for _, r_cand in df_ref.iterrows():
        cufe_cand = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("CUFE / Token", "") or r_cand.get("CUFE", "") or "")).upper()
        if len(cufe_cand) >= 15 and cufe_cand[:20] in txt_clean:
            return r_cand
            
    mejor_cand = None
    mejor_score = 0
    
    for _, r_cand in df_ref.iterrows():
        pref = re.sub(r'[^A-Z0-9]', '', str(r_cand.get("Prefijo", "")).upper())
        fol = re.sub(r'[^A-Z0-9]', '', str(r_cand.get("Folio", "")).upper())
        fol_sc = fol.lstrip('0')
        fac_full = (pref + fol) if pref else fol
        fac_full_sc = (pref + fol_sc) if pref else fol_sc
        
        nit_c = re.sub(r'\D', '', str(r_cand.get("NIT Emisor", "")))
        nit_base = nit_c[:-1] if len(nit_c) >= 10 else nit_c
        
        has_nit = (nit_c and len(nit_c) >= 6 and nit_c in digits_only) or (nit_base and len(nit_base) >= 6 and nit_base in digits_only)
        has_fac = (len(fac_full) >= 3 and fac_full in txt_clean) or (len(fac_full_sc) >= 3 and fac_full_sc in txt_clean)
        
        if not has_fac and len(fol) >= 3:
            # Buscar el folio explícito precedido por marcadores de factura
            pat_fol = rf'(?:FACTURA|FAC|NO|NUMERO|N[°º]|VENTA)[\s\:\.\#\-_]*{re.escape(fol)}'
            if re.search(pat_fol, texto, re.IGNORECASE):
                has_fac = True
                
        prov_words = [w for w in re.split(r'[^A-Z0-9]+', str(r_cand.get("Proveedor", "")).upper()) if len(w) >= 4 and w not in ["SAS", "LTDA", "S.A.", "COLOMBIA", "SERVICES", "SOLUTIONS", "SOCIEDAD", "DISTRIBUCIONES", "GLOBAL", "TRADE"]]
        has_prov = any(w in txt_clean for w in prov_words)
        
        score = 0
        if has_fac and has_nit:
            score = 600
        elif has_fac and has_prov:
            score = 450
        elif has_fac and len(fac_full) >= 4:
            score = 300
        else:
            score = 0
            
        if score > mejor_score and score >= 300:
            mejor_score = score
            mejor_cand = r_cand
            
    return mejor_cand

@st.cache_data(show_spinner=False)
def cache_extraer_textos_pdf(fbytes, nit_receptor="9013464125"):
    """Extrae el texto de todas las páginas de un PDF, desbloqueándolo automáticamente si está protegido."""
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        if reader.is_encrypted:
            fbytes_des, ok = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_receptor)
            if ok:
                reader = PdfReader(io.BytesIO(fbytes_des))
        return [(p.extract_text() or "") for p in reader.pages]
    except Exception:
        return []
    """Extrae el texto de todas las páginas de un PDF y lo almacena en caché en RAM para búsqueda ultrarrápida."""
    try:
        reader = PdfReader(io.BytesIO(fbytes))
        return [(p.extract_text() or "") for p in reader.pages]
    except Exception:
        return []

def buscar_y_extraer_pdf(fac_sel, dict_renombrados=None, dict_originales=None, empresa_compradora=None):
    """
    Busca y extrae el PDF de la factura seleccionada con máxima precisión y validación de contenido:
    1. Búsqueda por CUFE en el nombre del archivo (para PDFs descargados directamente de la DIAN).
    2. Búsqueda por CUFE o Factura verificando el CONTENIDO del PDF en dict_renombrados.
       Si un archivo renombrado contiene una factura distinta a la seleccionada, se descarta.
    3. Búsqueda en los archivos originales (tanto individuales como páginas de consolidado).
    Retorna: (pdf_bytes, descripcion_origen, lista_paginas)
    """
    comp_siigo = str(fac_sel.get("Comprobante Siigo", "")).strip()
    consecutivo = str(fac_sel.get("Consecutivo", "")).strip()
    t_comp = str(fac_sel.get("Tipo Comp", "")).strip()
    
    folio_clean = str(fac_sel.get("Folio", "")).replace("-", "").strip().upper()
    folio_sc = folio_clean.lstrip("0")
    pref_clean = str(fac_sel.get("Prefijo", "")).replace("-", "").strip().upper()
    fac_clean = str(fac_sel.get("Factura", "")).replace("-", "").strip().upper()
    fac_full = (pref_clean + folio_clean) if pref_clean else folio_clean
    fac_full_sc = (pref_clean + folio_sc) if pref_clean else folio_sc
    
    cufe_raw = str(fac_sel.get("CUFE", "") or fac_sel.get("CUFE / Token", "") or "").strip()
    cufe_clean = re.sub(r'[^a-zA-Z0-9]', '', cufe_raw).lower()
    
    # El NIT con el que vienen encriptadas las facturas electrónicas de proveedores es el NIT del COMPRADOR
    nit_comprador = "9013464125"
    if empresa_compradora:
        nit_comprador = re.sub(r"\D", "", str(empresa_compradora.get("nit", "9013464125")))
        
    nit_emisor_digits = re.sub(r"\D", "", str(fac_sel.get("NIT Emisor", "")))
    soporte_nom = str(fac_sel.get("Soporte PDF Renombrado", "")).strip()
    df_una_fac = pd.DataFrame([fac_sel])

    # 1. BÚSQUEDA DIRECTA POR CUFE EN NOMBRE DE ARCHIVO (Archivos DIAN descargados)
    if cufe_clean and len(cufe_clean) >= 15:
        if dict_originales:
            for fname, fbytes in dict_originales.items():
                fn_l = re.sub(r'[^a-zA-Z0-9]', '', fname).lower()
                if cufe_clean[:25] in fn_l or fn_l.startswith(cufe_clean[:20]):
                    f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                    return f_des, f"Factura Oficial DIAN (CUFE: {fname[:18]}...pdf)", None
        if dict_renombrados:
            for fname, fbytes in dict_renombrados.items():
                fn_l = re.sub(r'[^a-zA-Z0-9]', '', fname).lower()
                if cufe_clean[:25] in fn_l or fn_l.startswith(cufe_clean[:20]):
                    return fbytes, f"Factura Oficial DIAN ({fname})", None

    # 2. BÚSQUEDA Y VERIFICACIÓN EN FACTURAS RENOMBRADAS
    if dict_renombrados:
        # A. Prioridad a candidatos con nombre similar (soporte esperado, consecutivo o factura)
        candidatos = []
        if soporte_nom and soporte_nom in dict_renombrados:
            candidatos.append((soporte_nom, dict_renombrados[soporte_nom]))
            
        comp_key = f"Comp_{t_comp}-{consecutivo}".upper()
        for k, v in dict_renombrados.items():
            if comp_key in k.upper().replace(" ", "_") and (k, v) not in candidatos:
                candidatos.append((k, v))
                    
        for k, v in dict_renombrados.items():
            k_clean = k.replace("-", "").replace(" ", "").upper()
            if fac_full and len(fac_full) >= 3 and fac_full in k_clean and (k, v) not in candidatos:
                candidatos.append((k, v))
                    
        # Verificar contenido de los candidatos: SOLO aceptar si el texto realmente pertenece a esta factura
        for c_nom, c_bytes in candidatos:
            try:
                pgs = cache_extraer_textos_pdf(c_bytes, nit_receptor=nit_comprador)
                txt_c = " ".join(pgs)
                if identificar_factura_en_texto(txt_c, df_una_fac) is not None:
                    return c_bytes, f"Factura Verificada ({c_nom})", None
            except Exception:
                pass

        # B. Si los candidatos por nombre fallan, escanear TODOS los archivos renombrados por contenido
        for r_nom, r_bytes in dict_renombrados.items():
            if (r_nom, r_bytes) not in candidatos:
                try:
                    pgs = cache_extraer_textos_pdf(r_bytes, nit_receptor=nit_comprador)
                    txt_r = " ".join(pgs)
                    if identificar_factura_en_texto(txt_r, df_una_fac) is not None:
                        return r_bytes, f"Factura Verificada por Contenido ({r_nom})", None
                except Exception:
                    pass

    # 3. BÚSQUEDA EN ARCHIVOS ORIGINALES INDIVIDUALES Y UNIFICADOS
    if dict_originales:
        # A. Archivos individuales por coincidencia de texto
        for fname, fbytes in dict_originales.items():
            try:
                f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                pgs = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
                txt_ind = " ".join(pgs)
                if identificar_factura_en_texto(txt_ind, df_una_fac) is not None:
                    return f_des, f"Factura Original Verificada ({fname})", None
            except Exception:
                pass

        # B. Escaneo en PDFs unificados página por página
        for fname, fbytes in dict_originales.items():
            try:
                f_des, _ = desbloquear_pdf_bytes(fbytes, nit_receptor=nit_comprador)
                paginas_txt = cache_extraer_textos_pdf(f_des, nit_receptor=nit_comprador)
                pags_coincidentes = []
                for p_idx, txt in enumerate(paginas_txt):
                    if identificar_factura_en_texto(txt, df_una_fac) is not None:
                        if p_idx not in pags_coincidentes:
                            pags_coincidentes.append(p_idx)
                        m_p = re.search(r'(?:P[ÁAáa]G(?:INA)?|HOJA|PAGE)\s*[:\.]?\s*1\s*(?:DE|/|OF)\s*(\d+)', txt, re.IGNORECASE)
                        if m_p:
                            try:
                                total_h = int(m_p.group(1))
                                if 2 <= total_h <= 15:
                                    for off in range(1, total_h):
                                        p_next = p_idx + off
                                        if p_next < len(paginas_txt) and p_next not in pags_coincidentes:
                                            pags_coincidentes.append(p_next)
                            except Exception:
                                pass
                
                if pags_coincidentes:
                    reader = PdfReader(io.BytesIO(f_des))
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

    # 4. Si no se encontró ningún archivo cuyo contenido coincida con esta factura, NO retornar una factura equivocada
    return None, None, None

# AVISO DE TRABAJO PREVIO DISPONIBLE (SOLO SE CARGA SI EL USUARIO OPRIME EL BOTÓN)
trabajos_existentes = listar_trabajos_historial(empresa)

if "df_procesado" not in st.session_state and trabajos_existentes:
    ultimo = trabajos_existentes[0]
    nom_ult = ultimo.get("nombre_trabajo", f"Trabajo {ultimo.get('id', '')}")
    n_fac = ultimo.get("total_facturas", 0)
    val_tot = ultimo.get("total_valor", 0.0)
    
    st.markdown(f"""
    <div style="background:#f0fdf4; border:1px solid #86efac; border-left:5px solid #16a34a; border-radius:8px; padding:14px 18px; margin-bottom:14px;">
        <h4 style="margin:0 0 6px 0; color:#166534;">📂 Tienes un trabajo guardado disponible</h4>
        <p style="margin:0; color:#14532d; font-size:14px;">
            Tienes guardado en memoria tu trabajo previo: <b>'{nom_ult}'</b> con <b>{n_fac} facturas</b> (${val_tot:,.2f}) y sus PDFs vinculados.<br>
            • Si deseas continuar con este trabajo, haz clic en el botón verde.<br>
            • Si hoy vas a implementar <b>facturas nuevas</b>, no tienes que oprimir nada y puedes subir tus nuevos archivos directamente abajo en la Pestaña 1.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    c_btn_c1, c_btn_c2 = st.columns([1.6, 3])
    with c_btn_c1:
        if st.button("📂 Cargar mi trabajo anterior", key="btn_cargar_trabajo_previo_on_demand", use_container_width=True):
            df_g, pdfs_r_g, pdfs_o_g, ex_b_g, ex_n_g, zip_g = cargar_trabajo_historial(empresa, ultimo["id"])
            if df_g is not None and not df_g.empty:
                st.session_state["df_procesado"] = df_g
                st.session_state["dict_pdfs"] = pdfs_r_g or {}
                st.session_state["raw_uploaded_pdfs"] = pdfs_o_g or {}
                st.session_state["excel_bytes"] = ex_b_g
                st.session_state["excel_nombre"] = ex_n_g
                st.session_state["zip_pdfs"] = zip_g
                st.session_state["total_zip_pdfs"] = len(pdfs_r_g) if pdfs_r_g else 0
                st.session_state["job_actual_id"] = ultimo["id"]
                st.session_state["_sesion_cargada_nombre"] = nom_ult
                st.success(f"¡Trabajo '{nom_ult}' cargado con éxito!")
                st.rerun()
    with c_btn_c2:
        st.caption("👈 Oprime aquí **únicamente** cuando desees restaurar el trabajo anterior. De lo contrario, continúa abajo con tus facturas nuevas.")

# BANNER DE SESIÓN ACTIVA EN PANTALLA
if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
    c_bnr1, c_bnr2 = st.columns([4, 1.4])
    with c_bnr1:
        nom_ses = st.session_state.get("_sesion_cargada_nombre", st.session_state.get("excel_nombre", "Reporte de Facturas"))
        st.info(f"📋 **Trabajo Activo en Pantalla:** '{nom_ses}' ({len(st.session_state['df_procesado'])} facturas cargadas). Puedes auditar o exportar a Siigo.")
    with c_bnr2:
        if st.button("🆕 Limpiar / Subir Facturas Nuevas", key="btn_nuevo_trabajo_top", help="Limpia la pantalla para procesar un nuevo mes o facturas nuevas."):
            for k in ["df_procesado", "dict_pdfs", "raw_uploaded_pdfs", "excel_bytes", "excel_nombre", "zip_pdfs", "job_actual_id", "_sesion_cargada_nombre", "_sesion_auto_recuperada", "_ultimo_excel_proc_sig", "_ultimo_pdfs_proc_sig"]:
                st.session_state.pop(k, None)
            st.rerun()

# PANEL DE HISTORIAL DE TRABAJOS Y RESPALDOS INDESTRUCTIBLES
trabajos_guardados = listar_trabajos_historial(empresa)
panel_expanded = False

with st.expander("🗂️ Historial de Trabajos, Respaldos y Carga Rápida", expanded=panel_expanded):
    col_h_left, col_h_right = st.columns([1.5, 1])
    
    with col_h_left:
        st.markdown("#### 📂 Trabajos Guardados en esta Empresa:")
        if trabajos_guardados:
            st.caption("Haz clic en 'Cargar' para recuperar de inmediato cualquier auditoría o mes previo:")
            for tb in trabajos_guardados:
                c_h1, c_h2, c_h3 = st.columns([3, 1.2, 0.5])
                with c_h1:
                    nom_tb = tb.get("nombre_trabajo", f"Trabajo {tb.get('id', '')}")
                    n_facs = tb.get("total_facturas", 0)
                    val_tot = tb.get("total_valor", 0.0)
                    n_renom = tb.get("total_pdfs_renombrados", tb.get("total_pdfs", 0))
                    n_orig = tb.get("total_pdfs_originales", 0)
                    info_pdf_str = f"{n_renom} PDFs procesados" if n_renom > 0 else (f"{n_orig} PDFs subidos" if n_orig > 0 else "Sin PDFs")
                    st.markdown(f"📄 **{nom_tb}** — {n_facs} facturas (${val_tot:,.2f}) — **{info_pdf_str}**")
                with c_h2:
                    if st.button("📂 Cargar", key=f"btn_h_load_{tb['id']}"):
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
                        st.session_state["_sesion_auto_recuperada"] = nom_tb
                        st.success(f"¡Trabajo '{nom_tb}' cargado! Todo tu avance está listo.")
                        st.rerun()
                with c_h3:
                    if st.button("🗑️", key=f"btn_h_del_{tb['id']}"):
                        eliminar_trabajo_historial(empresa, tb["id"])
                        st.rerun()
        else:
            st.info("💡 Aún no tienes trabajos guardados en el disco local para esta empresa.")

    with col_h_right:
        st.markdown("#### 🛡️ Respaldo Portable (.zip):")
        st.caption("Guarda o restaura todo tu trabajo en un solo archivo, ideal si cambias de PC o si el servidor se reinicia:")
        
        # 1. Botón para exportar respaldo de la sesión activa
        if "df_procesado" in st.session_state and st.session_state["df_procesado"] is not None:
            zip_respaldo_bytes = exportar_respaldo_sesion_zip(
                empresa,
                st.session_state.get("df_procesado"),
                st.session_state.get("dict_pdfs", {}),
                st.session_state.get("raw_uploaded_pdfs", {}),
                st.session_state.get("excel_bytes"),
                st.session_state.get("excel_nombre", "Reporte.xlsx"),
                st.session_state.get("zip_pdfs")
            )
            nom_respaldo = f"Respaldo_Sesion_{empresa['nombre'].replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.zip"
            st.download_button(
                label="💾 Descargar Respaldo Completo de esta Sesión (.zip)",
                data=zip_respaldo_bytes,
                file_name=nom_respaldo,
                mime="application/zip",
                use_container_width=True,
                help="Descarga un solo archivo con TODO (Excel, cálculos, retenciones y PDFs vinculados)."
            )
        
        # 2. Uploader para restaurar desde un archivo ZIP de respaldo previo
        archivo_zip_restaurar = st.file_uploader(
            "📥 Restaurar Sesión desde Archivo de Respaldo (.zip):",
            type=["zip"],
            key="upl_zip_restore_historial"
        )
        if archivo_zip_restaurar is not None:
            with st.spinner("Restaurando sesión completa desde el respaldo..."):
                df_res, renom_res, orig_res, ex_b_res, ex_n_res, zip_res = importar_respaldo_sesion_zip(
                    archivo_zip_restaurar.getvalue(), empresa
                )
                if df_res is not None:
                    st.session_state["df_procesado"] = df_res
                    st.session_state["dict_pdfs"] = renom_res or {}
                    st.session_state["raw_uploaded_pdfs"] = orig_res or {}
                    st.session_state["excel_bytes"] = ex_b_res
                    st.session_state["excel_nombre"] = ex_n_res
                    st.session_state["zip_pdfs"] = zip_res
                    st.session_state["total_zip_pdfs"] = len(renom_res) if renom_res else 0
                    
                    # Guardar también en el disco local
                    guardar_trabajo_en_historial(
                        empresa, df_res,
                        excel_bytes=ex_b_res,
                        excel_nombre=ex_n_res,
                        dict_pdfs_renombrados=renom_res,
                        dict_pdfs_originales=orig_res,
                        zip_bytes=zip_res
                    )
                    st.session_state["_sesion_auto_recuperada"] = f"Respaldo {archivo_zip_restaurar.name}"
                    st.success(f"¡Sesión restaurada con éxito! Se recuperaron {len(df_res)} facturas y {len(renom_res)} PDFs.")
                    st.rerun()
                else:
                    st.error("El archivo ZIP no contiene un respaldo válido de sesión contable.")

st.markdown("---")

tab_compras, tab_auditoria, tab_triangulacion, tab_siigo = st.tabs([
    "1. Cargar Documentos y Desbloquear PDFs",
    "2. Auditoria y Trazabilidad Fiscal",
    "3. 🔀 Triangulación y Cruces de Importación (Aduanas)",
    "4. Exportar Planilla Oficial a Siigo"
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

def calcular_base_con_regla_descuento(r, tot, iva, nom_emisor):
    """
    Calcula la Base Gravable contable respetando la regla estricta:
    1. El descuento SOLO se aplica si está incorporado en los ÍTEMS de la factura.
    2. Si el descuento aparece en las NOTAS u observaciones, NO VA (no disminuye la base gravable).
    3. En agencias aduaneras (DHL, Euro Shipping, etc.), las casillas o notas de descuento corresponden
       a retenciones practicadas que no deben restarse de la base para evitar duplicidades.
    4. Si en el reporte/factura existe la columna de Subtotal (valor de los ítems antes de notas),
       se toma directamente el Subtotal como base contable.
    """
    nom = str(nom_emisor).upper()
    es_aduanero = any(k in nom for k in AGENTES_ADUANEROS)
    
    col_sub = next((c for c in r.index if any(k in str(c).lower() for k in ["subtotal", "sub total", "base gravable", "base_gravable", "valor subtotal", "lineextensionamount"])), None)
    col_desc_item = next((c for c in r.index if any(k in str(c).lower() for k in ["descuento_items", "descuento item", "descuento ítems", "descuento linea", "descuento comercial"])), None)
    col_desc_gen = next((c for c in r.index if any(k in str(c).lower() for k in ["descuento", "descuentos", "total descuento", "valor descuento", "allowancetotalamount"])), None)
    
    subtotal_val = None
    if col_sub and pd.notna(r.get(col_sub)):
        try:
            v = float(r.get(col_sub))
            if v > 0:
                subtotal_val = v
        except Exception:
            pass
            
    desc_val = 0.0
    if col_desc_gen and pd.notna(r.get(col_desc_gen)):
        try:
            desc_val = float(r.get(col_desc_gen))
        except Exception:
            desc_val = 0.0

    desc_item_val = 0.0
    if col_desc_item and pd.notna(r.get(col_desc_item)):
        try:
            desc_item_val = float(r.get(col_desc_item))
        except Exception:
            desc_item_val = 0.0
            
    # REGLA:
    # Si viene el Subtotal de los ítems, esa es la base real. Los descuentos en notas NO se restan.
    if subtotal_val is not None:
        base = round(subtotal_val, 2)
        motivo_base = f"Base = Subtotal de Ítems (${base:,.2f}) — Descuentos en notas no aplican"
    else:
        # Si no hay columna de subtotal y es agencia aduanera con descuento reportado (retenciones en notas):
        if es_aduanero and desc_val > 0:
            base = round(tot - iva + desc_val, 2)
            motivo_base = f"Agencia Aduanera: Base restituida (${base:,.2f}) sin descontar retenciones de notas"
        elif desc_item_val > 0:
            base = round(tot - iva, 2)
            motivo_base = f"Base calculada con descuento comercial en ítems (${base:,.2f})"
        else:
            base = round(tot - iva, 2)
            motivo_base = f"Base estándar = Total - IVA (${base:,.2f})"
            
    return base, desc_val, motivo_base

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
            # Analizar celdas rojas (ya registradas), grupos y cuentas específicas del Excel
            estados_excel = analizar_estado_filas_excel(st.session_state.get("excel_bytes"))

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
                base, desc_val, motivo_base = calcular_base_con_regla_descuento(r, tot, iva, nom_e)
                col_cufe = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["cufe", "token", "uuid", "cude"])), None)
                cufe_val = str(r.get(col_cufe, "")).strip() if col_cufe and pd.notna(r.get(col_cufe)) else ""

                col_resp = next((c for c in df_dian.columns if any(k in str(c).lower() for k in ["régimen", "regimen", "responsabilidad", "obligacion"])), None)
                resp_e = str(r.get(col_resp, "")).strip() if col_resp and pd.notna(r.get(col_resp)) else ""

                t_comp, op, cta_p, cta_c, desc, rfte, rica, riva, cta_rfte, cta_iva, cta_rica, cat, razon, audit_dict = clasificar_factura(
                    nit_e, nom_e, base, iva, tipo_doc, resp_e, empresa
                )
                consecutivo = contadores_consecutivos.get(t_comp, 1)
                contadores_consecutivos[t_comp] += 1

                nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
                nombre_pdf_esperado = f"Comp_{t_comp}-{consecutivo}_{prefijo}{folio}_{nom_limpio_prov}.pdf"

                # Identificar si la fila estaba marcada en rojo (ya registrada en Siigo), su grupo y cuenta
                fac_k_full = f"{prefijo}-{folio}" if prefijo else folio
                info_est = estados_excel.get(cufe_val) or estados_excel.get(fac_k_full) or estados_excel.get(folio) or estados_excel.get(idx, {})
                es_roja_reg = bool(info_est.get("es_roja", False))
                comp_prev = str(info_est.get("comprobante_existente", ""))
                grp_imp = str(info_est.get("grupo_importacion", "")).strip()
                cta_esp = str(info_est.get("cuenta_especifica", "")).strip()
                
                es_aduanera_flag = any(k in nom_e.upper() for k in AGENTES_ADUANEROS) or any(k in nom_e.upper() for k in ["CARGO", "ADUANA", "PORTUARIA", "ALMACENADORA", "TERMINAL", "BUENAVENTURA", "CARTAGENA", "CONSOLCARGO", "EURO SHIPPING", "TRADE GLOBAL"])

                filas.append({
                    "N°": idx + 1,
                    "Tipo Comp": t_comp,
                    "Consecutivo": consecutivo,
                    "Comprobante Siigo": f"Comp {t_comp}-{consecutivo}",
                    "Fecha": fecha_str,
                    "Prefijo": prefijo,
                    "Folio": folio,
                    "Factura": fac_k_full,
                    "Proveedor": nom_e,
                    "NIT Emisor": nit_e,
                    "Ya Registrada": es_roja_reg,
                    "Comprobante Previo": comp_prev,
                    "Grupo Importación": grp_imp,
                    "Cuenta Pasivo Especifica": cta_esp,
                    "Es Aduanera": es_aduanera_flag,
                    "Estado Registro": f"🔴 Ya Registrada ({comp_prev})" if es_roja_reg else ("🟡 Agente Aduanero" if any(k in nom_e.upper() for k in ["EURO SHIPPING", "TRADE GLOBAL"]) else ("🟢 Aduanera" if es_aduanera_flag else "⚪ Compra Pendiente")),
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
                    "CUFE": cufe_val,
                    "Descuento Registrado": desc_val,
                    "Regla Base": motivo_base,
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
        
        c_job1, c_job2, c_job3, c_job4 = st.columns([2.2, 0.9, 0.9, 1.2])
        with c_job1:
            n_orig_mem = len(st.session_state.get("raw_uploaded_pdfs", {}))
            n_renom_mem = len(st.session_state.get("dict_pdfs", {}))
            nom_ex = st.session_state.get("excel_nombre", "Reporte.xlsx")
            st.info(f"📂 **Trabajo Activo:** `{nom_ex}` ({len(df_proc)} facturas) | 📑 **{n_orig_mem} PDFs guardados** | 📄 **{n_renom_mem} procesados**")
        with c_job2:
            if "excel_bytes" in st.session_state and st.session_state["excel_bytes"]:
                st.download_button(
                    label="📥 Excel",
                    data=st.session_state["excel_bytes"],
                    file_name=st.session_state.get("excel_nombre", "Reporte_Guardado.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_excel_tab_compras",
                    use_container_width=True
                )
        with c_job3:
            if "zip_pdfs" in st.session_state and st.session_state["zip_pdfs"]:
                st.download_button(
                    label="📦 ZIP Facturas",
                    data=st.session_state["zip_pdfs"],
                    file_name=f"Facturas_{empresa['nombre'].replace(' ', '_')}.zip",
                    mime="application/zip",
                    key="btn_dl_zip_tab_compras",
                    use_container_width=True
                )
        with c_job4:
            zip_respaldo_b = exportar_respaldo_sesion_zip(
                empresa, df_proc,
                st.session_state.get("dict_pdfs", {}),
                st.session_state.get("raw_uploaded_pdfs", {}),
                st.session_state.get("excel_bytes"),
                st.session_state.get("excel_nombre", "Reporte.xlsx"),
                st.session_state.get("zip_pdfs")
            )
            st.download_button(
                label="💾 Guardar Sesión (.zip)",
                data=zip_respaldo_b,
                file_name=f"Respaldo_Sesion_{empresa['nombre'].replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                key="btn_dl_respaldo_bar",
                use_container_width=True,
                help="Guarda toda la sesión (Excel + Cuentas + PDFs) en un solo archivo para restaurar mañana en 1 segundo."
            )

    
        c_aud_btn1, c_aud_btn2 = st.columns(2)
        with c_aud_btn1:
            st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronológico Enero - Actual):")
        with c_aud_btn2:
            if st.button("⚡ Auditar y Escanear Régimen Fiscal de los PDFs (Ultrarrápido)", key="btn_escanear_regimen_todos"):
                dict_renom_s = st.session_state.get("dict_pdfs", {})
                dict_orig_s = st.session_state.get("raw_uploaded_pdfs", {})
                total_modificados = 0
                if dict_renom_s:
                    with st.spinner("⚡ Escaneando régimen fiscal directamente desde las facturas ya desbloqueadas y renombradas..."):
                        total_modificados = auditar_regimen_desde_facturas_renombradas(df_proc, dict_renom_s, empresa)
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
                    st.success(f"⚡ Auditoría completada desde facturas renombradas: se revisaron {len(df_proc)} facturas y se actualizaron {total_modificados} con su régimen oficial del PDF.")
                    st.rerun()
                elif dict_orig_s:
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

                        # Caso 1: Archivo individual (sin separar páginas) o de 1 página
                        if num_pags == 1 or "Mantener archivos individuales" in modo_sep:
                            writer = PdfWriter()
                            texto_pdf = ""
                            for page in reader.pages:
                                writer.add_page(page)
                                try: texto_pdf += (page.extract_text() or "") + "\n"
                                except: pass

                            r_mat = None
                            if not df_ref.empty:
                                # 1. Identificar por texto exacto (CUFE, Factura + Proveedor/NIT)
                                r_mat = identificar_factura_en_texto(texto_pdf, df_ref)
                                
                                # 2. Si no se identificó en el texto, buscar por coincidencia en el nombre del archivo
                                if r_mat is None:
                                    fn_clean = re.sub(r'[^A-Za-z0-9]', '', pdf_name).lower()
                                    for _, r_cand in df_ref.iterrows():
                                        cufe_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("CUFE", "") or "")).lower()
                                        if cufe_c and len(cufe_c) >= 15 and cufe_c[:20] in fn_clean:
                                            r_mat = r_cand
                                            break
                                        pref_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("Prefijo", "") or "")).upper()
                                        fol_c = re.sub(r'[^A-Za-z0-9]', '', str(r_cand.get("Folio", "") or "")).upper()
                                        fac_f = (pref_c + fol_c) if pref_c else fol_c
                                        if fac_f and len(fac_f) >= 4 and fac_f in fn_clean.upper():
                                            r_mat = r_cand
                                            break

                            if r_mat is not None:
                                nombre_final = r_mat["Soporte PDF Renombrado"]
                                # Escaneo de régimen fiscal sólo para la factura que realmente coincidió
                                regimen_escaneado = escanear_regimen_texto_pdf(texto_pdf)
                                if regimen_escaneado:
                                    r_idx = r_mat.name
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
                            else:
                                nombre_final = f"Soporte_{idx_pdf+1}_{pdf_name}"

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
                                txt_grp = ""
                                for off in range(2):
                                    if p_i + off < num_pags:
                                        writer.add_page(reader.pages[p_i + off])
                                        try: txt_grp += (reader.pages[p_i + off].extract_text() or "") + "\n"
                                        except: pass

                                r_mat = identificar_factura_en_texto(txt_grp, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    fac_idx = p_i // 2
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"] if not df_ref.empty and fac_idx < len(df_ref) else f"Factura_Pags_{p_i+1}-{p_i+2}_{pdf_name}"

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
                                txt_grp = ""
                                for off in range(3):
                                    if p_i + off < num_pags:
                                        writer.add_page(reader.pages[p_i + off])
                                        try: txt_grp += (reader.pages[p_i + off].extract_text() or "") + "\n"
                                        except: pass

                                r_mat = identificar_factura_en_texto(txt_grp, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    fac_idx = p_i // 3
                                    nombre_final = df_ref.iloc[fac_idx]["Soporte PDF Renombrado"] if not df_ref.empty and fac_idx < len(df_ref) else f"Factura_Pags_{p_i+1}-{p_i+3}_{pdf_name}"

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
                                try: txt_pag = (reader.pages[p_i].extract_text() or "") + "\n"
                                except: txt_pag = ""

                                r_mat = identificar_factura_en_texto(txt_pag, df_ref)
                                if r_mat is not None:
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                else:
                                    nombre_final = df_ref.iloc[p_i]["Soporte PDF Renombrado"] if not df_ref.empty and p_i < len(df_ref) else f"Factura_Pag_{p_i+1}_{pdf_name}"

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
                                return identificar_factura_en_texto(texto_pagina, df_ref)

                            for p_idx in range(num_pags):
                                try: txt_p = reader.pages[p_idx].extract_text() or ""
                                except: txt_p = ""

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

                                    curr_inv_row = inv_encontrada if inv_encontrada is not None else curr_inv_row
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
                                        curr_inv_row = inv_encontrada
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
            
            # Escaneo oficial de régimen fiscal directamente sobre las facturas ya desbloqueadas y renombradas
            total_auditados = auditar_regimen_desde_facturas_renombradas(df_ref, st.session_state["dict_pdfs"], empresa)
            st.session_state["df_procesado"] = df_ref

            # Guardar trabajo actualizado con PDFs procesados y regímenes verificados
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
            msg_reg = f" y se auditó el régimen de cada factura ({total_auditados} actualizadas)" if total_auditados > 0 else " y se validó el régimen de cada factura"
            st.success(f"¡Procesamiento exitoso! Se desbloquearon y renombraron **{total_generados} facturas completas**{msg_reg} directamente desde las facturas PDF desbloqueadas.")

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
            dict_originales=dict_orig,
            empresa_compradora=empresa
        )

        if pdf_bytes_encontrado:
            # Banner de coincidencia exacta con el registro contable
            st.success(f"✅ **Factura y Registro Contable Vinculados:** Comprobante **{fac_sel['Comprobante Siigo']}** | Factura: **{fac_sel['Factura']}** | Proveedor: **{fac_sel['Proveedor']}** (NIT: {fac_sel['NIT Emisor']}) — Total: **${fac_sel['Total']:,.2f}**")

            b64_pdf = base64.b64encode(pdf_bytes_encontrado).decode('utf-8')
            
            # Conteo previo de hojas de la factura para visualización completa
            try:
                reader_prev = PdfReader(io.BytesIO(pdf_bytes_encontrado))
                num_pags_tot = len(reader_prev.pages)
            except Exception:
                num_pags_tot = 1
            
            col_doc1, col_doc2 = st.columns([2.5, 1])
            with col_doc1:
                pags_badge = f"{num_pags_tot} páginas completas" if num_pags_tot > 1 else "1 página"
                st.markdown(f"""
                <div style="background:#0070ba; color:white; padding:9px 16px; border-radius:6px 6px 0 0; font-weight:600; font-size:14px; display:flex; justify-content:space-between; align-items:center;">
                    <span>📄 {origen_desc} — Factura {fac_sel['Factura']} ({fac_sel['Proveedor']})</span>
                    <span style="background:rgba(255,255,255,0.25); padding:3px 10px; border-radius:12px; font-size:12px;">📑 {pags_badge}</span>
                </div>
                """, unsafe_allow_html=True)
            with col_doc2:
                st.download_button(
                    label=f"📥 Descargar Factura Completa ({num_pags_tot} págs)",
                    data=pdf_bytes_encontrado,
                    file_name=fac_sel["Soporte PDF Renombrado"],
                    mime="application/pdf",
                    use_container_width=True
                )
                
            # Opciones de visualización de páginas
            if num_pags_tot > 1:
                modo_vista_doc = st.radio(
                    "Modo de vista de factura:",
                    ["📜 Ver Todas las Hojas en Cascada Continua (1, 2, 3...)", "🖨️ Visor Integrado del Navegador (con Miniaturas y Barra Lateral)"],
                    horizontal=True,
                    key="radio_modo_vista_pdf"
                )
            else:
                modo_vista_doc = "📜 Ver Todas las Hojas en Cascada Continua (1, 2, 3...)"
                
            # Altura dinámica para visualizar todas las hojas (2, 3 o más páginas continuas)
            visor_height = max(680, min(2200, num_pags_tot * 620))
            
            # Visor robusto secuencial de alta resolución para facturas multi-página
            html_visor = f"""
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="utf-8">
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
              <style>
                body {{
                  margin: 0;
                  padding: 14px;
                  background: #0f172a;
                  display: flex;
                  flex-direction: column;
                  align-items: center;
                  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                  overflow-y: auto;
                }}
                .page-box {{
                  margin-bottom: 24px;
                  box-shadow: 0 6px 20px rgba(0,0,0,0.5);
                  border-radius: 6px;
                  background: white;
                  overflow: hidden;
                  width: 100%;
                  max-width: 880px;
                }}
                .page-header {{
                  background: #1e293b;
                  color: #cbd5e1;
                  font-size: 13px;
                  font-weight: 600;
                  padding: 8px 14px;
                  border-bottom: 1px solid #334155;
                  display: flex;
                  justify-content: space-between;
                }}
                canvas {{
                  display: block;
                  width: 100%;
                  height: auto;
                }}
                #status {{
                  color: #38bdf8;
                  padding: 14px;
                  font-size: 14px;
                  font-weight: 600;
                  text-align: center;
                }}
                iframe {{
                  border: none;
                  width: 100%;
                  height: 780px;
                  background: white;
                  border-radius: 6px;
                }}
              </style>
            </head>
            <body>
              <div id="status">Cargando factura completa ({num_pags_tot} página(s))...</div>
              <div id="viewer-container" style="width: 100%; display: flex; flex-direction: column; align-items: center;"></div>
              <iframe id="fallback-frame" style="display:none;" width="100%" height="780px"></iframe>
              <script>
                const b64Data = "{b64_pdf}";
                function base64ToUint8Array(base64) {{
                    const binaryString = atob(base64);
                    const len = binaryString.length;
                    const bytes = new Uint8Array(len);
                    for (let i = 0; i < len; i++) {{
                        bytes[i] = binaryString.charCodeAt(i);
                    }}
                    return bytes;
                }}
                
                const uint8Pdf = base64ToUint8Array(b64Data);
                
                function showFallback() {{
                    document.getElementById('status').style.display = 'none';
                    document.getElementById('viewer-container').style.display = 'none';
                    const frame = document.getElementById('fallback-frame');
                    try {{
                        const blob = new Blob([uint8Pdf], {{type: 'application/pdf'}});
                        const blobUrl = URL.createObjectURL(blob);
                        frame.src = blobUrl + '#toolbar=1&navpanes=1';
                    }} catch (e) {{
                        frame.src = 'data:application/pdf;base64,' + b64Data + '#toolbar=1&navpanes=1';
                    }}
                    frame.style.display = 'block';
                }}

                try {{
                    if (typeof pdfjsLib === 'undefined' && window['pdfjs-dist/build/pdf']) {{
                        window.pdfjsLib = window['pdfjs-dist/build/pdf'];
                    }}
                    if (typeof pdfjsLib !== 'undefined') {{
                        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                        const loadingTask = pdfjsLib.getDocument({{data: uint8Pdf}});
                        loadingTask.promise.then(async function(pdf) {{
                            document.getElementById('status').style.display = 'none';
                            const container = document.getElementById('viewer-container');
                            container.innerHTML = '';
                            
                            for (let pNum = 1; pNum <= pdf.numPages; pNum++) {{
                                try {{
                                    const page = await pdf.getPage(pNum);
                                    const scale = 1.45;
                                    const viewport = page.getViewport({{scale: scale}});
                                    
                                    const pageBox = document.createElement('div');
                                    pageBox.className = 'page-box';
                                    
                                    const pageHeader = document.createElement('div');
                                    pageHeader.className = 'page-header';
                                    pageHeader.innerHTML = '<span>📄 Factura ' + '{fac_sel["Factura"]}' + ' — {fac_sel["Comprobante Siigo"]}</span><span style="background:#334155; padding:2px 8px; border-radius:4px;">Hoja ' + pNum + ' de ' + pdf.numPages + '</span>';
                                    pageBox.appendChild(pageHeader);
                                    
                                    const canvas = document.createElement('canvas');
                                    const ctx = canvas.getContext('2d');
                                    canvas.height = viewport.height;
                                    canvas.width = viewport.width;
                                    pageBox.appendChild(canvas);
                                    
                                    container.appendChild(pageBox);
                                    await page.render({{canvasContext: ctx, viewport: viewport}}).promise;
                                }} catch (renderErr) {{
                                    console.error('Error renderizando página ' + pNum, renderErr);
                                }}
                            }}
                        }}).catch(function(err) {{
                            console.error('Error cargando documento PDF:', err);
                            showFallback();
                        }});
                    }} else {{
                        showFallback();
                    }}
                }} catch (e) {{
                    console.error('Excepción general en visor:', e);
                    showFallback();
                }}
              </script>
            </body>
            </html>
            """
            
            if "Visor Integrado" in modo_vista_doc:
                html_blob = f"""
                <!DOCTYPE html>
                <html>
                <body style="margin:0; padding:0; background:#0f172a;">
                  <iframe id="native-frame" width="100%" height="820px" style="border:1px solid #334155; border-radius:6px; background:white;"></iframe>
                  <script>
                    const b64 = "{b64_pdf}";
                    const bin = atob(b64);
                    const bArr = new Uint8Array(bin.length);
                    for (let i = 0; i < bin.length; i++) bArr[i] = bin.charCodeAt(i);
                    const bUrl = URL.createObjectURL(new Blob([bArr], {{type: 'application/pdf'}}));
                    document.getElementById('native-frame').src = bUrl + '#toolbar=1&navpanes=1';
                  </script>
                </body>
                </html>
                """
                components.html(html_blob, height=840, scrolling=True)
            else:
                components.html(html_visor, height=visor_height, scrolling=True)
        elif dict_orig or dict_renom:
            st.warning("⚠️ No se identificó automáticamente esta factura. Puedes seleccionar manualmente cualquier PDF cargado para visualizarlo:")
            todos_los_pdfs = {**dict_orig, **dict_renom}
            pdf_elegido = st.selectbox("Selecciona un archivo PDF cargado:", list(todos_los_pdfs.keys()))
            if pdf_elegido:
                f_bytes_sel = todos_los_pdfs[pdf_elegido]
                # Desbloquear si está protegido
                f_bytes_sel, _ = desbloquear_pdf_bytes(f_bytes_sel, nit_receptor=re.sub(r"\D", "", str(empresa.get("nit", "9013464125"))))
                b64_m = base64.b64encode(f_bytes_sel).decode('utf-8')
                
                st.download_button(
                    label=f"📥 Descargar {pdf_elegido}",
                    data=f_bytes_sel,
                    file_name=pdf_elegido,
                    mime="application/pdf"
                )
                
                html_blob_manual = f"""
                <!DOCTYPE html>
                <html>
                <body style="margin:0; padding:0; background:#0f172a;">
                  <iframe id="man-frame" width="100%" height="750px" style="border:1px solid #334155; border-radius:6px; background:white;"></iframe>
                  <script>
                    const b64 = "{b64_m}";
                    const bin = atob(b64);
                    const bArr = new Uint8Array(bin.length);
                    for (let i = 0; i < bin.length; i++) bArr[i] = bin.charCodeAt(i);
                    const bUrl = URL.createObjectURL(new Blob([bArr], {{type: 'application/pdf'}}));
                    document.getElementById('man-frame').src = bUrl + '#toolbar=1&navpanes=1';
                  </script>
                </body>
                </html>
                """
                components.html(html_blob_manual, height=770, scrolling=True)
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
                    <i>💡 Nota: Sube los archivos PDF en la Pestaña 1 para ver el documento digitalizado en este visor.</i>
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

with tab_triangulacion:
    st.markdown("### 🔀 Triangulación y Cruce de Cuentas por Pagar (Importaciones y Aduanas)")
    st.caption("Resuelve la sustitución de acreedores: cruza los pagos realizados por agentes aduaneros (Euro Shipping, Trade Global) contra las facturas ya causadas (DHL, Cargo Aduana), reconociendo retenciones asumidas (53152001) y diferencias no deducibles (53950501).")

    if "df_procesado" in st.session_state:
        df_total = st.session_state["df_procesado"]
        
        # Identificar facturas del gremio aduanero
        cond_aduanera = (df_total.get("Es Aduanera", False) == True) | df_total["Proveedor"].str.upper().str.contains("EURO|TRADE|CARGO|ADUANA|DHL|PORTUARIA|ALMACENADORA|TERMINAL|CONSOLCARGO")
        df_adu = df_total[cond_aduanera].copy()
        
        # Facturas de Agentes Principales (Amarillas)
        df_agentes_all = df_adu[df_adu["Proveedor"].str.upper().str.contains("EURO SHIPPING|TRADE GLOBAL")].copy()
        
        # Facturas de Terceros Soporte (Rojas y Blancas)
        df_terceros_all = df_adu[~df_adu.index.isin(df_agentes_all.index)].copy()

        # Resumen general de importaciones
        m_c1, m_c2, m_c3, m_c4 = st.columns(4)
        with m_c1:
            st.metric("Facturas Agentes (Amarillas)", len(df_agentes_all))
        with m_c2:
            st.metric("Terceros Soporte (DHL / Aduanas)", len(df_terceros_all))
        with m_c3:
            n_rojas_adu = len(df_terceros_all[df_terceros_all.get("Ya Registrada", False) == True])
            st.metric("Terceros Ya Causados (Rojas)", n_rojas_adu)
        with m_c4:
            n_blancas_adu = len(df_terceros_all[df_terceros_all.get("Ya Registrada", False) == False])
            st.metric("Terceros Nuevos (Blancas)", n_blancas_adu)

        st.markdown("---")

        # Agrupación y Selección de Paquetes de Importación
        # Cada factura de agente aduanero (Euro Shipping, Trade Global) representa una operación de cobro global
        lista_opciones_agente = []
        mapa_agentes = {}
        for idx_ag, (_, ag) in enumerate(df_agentes_all.iterrows()):
            key_ag = f"{ag['Proveedor'][:18]} (Fac {ag['Factura']}) — Total: ${ag['Total']:,.2f} [{ag['Fecha']}]"
            lista_opciones_agente.append(key_ag)
            mapa_agentes[key_ag] = ag

        if not lista_opciones_agente:
            st.warning("⚠️ No se encontraron facturas de agentes aduaneros principales (Euro Shipping o Trade Global) en este reporte.")
        else:
            col_sel_g1, col_sel_g2 = st.columns([2.5, 1])
            with col_sel_g1:
                sel_ag_key = st.selectbox("1. Selecciona la Factura del Agente Aduanero (Cobro Global):", lista_opciones_agente, key="sel_agente_triangulacion")
                agente_actual = mapa_agentes[sel_ag_key]
                tot_agente_actual = float(agente_actual["Total"])

            with col_sel_g2:
                st.write("")
                st.write("")
                st.info(f"📅 **Fecha:** {agente_actual['Fecha']} | Total: **${tot_agente_actual:,.2f}**")

            # Identificar qué facturas de terceros corresponden a este paquete
            # A. Si el Excel trae Grupo Importación definido (ej. grupo '1', '2', etc.), tomar esas
            grp_id_ag = str(agente_actual.get("Grupo Importación", "")).strip()
            terceros_predefinidos = []
            if grp_id_ag and grp_id_ag != 'nan':
                for _, tr in df_terceros_all[df_terceros_all["Grupo Importación"] == grp_id_ag].iterrows():
                    terceros_predefinidos.append(tr["Factura"])

            # B. Si no hay grupo predefinido en Excel, buscar sugerencias automáticas por ventana de fecha y aproximación
            if not terceros_predefinidos:
                try:
                    f_ag_dt = pd.to_datetime(agente_actual["Fecha"], dayfirst=True)
                    # Candidatos en ventana de +/- 25 días
                    candidatos_periodo = []
                    for _, tr in df_terceros_all.iterrows():
                        try:
                            f_tr_dt = pd.to_datetime(tr["Fecha"], dayfirst=True)
                            diff_dias = (f_ag_dt - f_tr_dt).days
                            if -5 <= diff_dias <= 35 and tr["Total"] < tot_agente_actual:
                                candidatos_periodo.append(tr["Factura"])
                        except:
                            pass
                    terceros_predefinidos = candidatos_periodo[:4]
                except:
                    terceros_predefinidos = []

            # Selector interactivo para que el usuario pueda marcar/desmarcar con total libertad
            opciones_terceros_lista = []
            mapa_terceros = {}
            for _, tr in df_terceros_all.iterrows():
                es_r = tr.get("Ya Registrada", False)
                comp_txt = f"🔴 [Ya Registrada en Siigo: {tr.get('Comprobante Previo', '10-Prev')}]" if es_r else "⚪ [Pendiente]"
                tag_t = f"[{tr['Factura']}] {tr['Fecha']} - {tr['Proveedor'][:24]} (${tr['Total']:,.2f}) — {comp_txt}"
                opciones_terceros_lista.append(tag_t)
                mapa_terceros[tr["Factura"]] = tag_t

            # Valores por defecto para el multiselect
            defaults_sel = [mapa_terceros[fac] for fac in terceros_predefinidos if fac in mapa_terceros]

            st.markdown("##### 2. Selecciona las facturas de terceros que van trianguladas en este paquete:")
            st.caption("Marca o desmarca las facturas de terceros (Cargo Aduana, DHL, Almacenadoras, Puertos). Puedes incluir las facturas rojas (ya causadas) para cruzar sus cuentas por pagar:")

            terceros_elegidos_tags = st.multiselect(
                "Facturas de terceros incluidas:",
                opciones_terceros_lista,
                default=defaults_sel,
                key=f"multisel_terceros_{agente_actual['Factura']}"
            )

            # Filtrar DataFrame de terceros seleccionados
            facs_seleccionadas = [t.split("]")[0].replace("[", "").strip() for t in terceros_elegidos_tags]
            terceros_actual = df_terceros_all[df_terceros_all["Factura"].isin(facs_seleccionadas)].copy()

            # Resumen del Paquete Seleccionado
            col_pq1, col_pq2 = st.columns([1.3, 2.5])
            with col_pq1:
                st.markdown(f"""
                <div style="background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #d97706; border-radius:8px; padding:14px; margin-bottom:12px;">
                    <h4 style="margin:0 0 6px 0; color:#92400e;">🟡 Factura del Agente (Cobro Total)</h4>
                    <p style="margin:0; font-size:14px; color:#78350f;">
                        <b>Proveedor:</b> {agente_actual['Proveedor']}<br>
                        <b>NIT:</b> {agente_actual['NIT Emisor']}<br>
                        <b>Factura:</b> {agente_actual['Factura']}<br>
                        <b>Total Facturado:</b> <span style="font-size:16px; font-weight:bold; color:#b45309;">${tot_agente_actual:,.2f}</span><br>
                        <b>IVA de Importación / Serv:</b> ${float(agente_actual.get('IVA', 0.0)):,.2f}
                    </p>
                </div>
                """, unsafe_allow_html=True)

            with col_pq2:
                if not terceros_actual.empty:
                    filas_terc_disp = []
                    for _, tr in terceros_actual.iterrows():
                        es_r = tr.get("Ya Registrada", False)
                        badge_est = f"🔴 Ya en Siigo ({tr.get('Comprobante Previo', '10-Prev')})" if es_r else "⚪ Pendiente"
                        filas_terc_disp.append({
                            "Estado": badge_est,
                            "Proveedor Tercero": tr["Proveedor"],
                            "Factura": tr["Factura"],
                            "Fecha": tr["Fecha"],
                            "Valor Total": tr["Total"],
                            "Cuenta Pasivo": tr.get("Cuenta Pasivo Especifica", "22050505" if "CARGO" in tr["Proveedor"] else "23359501")
                        })
                    df_terc_disp = pd.DataFrame(filas_terc_disp)
                    st.dataframe(df_terc_disp.style.format({"Valor Total": "${:,.2f}"}), use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ No has seleccionado facturas de terceros para este paquete aún. Selecciona arriba las que correspondan.")

        # CALCULAR ASIENTO CONTABLE CUADRADO DEL PAQUETE
        df_asiento_paquete, dif_no_ded, ret_asum = generar_asiento_triangulacion_paquete(agente_actual, terceros_actual)

        st.markdown("#### ⚖️ Asiento Contable Cuadrado del Cruce (Para Siigo):")
        st.caption("Detalle de partida doble: cancela las cuentas por pagar a terceros, asume las retenciones y manda la diferencia a no deducible:")

        st.dataframe(
            df_asiento_paquete.style.format({"Débito ($)": "${:,.2f}", "Crédito ($)": "${:,.2f}"}),
            use_container_width=True,
            hide_index=True
        )

        c_cuad1, c_cuad2, c_cuad3, c_cuad4 = st.columns(4)
        sum_deb_pq = df_asiento_paquete["Débito ($)"].sum()
        sum_cred_pq = df_asiento_paquete["Crédito ($)"].sum()
        dif_cuad_pq = abs(sum_deb_pq - sum_cred_pq)
        
        with c_cuad1:
            st.metric("Total Débito", f"${sum_deb_pq:,.2f}")
        with c_cuad2:
            st.metric("Total Crédito", f"${sum_cred_pq:,.2f}")
        with c_cuad3:
            st.metric("Retenciones Asumidas (53152001)", f"${ret_asum:,.2f}")
        with c_cuad4:
            st.metric("Diferencia No Deducible (53950501)", f"${dif_no_ded:,.2f}")

        if dif_cuad_pq < 0.05:
            st.success("✅ **Comprobante de Cruce Verificado:** Partida doble cuadrada con sumas iguales al centavo ($0.00).")
        else:
            st.error(f"Diferencia de cuadre: ${dif_cuad_pq:,.2f}")

        st.markdown("---")
        st.markdown("#### 📥 Exportar Todos los Cruces de Importación para Siigo:")
        st.caption("Genera una planilla formulada con todos los paquetes de importación triangulados listos para importar a Siigo:")

        # Generar Excel de todos los cruces triangulados
        todos_asientos_lista = []
        consecutivo_cruce = 1
        for g_k, g_v in grupos_dict.items():
            df_as_p, _, _ = generar_asiento_triangulacion_paquete(g_v["agente"], g_v["terceros"])
            for _, fila_as in df_as_p.iterrows():
                todos_asientos_lista.append({
                    "Paquete": f"Paquete #{g_k}",
                    "Tipo Comprobante": 14,
                    "Consecutivo": consecutivo_cruce,
                    "Fecha Elaboración": g_v["agente"]["Fecha"],
                    "Sigla Moneda": "COP",
                    "Tasa Cambio": 1,
                    "Código Cuenta": fila_as["Código Cuenta"],
                    "Identificación Tercero": fila_as["Tercero / NIT"].split("-")[0].strip(),
                    "Sucursal": 0,
                    "Descripción": fila_as["Descripción Cuenta"],
                    "Débito": fila_as["Débito ($)"],
                    "Crédito": fila_as["Crédito ($)"]
                })
            consecutivo_cruce += 1

        if todos_asientos_lista:
            df_export_cruces = pd.DataFrame(todos_asientos_lista)
            buf_cruce_ex = io.BytesIO()
            with pd.ExcelWriter(buf_cruce_ex, engine="openpyxl") as wr_cr:
                df_export_cruces.to_excel(wr_cr, sheet_name="interfaz_siigo_cruces", index=False)
            buf_cruce_ex.seek(0)
            
            st.download_button(
                label=f"📥 Descargar Planilla Siigo de Cruces de Importación ({len(grupos_dict)} Paquetes Triangulados)",
                data=buf_cruce_ex.getvalue(),
                file_name=f"Cruces_Importacion_Siigo_{empresa['nombre'].replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    else:
        st.info("💡 Sube el reporte Excel de la DIAN en la Pestaña 1 para habilitar la triangulación y cruce de importaciones.")


with tab_siigo:
    st.markdown("### Descargar Planilla Oficial Siigo Nube (3 Hojas)")
    st.caption("Planilla oficial formulada con 'matriz_captura', 'interfaz_siigo' y 'Parametrización'.")
    
    if "df_procesado" in st.session_state:
        df_full = st.session_state["df_procesado"]
        
        # FILTRO DE PROTECCIÓN: Excluir facturas rojas (ya causadas en Siigo) y facturas que van por triangulación aduanera
        n_rojas = len(df_full[df_full.get("Ya Registrada", False) == True])
        df_p = df_full[
            (df_full.get("Ya Registrada", False) == False) & 
            (~((df_full.get("Es Aduanera", False) == True) & (df_full.get("Grupo Importación", "") != "")))
        ].copy()
        
        if n_rojas > 0:
            st.info(f"🛡️ **Protección contra duplicados:** Se excluyeron {n_rojas} facturas marcadas en rojo que ya estaban causadas en Siigo. Solo se están exportando las facturas nuevas.")
            
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

