import streamlit as st
import pandas as pd
import openpyxl
import io
import time
from datetime import datetime
import zipfile
import re
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Sistema ERP y Auditoría Contable DIAN", layout="wide", page_icon="🏢")

# 1. TEMPORIZADOR DE INACTIVIDAD (30 MINUTOS = 1800 SEG)
TIEMPO_MAX_INACTIVIDAD = 30 * 60

st.markdown("""
<script>
let tiempoLimite = 30 * 60 * 1000;
let temporizador;
function resetInactividad() {
    clearTimeout(temporizador);
    temporizador = setTimeout(() => {
        alert("Tu sesión ha expirado por 30 minutos de inactividad.");
        window.location.search = "?sesion_expirada=1";
    }, tiempoLimite);
}
window.onload = resetInactividad;
document.onmousemove = resetInactividad;
document.onkeydown = resetInactividad;
document.onclick = resetInactividad;
document.onscroll = resetInactividad;
</script>
""", unsafe_allow_html=True)

# EMPRESAS REGISTRADAS
EMPRESAS_DISPONIBLES = [
    {
        "nombre": "INDUMAQ ER SAS",
        "nit": "901.346.412-5",
        "actividad": "Comercio y Reparación de Maquinaria / Importaciones",
        "regimen": "Responsable de IVA",
        "estado": "ACTIVA"
    },
    {
        "nombre": "ASMINCOL S.A.S.",
        "nit": "900.467.519-1",
        "actividad": "Servicios Mineros y Construcción",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    },
    {
        "nombre": "CONSTRUDISENO CT SAS",
        "nit": "900.524.356-8",
        "actividad": "Construcción y Obras Civiles",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    },
    {
        "nombre": "SCG TRANSPORTES",
        "nit": "901.700.731-8",
        "actividad": "Transporte de Carga y Logística",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    },
    {
        "nombre": "OPJ SAS",
        "nit": "901.425.101-3",
        "actividad": "Servicios Generales y Operaciones",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    }
]

# Inicialización de variables de sesión
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "empresa_activa" not in st.session_state:
    st.session_state["empresa_activa"] = None
if "proceso_activo" not in st.session_state:
    st.session_state["proceso_activo"] = None
if "ultima_actividad" not in st.session_state:
    st.session_state["ultima_actividad"] = time.time()

# Alerta si expiró por inactividad
if st.query_params.get("sesion_expirada") == "1":
    st.session_state["autenticado"] = False
    st.session_state["empresa_activa"] = None
    st.session_state["proceso_activo"] = None
    st.query_params.clear()
    st.warning("⏳ Tu sesión se cerró automáticamente por 30 minutos de inactividad.")

# Control de inactividad y persistencia
ahora = time.time()
if st.session_state["autenticado"]:
    if ahora - st.session_state["ultima_actividad"] > TIEMPO_MAX_INACTIVIDAD:
        st.session_state["autenticado"] = False
        st.session_state["empresa_activa"] = None
        st.session_state["proceso_activo"] = None
        st.query_params.clear()
        st.warning("⏳ Sesión cerrada por 30 minutos de inactividad.")
        st.stop()
    else:
        st.session_state["ultima_actividad"] = ahora
else:
    if st.query_params.get("auth") == "1":
        t_param = int(st.query_params.get("t", 0))
        if ahora - t_param < TIEMPO_MAX_INACTIVIDAD:
            st.session_state["autenticado"] = True
            st.session_state["ultima_actividad"] = ahora
            emp_param = st.query_params.get("emp", "")
            if emp_param:
                for e in EMPRESAS_DISPONIBLES:
                    if emp_param in e["nombre"]:
                        st.session_state["empresa_activa"] = e
                        break
            proc_param = st.query_params.get("proc", "")
            if proc_param:
                st.session_state["proceso_activo"] = proc_param

# PANTALLA 1: LOGIN (Línea con spec explícito para evitar TypeError)
if not st.session_state["autenticado"]:
    col1, col2, col3 = st.columns()
    with col2:
        st.markdown("<h2 style='text-align: center;'>Portal ERP y Auditoría Contable</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Plataforma unificada para gestión contable y DIAN (Siigo / World Office)</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario / Correo Electrónico", value="deibydaza2014@gmail.com")
            password = st.text_input("Contraseña", type="password", value="123456")
            submit = st.form_submit_button("Ingresar al Ecosistema", use_container_width=True)
            
            if submit:
                if usuario and password:
                    st.session_state["autenticado"] = True
                    st.session_state["ultima_actividad"] = time.time()
                    st.query_params["auth"] = "1"
                    st.query_params["t"] = str(int(time.time()))
                    st.rerun()
                else:
                    st.error("Ingresa usuario y contraseña.")
    st.stop()

# Barra superior con botón para Cerrar Sesión manual
col_top_l, col_top_r = st.columns()
with col_top_r:
    if st.button("🚪 Cerrar Sesión"):
        st.session_state["autenticado"] = False
        st.session_state["empresa_activa"] = None
        st.session_state["proceso_activo"] = None
        st.query_params.clear()
        st.rerun()

# PANTALLA 2: SELECTOR DE EMPRESAS
if not st.session_state["empresa_activa"]:
    st.title("Selección de Empresa")
    st.caption("Selecciona la entidad sobre la cual vas a trabajar:")
    
    c1, c2 = st.columns(2)
    for i, emp in enumerate(EMPRESAS_DISPONIBLES):
        col = c1 if i % 2 == 0 else c2
        with col:
            with st.container(border=True):
                st.subheader(f"🏢 {emp['nombre']}")
                st.caption(f"**NIT:** {emp['nit']} | **Régimen:** {emp['regimen']}")
                st.write(emp['actividad'])
                if emp["estado"] == "ACTIVA":
                    st.success("Estado: ACTIVA")
                    if st.button(f"Ingresar a {emp['nombre']}", key=f"btn_emp_{i}", use_container_width=True):
                        st.session_state["empresa_activa"] = emp
                        st.session_state["proceso_activo"] = None
                        st.session_state["ultima_actividad"] = time.time()
                        st.query_params["emp"] = emp["nombre"]
                        st.query_params["t"] = str(int(time.time()))
                        st.rerun()
                else:
                    st.info(f"Estado: {emp['estado']}")
                    st.button(f"Pendiente Parametrización ({emp['estado']})", key=f"btn_emp_{i}", disabled=True, use_container_width=True)
    st.stop()

empresa = st.session_state["empresa_activa"]

# PANTALLA 3: MENU DE PROCESOS OPERATIVOS
PROCESOS_SISTEMA = [
    {
        "id": "facturacion",
        "titulo": "Facturas de Compra, Venta y Devoluciones",
        "desc": "Carga de reportes DIAN, separación y renombrado de facturas compiladas por comprobante, auditoría contable y Siigo.",
        "estado": "ACTIVO"
    },
    {
        "id": "nomina",
        "titulo": "Gestión Laboral y Nómina Electrónica",
        "desc": "Cálculo de liquidación de nómina, provisiones de prestaciones sociales y soportes electrónicos DIAN.",
        "estado": "PROXIMAMENTE"
    },
    {
        "id": "conciliacion",
        "titulo": "Tesorería y Conciliación Bancaria",
        "desc": "Cruce automatizado de extractos bancarios contra libros auxiliares y control de partidas conciliatorias.",
        "estado": "PROXIMAMENTE"
    },
    {
        "id": "notas",
        "titulo": "Notas de Contabilidad y Cierre Fiscal",
        "desc": "Comprobantes de ajuste, amortizaciones de intangibles (NIC 38) y depreciaciones.",
        "estado": "PROXIMAMENTE"
    }
]

if not st.session_state["proceso_activo"]:
    col_t1, col_t2 = st.columns()
    with col_t1:
        st.subheader(f"🏢 {empresa['nombre']} — Panel de Procesos")
        st.caption(f"NIT: {empresa['nit']} | Selecciona el módulo de trabajo:")
    with col_t2:
        if st.button("Cambiar Empresa"):
            st.session_state["empresa_activa"] = None
            st.session_state["proceso_activo"] = None
            if "emp" in st.query_params: del st.query_params["emp"]
            if "proc" in st.query_params: del st.query_params["proc"]
            st.rerun()

    st.markdown("---")
    
    cp1, cp2 = st.columns(2)
    for idx, proc in enumerate(PROCESOS_SISTEMA):
        col_p = cp1 if idx % 2 == 0 else cp2
        with col_p:
            with st.container(border=True):
                st.subheader(proc['titulo'])
                st.write(proc['desc'])
                if proc["estado"] == "ACTIVO":
                    st.success("Módulo: ACTIVO")
                    if st.button("Abrir Módulo de Facturación", key=f"btn_proc_{idx}", use_container_width=True):
                        st.session_state["proceso_activo"] = proc["id"]
                        st.session_state["ultima_actividad"] = time.time()
                        st.query_params["proc"] = proc["id"]
                        st.query_params["t"] = str(int(time.time()))
                        st.rerun()
                else:
                    st.info(f"Módulo: {proc['estado']}")
                    st.button("En Construcción", key=f"btn_proc_{idx}", disabled=True, use_container_width=True)
    st.stop()

# PANTALLA 4: FACTURACION, AUDITORIA Y SIIGO
col_nav1, col_nav2 = st.columns()
with col_nav1:
    st.subheader(f"🏢 {empresa['nombre']} — Facturación y Auditoría")
    st.caption(f"NIT: {empresa['nit']} | Módulo: Facturas de Compra, Venta y Devoluciones")
with col_nav2:
    if st.button("Volver a Procesos"):
        st.session_state["proceso_activo"] = None
        if "proc" in st.query_params: del st.query_params["proc"]
        st.rerun()

st.markdown("---")

tab_compras, tab_auditoria, tab_siigo = st.tabs([
    "1. Cargar Excel, Desglosar y Renombrar PDFs",
    "2. Auditoría y Trazabilidad Fiscal",
    "3. Exportar Planilla Oficial a Siigo (3 Hojas)"
])

AGENTES_ADUANEROS = ["DHL", "ADUANA", "EURO SHIPPING", "PORTUARIA", "ALMACENADORA", "CARGO", "TRADE GLOBAL", "TERMINAL", "BUENAVENTURA"]

def safe_read_pdf(file_obj):
    """Lee el buffer completo del archivo de Streamlit sin problemas de puntero vacío"""
    try:
        data = file_obj.getvalue()
    except Exception:
        file_obj.seek(0)
        data = file_obj.read()
    if not data:
        file_obj.seek(0)
        data = file_obj.read()
    if not data:
        raise ValueError("El archivo se leyó vacío (0 bytes).")
    return PdfReader(io.BytesIO(data))

def parse_fecha(fecha_val):
    """Interpreta fechas en formatos colombianos DD/MM/AAAA o AAAA-MM-DD"""
    if pd.isna(fecha_val):
        return None
    try:
        if isinstance(fecha_val, datetime):
            return fecha_val
        s = str(fecha_val).split()[0].replace("-", "/").strip()
        partes = s.split("/")
        if len(partes) == 3:
            if len(partes[0]) == 4:
                return datetime(int(partes[0]), int(partes), int(partes))
            elif len(partes) == 4:
                return datetime(int(partes), int(partes), int(partes[0]))
        return pd.to_datetime(fecha_val, dayfirst=True).to_pydatetime()
    except:
        return None

def get_parametros_tributarios_fecha(dt_obj, uvt_val=52374):
    """
    Aplica las variaciones históricas del año 2026:
    1. 1 Ene - 7 May 2026: Decreto 0572 de 2025 (Compras 10 UVT, Servicios 2 UVT)
    2. 8 May - 30 Jun 2026: Suspensión Consejo de Estado / Comunicado DIAN 070 (Decreto 1625: Compras 27 UVT, Servicios 4 UVT)
    3. 1 Jul 2026 en adelante: Reactivación Decreto 0572 (Compras 10 UVT, Servicios 2 UVT)
    """
    if dt_obj and datetime(2026, 5, 8) <= dt_obj <= datetime(2026, 6, 30):
        uvt_compras = 27
        uvt_servicios = 4
        norma_desc = "Decreto 1625/2016 (Suspensión Consejo de Estado 8 May - 30 Jun)"
        badge_norma = "Suspensión C.E. (27 UVT)"
    else:
        uvt_compras = 10
        uvt_servicios = 2
        if dt_obj and dt_obj >= datetime(2026, 7, 1):
            norma_desc = "Decreto 0572/2025 Reactivado (1 Jul en adelante)"
            badge_norma = "D.0572 Reactivado (10 UVT)"
        else:
            norma_desc = "Decreto 0572/2025 Vigente (1 Ene - 7 May)"
            badge_norma = "Decreto 0572 (10 UVT)"
            
    tope_compras = uvt_compras * uvt_val
    tope_servicios = uvt_servicios * uvt_val
    return uvt_compras, tope_compras, uvt_servicios, tope_servicios, norma_desc, badge_norma

def clasificar_factura_completa(nit_emisor, nombre_emisor, valor_base, tipo_doc, dt_obj, uvt_val=52374):
    """Clasifica contablemente según PUC, fecha de emisión y régimen tributario vigente en esa fecha"""
    nombre = str(nombre_emisor).upper()
    es_nc = "CREDITO" in str(tipo_doc).upper() or "CRÉDITO" in str(tipo_doc).upper()
    op = "Devolucion Compra" if es_nc else "Compra"
    
    u_comp, tope_comp, u_serv, tope_serv, norma_desc, badge_norma = get_parametros_tributarios_fecha(dt_obj, uvt_val)
    
    # 1. Agenciamiento Aduanero, Fletes e Importaciones (Cuenta 146505)
    if any(k in nombre for k in AGENTES_ADUANEROS):
        rfte = round(valor_base * 0.04, 2) if valor_base >= tope_servicios else 0.0
        return op, "146505", "22050501", f"Importación / Tránsito - {nombre_emisor[:25]}", rfte, "Importación (1465)", f"Agenciamiento / Tránsito ({norma_desc}, Base: {u_serv} UVT = ${tope_servicios:,.0f})", badge_norma
        
    # 2. Repuestos y Mantenimiento de Maquinaria
    repuestos_kw = ["FERROMENDEZ", "TORNILLOLOCO", "CAUCHOS", "ASIMFER", "MAFLEXCOL", "EMPRECOL", "BATTS ZONE", "MECANIZAR", "HIDRAHULICAS", "BAMACOLGROUP"]
    if any(k in nombre for k in repuestos_kw):
        if valor_base >= tope_comp:
            rfte = round(valor_base * 0.025, 2)
            return op, "14350101", "22050501", f"Repuestos / Inventario - {nombre_emisor[:25]}", rfte, "Inventario", f"Repuestos >= {u_comp} UVT ({norma_desc})", badge_norma
        else:
            return op, "61800101", "23359501", f"Mantenimiento menor - {nombre_emisor[:25]}", 0.0, "Costo Mantenimiento", f"Repuestos menores a {u_comp} UVT (${tope_comp:,.0f}) sin retención", badge_norma
            
    # 3. Hoteles y Hospedajes
    if any(k in nombre for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK"]):
        rfte = round(valor_base * 0.035, 2) if valor_base >= tope_servicios else 0.0
        return op, "51550501", "23359501", f"Alojamiento / Viaje - {nombre_emisor[:25]}", rfte, "Gasto Viaje", f"Hospedaje de personal (ReteFuente 3.5%, base {u_serv} UVT)", badge_norma
        
    # 4. Suscripciones y Software (Autorretenedores)
    if "SIIGO" in nombre:
        return op, "51352001", "23359501", f"Software Siigo - {nombre_emisor[:25]}", 0.0, "Software", "Autorretenedor de renta (Sin ReteFuente)", badge_norma
        
    # 5. Papelería y Útiles
    if "PANAMERICANA" in nombre:
        return op, "51953001", "23359501", f"Papelería - {nombre_emisor[:25]}", 0.0, "Gastos Papelería", "Útiles de oficina sin retención", badge_norma
        
    # 6. Compras Generales
    if valor_base >= tope_comp:
        rfte = round(valor_base * 0.025, 2)
        return op, "14350101", "22050501", f"Compra mercancías - {nombre_emisor[:25]}", rfte, "Mercancía", f"Compra general >= {u_comp} UVT (${tope_comp:,.0f})", badge_norma
    else:
        return op, "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}", 0.0, "Gasto General", f"Compra menor general < {u_comp} UVT", badge_norma

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF (Compilado o Individuales)")
    
    # PARAMETRIZACION DE CONSECUTIVOS Y REGIMEN TEMPORAL TRIBUTARIO
    with st.expander("⚙️ **Configuración de Comprobantes, Consecutivos y Marco Tributario 2026**", expanded=True):
        st.markdown("##### 1. Numeración de Comprobantes:")
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            st.markdown("**1er Semestre (1 Ene - 30 Jun)**")
            t_comp_sem1 = st.number_input("Tipo Comprobante (Ene-Jun):", min_value=1, max_value=99, value=16, step=1)
            consec_ini_sem1 = st.number_input("Consecutivo Inicial (Ene-Jun):", min_value=1, value=1, step=1)
        with col_c2:
            st.markdown("**2do Semestre (1 Jul en adelante)**")
            t_comp_sem2 = st.number_input("Tipo Comprobante (Jul en adelante):", min_value=1, max_value=99, value=10, step=1)
            consec_ini_sem2 = st.number_input("Consecutivo Inicial (Jul en adelante):", min_value=1, value=791, step=1)
        with col_c3:
            st.markdown("**Notas Crédito / Devoluciones**")
            t_comp_nc = st.number_input("Tipo Comprobante Devolución:", min_value=1, max_value=99, value=17, step=1)
            consec_ini_nc = st.number_input("Consecutivo Inicial NC:", min_value=1, value=1, step=1)

        st.markdown("---")
        st.markdown("##### 2. Marco Tributario 2026 Integrado (UVT: $52.374):")
        st.info("""
        **Reglas temporales aplicadas automáticamente según la fecha de cada factura:**
        * **1 Ene - 7 May 2026:** Decreto 0572/2025 ➔ Compras base **10 UVT ($523.740)** | Servicios base **2 UVT ($104.748)**.
        * **8 May - 30 Jun 2026:** Suspensión Consejo de Estado / Comunicado DIAN 070 ➔ Regreso a Decreto 1625: Compras base **27 UVT ($1.414.098)** | Servicios base **4 UVT ($209.496)**.
        * **1 Jul 2026 en adelante:** Reactivación Decreto 0572 ➔ Compras base **10 UVT ($523.740)** | Servicios base **2 UVT ($104.748)**.
        """)

    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o matriz de compras, y el archivo **PDF compilado** para desglose y renombrado automático por comprobante.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel (DIAN prueba.xlsx o Matriz)", type=["xlsx", "xls"])
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF (Sube el PDF compilado o sueltas)", type=["pdf"], accept_multiple_files=True)
        
    if archivo_excel is not None:
        try:
            df_dian = pd.read_excel(archivo_excel)
        except Exception as e:
            df_dian = pd.DataFrame()
            st.error(f"Error al leer el archivo Excel: {e}")
            
        if not df_dian.empty:
            st.success(f"Reporte cargado con éxito: **{len(df_dian)} registros identificados**.")
            
            c_sem1 = int(consec_ini_sem1)
            c_sem2 = int(consec_ini_sem2)
            c_nc = int(consec_ini_nc)
            
            filas = []
            for idx, r in df_dian.iterrows():
                tipo_doc = r.get("Tipo de documento", r.get("Operacion", "Factura electrónica"))
                folio = str(r.get("Folio", r.get("Factura Num", f"{idx+1}"))).strip()
                prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
                if prefijo.lower() == "nan": prefijo = ""
                fecha_str = str(r.get("Fecha Emisión", r.get("Fecha (DD/MM/AAAA)", r.get("Fecha", "01/01/2026")))).split()[0]
                dt_obj = parse_fecha(fecha_str)
                fecha_formateada = dt_obj.strftime("%d/%m/%Y") if dt_obj else fecha_str
                
                nit_e = str(r.get("NIT Emisor", r.get("NIT", ""))).strip().split("-")[0].replace(".", "")
                nom_e = str(r.get("Nombre Emisor", r.get("Descripcion", "Proveedor"))).strip()
                
                tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
                iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
                base_col = float(r.get("Valor Base", 0.0)) if pd.notna(r.get("Valor Base")) else 0.0
                base = base_col if base_col > 0 else round(tot - iva, 2)
                if tot == 0.0 and base > 0:
                    tot = base + iva
                
                op, cta_p, cta_c, desc, rfte, cat, razon, badge_norma = clasificar_factura_completa(
                    nit_e, nom_e, base, tipo_doc, dt_obj
                )
                
                es_devolucion = op == "Devolucion Compra"
                if es_devolucion:
                    t_comp_final = int(t_comp_nc)
                    consecutivo_final = c_nc
                    c_nc += 1
                else:
                    if dt_obj and dt_obj <= datetime(2026, 6, 30):
                        t_comp_final = int(t_comp_sem1)
                        consecutivo_final = c_sem1
                        c_sem1 += 1
                    else:
                        t_comp_final = int(t_comp_sem2)
                        consecutivo_final = c_sem2
                        c_sem2 += 1
                
                nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
                nombre_pdf_esperado = f"Comp_{t_comp_final}-{consecutivo_final}_{prefijo}{folio}_{nom_limpio_prov}.pdf"
                
                filas.append({
                    "N°": idx + 1,
                    "Tipo Comp": t_comp_final,
                    "Consecutivo": consecutivo_final,
                    "Comprobante Siigo": f"Comp {t_comp_final}-{consecutivo_final}",
                    "Fecha (DD/MM/AAAA)": fecha_formateada,
                    "Factura": f"{prefijo}-{folio}" if prefijo else folio,
                    "Folio": folio,
                    "Prefijo": prefijo,
                    "Proveedor": nom_e,
                    "NIT": nit_e,
                    "NIT Emisor": nit_e,
                    "Descripcion": desc,
                    "Operacion": op,
                    "Cta Principal": cta_p,
                    "Categoría": cat,
                    "Valor Base": base,
                    "Base": base,
                    "IVA": iva,
                    "ReteFuente": rfte,
                    "ReteICA": 0.0,
                    "ReteIVA": 0.0,
                    "Total": tot,
                    "Cta Contrapartida": cta_c,
                    "Razón Contable": razon,
                    "Norma Aplicada": badge_norma,
                    "Soporte PDF Renombrado": nombre_pdf_esperado
                })
                
            df_proc = pd.DataFrame(filas)
            st.session_state["df_procesado"] = df_proc
            
            st.markdown("#### Matriz Contable Generada según Fechas, Consecutivos y Marco Tributario:")
            st.dataframe(df_proc[["Comprobante Siigo", "Fecha (DD/MM/AAAA)", "Factura", "Proveedor", "Cta Principal", "Valor Base", "IVA", "ReteFuente", "Total", "Norma Aplicada", "Soporte PDF Renombrado"]], use_container_width=True)

    # Procesar, Separar y Renombrar PDFs
    if archivos_pdfs:
        st.markdown("---")
        st.markdown("#### ✂️ Motor de Desglose, Separación y Renombrado de PDFs")
        
        modo_corte = st.radio(
            "Selecciona el método de corte y separación:",
            [
                "🔍 Modo Inteligente (Escanea y busca el número de factura/prefijo en cada hoja del PDF compilado)",
                "📑 Modo Secuencial (1 Factura por hoja en el mismo orden del Excel)",
                "📁 Modo Archivos Sueltos (Ya tienes múltiples PDFs individuales y deseas renombrarlos por comprobante)"
            ]
        )
        
        if st.button("🚀 Ejecutar Separación y Guardado por Comprobante"):
            if "df_procesado" not in st.session_state:
                st.warning("Primero debes subir el archivo Excel para saber los números de comprobante.")
            else:
                df_proc = st.session_state["df_procesado"]
                buffer_zip = io.BytesIO()
                nit_limpio = "901346412"
                contrasenas = [nit_limpio, f"{nit_limpio}5", f"{nit_limpio}-5", ""]
                
                total_paginas = []
                for pdf_file in archivos_pdfs:
                    try:
                        reader = safe_read_pdf(pdf_file)
                        if reader.is_encrypted:
                            for pwd in contrasenas:
                                try:
                                    if reader.decrypt(pwd) > 0: break
                                except: pass
                        for p in reader.pages:
                            txt = p.extract_text() or ""
                            total_paginas.append((p, txt))
                    except Exception as e:
                        st.error(f"Error procesando {pdf_file.name}: {e}")
                        
                if not total_paginas:
                    st.error("No se pudieron extraer páginas de los archivos PDF subidos.")
                else:
                    st.info(f"Se cargaron un total de **{len(total_paginas)} páginas de facturas** para procesar.")
                    facturas_asignadas = {}
                    
                    if "Modo Secuencial" in modo_corte:
                        for idx_f, row in df_proc.iterrows():
                            if idx_f < len(total_paginas):
                                facturas_asignadas[idx_f] = [total_paginas[idx_f][0]]
                                
                    elif "Modo Archivos Sueltos" in modo_corte:
                        for idx_pdf, pdf_file in enumerate(archivos_pdfs):
                            try:
                                r_single = safe_read_pdf(pdf_file)
                                if r_single.is_encrypted:
                                    for pwd in contrasenas:
                                        try:
                                            if r_single.decrypt(pwd) > 0: break
                                        except: pass
                                txt_single = "".join([p.extract_text() or "" for p in r_single.pages])
                                txt_clean = re.sub(r'[^a-zA-Z0-9]', '', txt_single.upper())
                                
                                matched_idx = idx_pdf if idx_pdf < len(df_proc) else None
                                for idx_f, row in df_proc.iterrows():
                                    fac_num = re.sub(r'[^a-zA-Z0-9]', '', f"{row['Prefijo']}{row['Folio']}".upper())
                                    if fac_num and len(fac_num) >= 4 and fac_num in txt_clean:
                                        matched_idx = idx_f
                                        break
                                if matched_idx is not None:
                                    facturas_asignadas[matched_idx] = list(r_single.pages)
                            except Exception as e:
                                st.error(f"Error procesando {pdf_file.name}: {e}")
                    else:
                        inv_actual = 0
                        for p_idx, (p_obj, p_txt) in enumerate(total_paginas):
                            clean_txt = re.sub(r'[^a-zA-Z0-9]', '', p_txt.upper())
                            match_encontrado = None
                            for idx_f, row in df_proc.iterrows():
                                num_compuesto = re.sub(r'[^a-zA-Z0-9]', '', f"{row['Prefijo']}{row['Folio']}".upper())
                                num_solo = re.sub(r'[^a-zA-Z0-9]', '', str(row['Folio']).upper())
                                nit_prov = str(row['NIT']).replace(".", "")
                                
                                if num_compuesto and len(num_compuesto) >= 4 and num_compuesto in clean_txt:
                                    match_encontrado = idx_f
                                    break
                                elif num_solo and len(num_solo) >= 4 and num_solo in clean_txt:
                                    if nit_prov and nit_prov in clean_txt:
                                        match_encontrado = idx_f
                                        break
                                    elif len(num_solo) >= 6:
                                        match_encontrado = idx_f
                                        break
                                        
                            if match_encontrado is not None:
                                inv_actual = match_encontrado
                                if inv_actual not in facturas_asignadas:
                                    facturas_asignadas[inv_actual] = []
                                facturas_asignadas[inv_actual].append(p_obj)
                            else:
                                if inv_actual not in facturas_asignadas:
                                    facturas_asignadas[inv_actual] = []
                                facturas_asignadas[inv_actual].append(p_obj)
                                
                    # Empacar en ZIP con el nombre exacto del comprobante contable
                    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        for inv_idx, paginas in facturas_asignadas.items():
                            if inv_idx < len(df_proc):
                                info_row = df_proc.iloc[inv_idx]
                                nombre_archivo = info_row["Soporte PDF Renombrado"]
                            else:
                                nombre_archivo = f"Comprobante_Adicional_{inv_idx+1}.pdf"
                                
                            w_out = PdfWriter()
                            for p in paginas:
                                w_out.add_page(p)
                                
                            pdf_stream = io.BytesIO()
                            w_out.write(pdf_stream)
                            zf.writestr(nombre_archivo, pdf_stream.getvalue())
                            
                    st.success(f"🎉 ¡Proceso completado! Se generaron **{len(facturas_asignadas)} archivos de factura individuales** nombrados con su comprobante contable.")
                    
                    buffer_zip.seek(0)
                    st.download_button(
                        label="📥 Descargar Paquete Completo de Facturas Separadas y Renombradas (.ZIP)",
                        data=buffer_zip,
                        file_name="Facturas_INDUMAQ_Separadas_Por_Comprobante.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

with tab_auditoria:
    st.markdown("### Módulo de Auditoría Contable y Trazabilidad")
    st.caption("Inspección de cuentas, deducciones y separación de gastos por cuenta de terceros.")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        opciones_fac = [f"[{r['Comprobante Siigo']}] {r['Factura']} - {r['Proveedor']} (${r['Total']:,.0f})" for _, r in df_p.iterrows()]
        seleccion = st.selectbox("Selecciona una factura para auditar:", opciones_fac)
        
        comp_sel = seleccion.split("]")[0].replace("[", "")
        fac_sel = df_p[df_p["Comprobante Siigo"] == comp_sel].iloc[0]
        es_aduanero = any(k in fac_sel["Proveedor"].upper() for k in AGENTES_ADUANEROS)
        
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            with st.container(border=True):
                st.subheader(f"Comprobante: {fac_sel['Comprobante Siigo']}")
                st.write(f"**Soporte PDF Vinculado:** `{fac_sel['Soporte PDF Renombrado']}`")
                st.write(f"**Factura:** {fac_sel['Factura']} — **Proveedor:** {fac_sel['Proveedor']} (NIT: {fac_sel['NIT']})")
                st.write(f"**Fecha de Emisión:** {fac_sel['Fecha (DD/MM/AAAA)']} | **Total:** ${fac_sel['Total']:,.2f}")
                st.write(f"**Marco Tributario Aplicado:** `{fac_sel['Norma Aplicada']}`")
                st.markdown("---")
                st.info(f"**Cuenta Asignada:** {fac_sel['Cta Principal']} ({fac_sel['Categoría']})\n\n**Motivo Técnico:** {fac_sel['Razón Contable']}")
                st.write(f"• **Base Gravable:** ${fac_sel['Valor Base']:,.2f}")
                st.write(f"• **IVA Liquidado:** ${fac_sel['IVA']:,.2f}")
                st.write(f"• **Retención en la Fuente:** ${fac_sel['ReteFuente']:,.2f}")
                
            if es_aduanero:
                st.warning(
                    "⚠️ **Alerta de Importación / Agenciamiento Aduanero**\n\n"
                    "En este documento intervienen gastos por cuenta de terceros y honorarios propios del agente:\n"
                    "• **Honorarios / Comisión:** Gravados con IVA 19% y ReteFuente (4% u 11%).\n"
                    "• **Pagos por cuenta de terceros (Tributos/Fletes/Bodegajes):** Imputables a la cuenta 146505. No llevan IVA del agente ni retención al intermediario."
                )
                
        with col_a2:
            with st.container(border=True):
                st.subheader("Resumen Financiero")
                st.metric("Base Gravable", f"${fac_sel['Valor Base']:,.0f}")
                st.metric("IVA Liquidado", f"${fac_sel['IVA']:,.0f}")
                st.metric("ReteFuente", f"${fac_sel['ReteFuente']:,.0f}")
                st.metric("Total Neto a Pagar", f"${fac_sel['Total'] - fac_sel['ReteFuente']:,.0f}")
    else:
        st.info("Carga el archivo Excel en la Pestaña 1 para habilitar la auditoría.")

with tab_siigo:
    st.markdown("### Exportar Libro Oficial de Importación Siigo Nube (3 Hojas Completas)")
    st.write("Genera el libro Excel estructurado con **matriz_captura**, **interfaz_siigo** (27 columnas oficiales) y **Parametrización**.")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. Hoja matriz_captura
            cols_captura = ["Tipo Comp", "Consecutivo", "Fecha (DD/MM/AAAA)", "NIT", "Prefijo", "Factura Num", "Descripcion", "Operacion", "Cta Principal", "Valor Base", "IVA", "ReteFuente", "ReteICA", "ReteIVA", "Cta Contrapartida"]
            df_matriz_exp = df_p.copy()
            df_matriz_exp["Factura Num"] = df_matriz_exp["Folio"]
            df_matriz_exp = df_matriz_exp[[c for c in cols_captura if c in df_matriz_exp.columns]]
            df_matriz_exp.to_excel(writer, sheet_name="matriz_captura", index=False)
            
            # 2. Hoja interfaz_siigo (27 columnas estándar Siigo)
            filas_interfaz = []
            for _, r in df_p.iterrows():
                es_nc = r.get("Operacion") == "Devolucion Compra"
                val_base = float(r.get("Valor Base", 0.0))
                debito = 0.0 if es_nc else val_base
                credito = val_base if es_nc else 0.0
                
                filas_interfaz.append({
                    "Tipo de comprobante": r.get("Tipo Comp"),
                    "Consecutivo comprobante": r.get("Consecutivo"),
                    "Fecha de elaboración": r.get("Fecha (DD/MM/AAAA)"),
                    "Sigla moneda": "COP",
                    "Tasa de cambio": 1,
                    "Código cuenta contable": str(r.get("Cta Principal")),
                    "Identificación tercero": str(r.get("NIT")),
                    "Sucursal": 0,
                    "Código producto": "",
                    "Código de bodega": "",
                    "Acción": "",
                    "Cantidad producto": "",
                    "Prefijo": r.get("Prefijo", ""),
                    "Consecutivo": r.get("Folio", ""),
                    "No. cuota": 1,
                    "Fecha vencimiento": r.get("Fecha (DD/MM/AAAA)"),
                    "Código impuesto": "",
                    "Código grupo activo fijo": "",
                    "Código activo fijo": "",
                    "Descripción": r.get("Descripcion"),
                    "Código centro/subcentro de costos": "",
                    "Débito": debito,
                    "Crédito": credito,
                    "Observaciones": f"Factura {r.get('Prefijo','')}{r.get('Folio','')}",
                    "Base gravable libro compras/ventas": val_base,
                    "Base exenta libro compras/ventas": 0.0,
                    "Mes de cierre": ""
                })
            df_interfaz = pd.DataFrame(filas_interfaz)
            df_interfaz.to_excel(writer, sheet_name="interfaz_siigo", index=False)
            
            # 3. Hoja Parametrización
            data_puc = [
                {"Cuenta": "146505", "Nombre": "Inventarios en Tránsito / Costos Importación", "Naturaleza": "Débito", "Uso": "DHL, Euro Shipping, Agencias Aduaneras, Puertos"},
                {"Cuenta": "14350101", "Nombre": "Mercancías / Repuestos Mayores a Base UVT", "Naturaleza": "Débito", "Uso": "Inventario repuestos (Base según periodo de emisión)"},
                {"Cuenta": "61800101", "Nombre": "Costos Mantenimiento / Menores a Base UVT", "Naturaleza": "Débito", "Uso": "Consumo y mantenimiento directo de maquinaria"},
                {"Cuenta": "51550501", "Nombre": "Gastos de Viaje y Alojamiento", "Naturaleza": "Débito", "Uso": "Hospedaje y hoteles de operarios / personal (3.5%)"},
                {"Cuenta": "51352001", "Nombre": "Suscripciones y Software", "Naturaleza": "Débito", "Uso": "Plataforma Siigo y software contable (Autorretenedor)"},
                {"Cuenta": "51953001", "Nombre": "Útiles, Papelería y Fotocopias", "Naturaleza": "Débito", "Uso": "Gastos de papelería de oficina"},
                {"Cuenta": "51959501", "Nombre": "Gastos Generales Diversos", "Naturaleza": "Débito", "Uso": "Compras menores y gastos varios"},
                {"Cuenta": "24080101", "Nombre": "Impuesto a las Ventas (IVA) Descontable 19%", "Naturaleza": "Débito", "Uso": "IVA pagado en compras y servicios gravados"},
                {"Cuenta": "23654001", "Nombre": "ReteFuente Compras Declarantes (2.5%)", "Naturaleza": "Crédito", "Uso": "10 UVT ($523.740) / 27 UVT ($1.414.098 en suspensión 8 May-30 Jun)"},
                {"Cuenta": "23652501", "Nombre": "ReteFuente Servicios General (4.0%)", "Naturaleza": "Crédito", "Uso": "2 UVT ($104.748) / 4 UVT ($209.496 en suspensión 8 May-30 Jun)"},
                {"Cuenta": "22050501", "Nombre": "Proveedores Nacionales", "Naturaleza": "Crédito", "Uso": "Contrapartida de compras y materias primas"},
                {"Cuenta": "23359501", "Nombre": "Costos y Gastos por Pagar", "Naturaleza": "Crédito", "Uso": "Contrapartida de servicios y gastos operativos"}
            ]
            df_puc = pd.DataFrame(data_puc)
            df_puc.to_excel(writer, sheet_name="Parametrización", index=False)
            
        output.seek(0)
        st.download_button(
            label=f"📥 Descargar Plantilla Siigo Completa (3 Hojas: matriz, interfaz y parametrización)",
            data=output,
            file_name=f"Plantilla_Siigo_{empresa['nombre'].replace(' ', '_')}_Completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("Vista Previa de la Hoja 'interfaz_siigo' (Para Importación en Siigo):")
        st.dataframe(df_interfaz.head(10), use_container_width=True)
    else:
        st.info("Primero procesa los documentos en la Pestaña 1 para habilitar la descarga.")
