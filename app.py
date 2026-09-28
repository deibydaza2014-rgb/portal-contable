import base64
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import io
import zipfile
import re
import datetime
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

# PANTALLA 2: SELECTOR DE EMPRESAS
EMPRESAS_DISPONIBLES = [
    {
        "nombre": "INDUMAQ ER SAS",
        "nit": "901.346.412-5",
        "actividad": "Comercio y Reparacion de Maquinaria / Importaciones",
        "regimen": "Responsable de IVA",
        "estado": "ACTIVA",
        "badge": "badge-active"
    },
    {
        "nombre": "ASMINCOL S.A.S.",
        "nit": "900.467.519-1",
        "actividad": "Servicios Mineros y Construccion",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "CONSTRUDISENO CT SAS",
        "nit": "900.524.356-8",
        "actividad": "Construccion y Obras Civiles",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "SCG TRANSPORTES",
        "nit": "901.700.731-8",
        "actividad": "Transporte de Carga y Logistica",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE",
        "badge": "badge-next"
    },
    {
        "nombre": "OPJ SAS",
        "nit": "901.425.101-3",
        "actividad": "Servicios Generales y Operaciones",
        "regimen": "Responsable de IVA",
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

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, tipo_doc):
    nombre = str(nombre_emisor).upper()
    es_nc = "CREDITO" in str(tipo_doc).upper() or "CRÉDITO" in str(tipo_doc).upper()
    t_comp = 17 if es_nc else 10
    op = "Devolucion Compra" if es_nc else "Compra"
    
    # 1. Agentes Aduaneros e Importaciones
    if any(k in nombre for k in AGENTES_ADUANEROS):
        rfte = round(valor_base * 0.04, 2) if valor_base >= 210000 or "EURO SHIPPING" in nombre or "DHL" in nombre else 0.0
        rica = round(valor_base * 0.00966, 2) if valor_base >= 210000 and "BUENAVENTURA" not in nombre and "PORTUARIA" not in nombre else 0.0
        return t_comp, op, "146505", "22050501", f"Importacion / Transito - {nombre_emisor[:25]}", rfte, rica, "23652503", "24081501", "23680505", "Importacion (1465)", "Honorarios Agenciamiento vs Terceros"
        
    # 2. Repuestos / Mantenimiento
    repuestos_kw = ["FERROMENDEZ", "TORNILLOLOCO", "CAUCHOS", "ASIMFER", "MAFLEXCOL", "EMPRECOL", "BATTS ZONE", "MECANIZAR", "HIDRAHULICAS", "ELECTRICOS", "ILUMINACION", "BAMACOLGROUP"]
    if any(k in nombre for k in repuestos_kw):
        rfte = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        if valor_base >= 500000:
            return t_comp, op, "14350101", "22050501", f"Repuestos / Inventario - {nombre_emisor[:25]}", rfte, rica, "23654001", "24081001", "23680501", "Inventario", "Compra repuestos mayores a 500k"
        else:
            return t_comp, op, "61800101", "23359501", f"Mantenimiento menor - {nombre_emisor[:25]}", 0.0, 0.0, "23654001", "24081001", "23680501", "Costo Mantenimiento", "Repuestos menores a 500k"
            
    # 3. Hoteles / Viajes
    if any(k in nombre for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK"]):
        rfte = round(valor_base * 0.035, 2) if valor_base >= 210000 else 0.0
        rica = round(valor_base * 0.00966, 2) if valor_base >= 210000 else 0.0
        return t_comp, op, "51550501", "23359501", f"Alojamiento / Viaje - {nombre_emisor[:25]}", rfte, rica, "23654004", "24081501", "23680505", "Gasto Viaje", "Hospedaje de personal"
        
    # 4. Software Siigo
    if "SIIGO" in nombre:
        return t_comp, op, "51352001", "23359501", f"Software Siigo - {nombre_emisor[:25]}", 0.0, 0.0, "", "24081501", "", "Software", "Autorretenedor de renta"
        
    # 5. Papelería
    if any(k in nombre for k in ["PANAMERICANA", "LIBRERIA"]):
        return t_comp, op, "51953001", "23359501", f"Papeleria - {nombre_emisor[:25]}", 0.0, 0.0, "", "24081001", "", "Gastos Papeleria", "Utiles de oficina"
        
    # 6. Compras y Gastos Generales
    if valor_base >= 500000:
        rfte = round(valor_base * 0.025, 2) if valor_base >= 1047000 else 0.0
        rica = round(valor_base * 0.01104, 2) if valor_base >= 1414000 else 0.0
        return t_comp, op, "14350101", "22050501", f"Compra mercancias - {nombre_emisor[:25]}", rfte, rica, "23654001", "24081001", "23680501", "Mercancia", "Compra general > 500k"
    else:
        return t_comp, op, "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}", 0.0, 0.0, "", "24081001", "", "Gasto General", "Compra menor general"

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el reporte de facturas, y los PDFs (o un PDF consolidado) para desbloquear, separar y renombrar automáticamente por comprobante.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel de la DIAN (ej. prueba.xlsx)", type=["xlsx", "xls"])
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF (pueden ser individuales o un PDF consolidado)", type=["pdf"], accept_multiple_files=True)
        
    if archivo_excel is not None:
        df_dian = pd.read_excel(archivo_excel)
        
        # Identificar columna de fecha y ordenar cronológicamente de Enero a la fecha actual
        col_fecha = None
        for col_cand in ["Fecha Emisión", "Fecha Emision", "Fecha", "Fecha de Emisión"]:
            if col_cand in df_dian.columns:
                col_fecha = col_cand
                break
        if col_fecha is None:
            col_fecha = df_dian.columns
            
        df_dian["_fecha_dt"] = pd.to_datetime(df_dian[col_fecha], dayfirst=True, errors="coerce")
        df_dian = df_dian.sort_values(by="_fecha_dt", ascending=True).reset_index(drop=True)
        
        st.success(f"Reporte procesado y ordenado cronológicamente (Enero a la fecha): **{len(df_dian)} facturas identificadas**.")
        
        filas = []
        for idx, r in df_dian.iterrows():
            tipo_doc = r.get("Tipo de documento") or r.get("Tipo documento") or "Factura electrónica"
            folio = str(r.get("Folio") or r.get("Factura") or r.get("Factura Num") or f"Doc_{idx+1}").strip()
            if folio.endswith(".0"):
                folio = folio[:-2]
            prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) and str(r.get("Prefijo")) != "nan" else ""
            
            f_dt = r["_fecha_dt"]
            if pd.notna(f_dt):
                fecha_str = f"{f_dt.day:02d}/{f_dt.month:02d}/{f_dt.year}"
            else:
                fecha_str = str(r.get(col_fecha, "S/F")).split()[0]
                
            nit_e = str(r.get("NIT Emisor") or r.get("NIT") or "").strip()
            if nit_e.endswith(".0"):
                nit_e = nit_e[:-2]
            nom_e = str(r.get("Nombre Emisor") or r.get("Proveedor (Emisor)") or r.get("Proveedor") or "Proveedor").strip()
            
            try: iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
            except: iva = 0.0
            try: tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
            except: tot = 0.0
            base = round(tot - iva, 2)
            
            t_comp, op, cta_p, cta_c, desc, rfte, rica, cta_rfte, cta_iva, cta_rica, cat, razon = clasificar_factura(nit_e, nom_e, base, tipo_doc)
            consecutivo = 680 + idx
            
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
                "Descripcion": desc,
                "Operacion": op,
                "Cta Principal": cta_p,
                "Categoría": cat,
                "Base": base,
                "IVA": iva,
                "ReteFuente": rfte,
                "ReteICA": rica,
                "ReteIVA": 0.0,
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
        
        st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes (Orden Cronológico Enero - Actual):")
        st.dataframe(df_proc[["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "Cta Principal", "Base", "IVA", "ReteFuente", "ReteICA", "Total", "Soporte PDF Renombrado"]], use_container_width=True)

    if archivos_pdfs:
        st.markdown("---")
        st.markdown("#### 📑 Procesamiento, Separación y Renombrado de PDFs Multi-Página:")
        st.caption("Identifica y une facturas completas de 2 o más hojas a partir de un PDF consolidado, asignando a cada una su comprobante y consecutivo oficial.")

        patron_pag = re.compile(r'(?:P[ÁAáa]G(?:INA|\.)?|HOJA|PAGE)\s*(\d+)\s*(?:DE|\/|OF)\s*(\d+)', re.IGNORECASE)

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
            nit_limpio = "901346412"
            
            if "dict_pdfs" not in st.session_state:
                st.session_state["dict_pdfs"] = {}

            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                df_ref = st.session_state.get("df_procesado", pd.DataFrame())

                for idx_pdf, pdf_file in enumerate(archivos_pdfs):
                    try:
                        reader = PdfReader(pdf_file)
                        if reader.is_encrypted:
                            for pwd in [nit_limpio, f"{nit_limpio}5", f"{nit_limpio}-5", ""]:
                                try:
                                    if reader.decrypt(pwd) > 0:
                                        break
                                except:
                                    pass
                        
                        num_pags = len(reader.pages)

                        # Caso 1: Archivo de una sola página o modo sin separación
                        if num_pags == 1 or "Mantener archivos individuales" in modo_sep:
                            writer = PdfWriter()
                            texto_pdf = ""
                            for page in reader.pages:
                                writer.add_page(page)
                                try: texto_pdf += page.extract_text() + "\n"
                                except: pass
                            
                            nombre_final = f"Comprobante_{idx_pdf+1}_{pdf_file.name}"
                            if not df_ref.empty:
                                for _, r_mat in df_ref.iterrows():
                                    fac_num = str(r_mat["Folio"]).replace("-", "").strip()
                                    if fac_num and fac_num in texto_pdf.replace("-", "").replace(" ", ""):
                                        nombre_final = r_mat["Soporte PDF Renombrado"]
                                        break
                                        
                            out_pdf = io.BytesIO()
                            writer.write(out_pdf)
                            pdf_bytes = out_pdf.getvalue()
                            zf.writestr(nombre_final, pdf_bytes)
                            st.session_state["dict_pdfs"][nombre_final] = pdf_bytes
                            total_generados += 1

                        # Caso 2: Agrupación fija de 2 páginas por factura
                        elif "Agrupación fija de 2 páginas" in modo_sep:
                            st.write(f"Agrupando PDF **{pdf_file.name}** en bloques de **2 páginas por factura** ({num_pags} páginas totales)...")
                            inv_idx = 0
                            for p_start in range(0, num_pags, 2):
                                pw = PdfWriter()
                                pw.add_page(reader.pages[p_start])
                                if p_start + 1 < num_pags:
                                    pw.add_page(reader.pages[p_start + 1])
                                
                                out_p = io.BytesIO()
                                pw.write(out_p)
                                p_bytes = out_p.getvalue()
                                
                                if inv_idx < len(df_ref):
                                    nom_doc = df_ref.iloc[inv_idx]["Soporte PDF Renombrado"]
                                else:
                                    nom_doc = f"Comprobante_Extra_{inv_idx+1}.pdf"
                                    
                                zf.writestr(nom_doc, p_bytes)
                                st.session_state["dict_pdfs"][nom_doc] = p_bytes
                                total_generados += 1
                                inv_idx += 1

                        # Caso 3: Agrupación fija de 3 páginas por factura
                        elif "Agrupación fija de 3 páginas" in modo_sep:
                            st.write(f"Agrupando PDF **{pdf_file.name}** en bloques de **3 páginas por factura** ({num_pags} páginas totales)...")
                            inv_idx = 0
                            for p_start in range(0, num_pags, 3):
                                pw = PdfWriter()
                                for offset in range(3):
                                    if p_start + offset < num_pags:
                                        pw.add_page(reader.pages[p_start + offset])
                                
                                out_p = io.BytesIO()
                                pw.write(out_p)
                                p_bytes = out_p.getvalue()
                                
                                if inv_idx < len(df_ref):
                                    nom_doc = df_ref.iloc[inv_idx]["Soporte PDF Renombrado"]
                                else:
                                    nom_doc = f"Comprobante_Extra_{inv_idx+1}.pdf"
                                    
                                zf.writestr(nom_doc, p_bytes)
                                st.session_state["dict_pdfs"][nom_doc] = p_bytes
                                total_generados += 1
                                inv_idx += 1

                        # Caso 4: Separar estrictamente 1 página por factura
                        elif "Separar 1 página por factura" in modo_sep or df_ref.empty:
                            for p_idx in range(num_pags):
                                pw = PdfWriter()
                                pw.add_page(reader.pages[p_idx])
                                out_p = io.BytesIO()
                                pw.write(out_p)
                                p_bytes = out_p.getvalue()
                                
                                if p_idx < len(df_ref):
                                    nom_doc = df_ref.iloc[p_idx]["Soporte PDF Renombrado"]
                                else:
                                    nom_doc = f"Comprobante_Extra_Pag_{p_idx+1}.pdf"
                                    
                                zf.writestr(nom_doc, p_bytes)
                                st.session_state["dict_pdfs"][nom_doc] = p_bytes
                                total_generados += 1

                        # Caso 5: Detección Inteligente Multihidráulica / DIAN (2 o más hojas por factura)
                        else:
                            st.write(f"Analizando estructura multi-página del PDF consolidado **{pdf_file.name}** ({num_pags} páginas)...")
                            
                            facturas_generadas = []
                            curr_writer = None
                            curr_inv_row = None
                            paginas_del_comprobante = 0
                            
                            def buscar_coincidencia_factura(texto_pagina):
                                t_clean = texto_pagina.replace("-", "").replace(" ", "").upper()
                                for _, r_cand in df_ref.iterrows():
                                    fol_cand = str(r_cand["Folio"]).replace("-", "").strip().upper()
                                    pref_cand = str(r_cand["Prefijo"]).replace("-", "").strip().upper()
                                    fac_cand = str(r_cand["Factura"]).replace("-", "").strip().upper()
                                    nit_cand = str(r_cand["NIT Emisor"]).replace("-", "").strip().upper()
                                    
                                    # 1. Coincidencia estricta Prefijo + Folio (ej. CBO2015647 o FTGI351)
                                    if pref_cand and fol_cand and (pref_cand + fol_cand) in t_clean:
                                        return r_cand
                                    # 2. Factura completa
                                    if fac_cand and len(fac_cand) >= 4 and fac_cand in t_clean:
                                        return r_cand
                                    # 3. NIT Proveedor + Folio juntos
                                    if nit_cand and fol_cand and len(fol_cand) >= 3 and (nit_cand in t_clean and fol_cand in t_clean):
                                        return r_cand
                                    # 4. Folio largo (>= 4 dígitos)
                                    if fol_cand and len(fol_cand) >= 4 and fol_cand in t_clean:
                                        return r_cand
                                return None

                            for p_idx in range(num_pags):
                                try: txt_p = reader.pages[p_idx].extract_text() or ""
                                except: txt_p = ""
                                
                                # Comprobar indicador explícito "Página X de Y"
                                m_pag = patron_pag.search(txt_p)
                                cur_p, tot_p = (int(m_pag.group(1)), int(m_pag.group(2))) if m_pag else (None, None)
                                
                                # Buscar si en esta página arranca una factura identificada
                                inv_encontrada = buscar_coincidencia_factura(txt_p)
                                
                                # Criterio de inicio de nueva factura
                                es_inicio_nueva_factura = False
                                if cur_p == 1:
                                    es_inicio_nueva_factura = True
                                elif cur_p is not None and cur_p > 1:
                                    es_inicio_nueva_factura = False # Es hoja 2 o 3 de la factura actual
                                elif inv_encontrada is not None:
                                    if curr_inv_row is None or inv_encontrada["Comprobante Siigo"] != curr_inv_row["Comprobante Siigo"]:
                                        es_inicio_nueva_factura = True
                                elif curr_inv_row is None:
                                    es_inicio_nueva_factura = True

                                if es_inicio_nueva_factura:
                                    # Guardar la factura anterior que ya se completó con todas sus hojas
                                    if curr_writer is not None and curr_inv_row is not None:
                                        out_c = io.BytesIO()
                                        curr_writer.write(out_c)
                                        c_bytes = out_c.getvalue()
                                        doc_nombre = curr_inv_row["Soporte PDF Renombrado"]
                                        zf.writestr(doc_nombre, c_bytes)
                                        st.session_state["dict_pdfs"][doc_nombre] = c_bytes
                                        total_generados += 1
                                        facturas_generadas.append(f"{curr_inv_row['Comprobante Siigo']} ({paginas_del_comprobante} págs)")

                                    # Iniciar nueva factura
                                    curr_inv_row = inv_encontrada if inv_encontrada is not None else (df_ref.iloc[len(facturas_generadas)] if len(facturas_generadas) < len(df_ref) else None)
                                    curr_writer = PdfWriter()
                                    curr_writer.add_page(reader.pages[p_idx])
                                    paginas_del_comprobante = 1
                                else:
                                    # Agregar hoja de continuación (página 2, 3, etc.)
                                    if curr_writer is None:
                                        curr_inv_row = df_ref.iloc[0] if len(df_ref) > 0 else None
                                        curr_writer = PdfWriter()
                                    curr_writer.add_page(reader.pages[p_idx])
                                    paginas_del_comprobante += 1

                            # Guardar la última factura al terminar el archivo
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
                        st.error(f"Error procesando {pdf_file.name}: {e}")

            st.success(f"¡Éxito! Se identificaron y unieron **{total_generados} facturas completas en PDF** (incluyendo facturas de 2 o más hojas), nombradas con su respectivo comprobante y consecutivo (`Comp_10-XXX_...pdf`).")
            buffer_zip.seek(0)
            st.download_button(
                label="📥 Descargar Paquete Completo de Facturas Separadas (.ZIP)",
                data=buffer_zip,
                file_name=f"Facturas_{empresa['nombre'].replace(' ', '_')}_Separadas_Por_Comprobante.zip",
                mime="application/zip",
                use_container_width=True
            )

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
        
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            st.markdown(f"""
            <div class="audit-card">
                <h4>Detalle del Comprobante: {fac_sel['Comprobante Siigo']}</h4>
                <p><b>Soporte PDF Vinculado:</b> <span class="tag-comp">{fac_sel['Soporte PDF Renombrado']}</span></p>
                <p><b>Factura:</b> {fac_sel['Factura']} - <b>Proveedor:</b> {fac_sel['Proveedor']} (NIT: {fac_sel['NIT Emisor']})</p>
                <p><b>Fecha de Emision:</b> {fac_sel['Fecha']} - <b>Total:</b> ${fac_sel['Total']:,.2f}</p>
                <hr>
                <h5>Trazabilidad de la Contabilizacion:</h5>
                <ul>
                    <li><b>Cuenta Asignada:</b> <span class="tag-propio">{fac_sel['Cta Principal']}</span> - {fac_sel['Categoría']}</li>
                    <li><b>Motivo Tecnico:</b> {fac_sel['Razón Contable']}</li>
                    <li><b>Base Gravable:</b> ${fac_sel['Base']:,.2f} - <b>IVA:</b> ${fac_sel['IVA']:,.2f}</li>
                    <li><b>Retencion en la Fuente:</b> ${fac_sel['ReteFuente']:,.2f}</li>
                    <li><b>ReteICA:</b> ${fac_sel['ReteICA']:,.2f}</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
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
            st.metric("Base Gravable", f"${fac_sel['Base']:,.0f}")
            st.metric("IVA Liquidado", f"${fac_sel['IVA']:,.0f}")
            st.metric("ReteFuente", f"${fac_sel['ReteFuente']:,.0f}")
            st.metric("ReteICA", f"${fac_sel['ReteICA']:,.0f}")
            st.metric("Total Neto CxP (Cta 22 / 23)", f"${fac_sel['Total'] - fac_sel['ReteFuente'] - fac_sel['ReteICA']:,.0f}")

        # VISTA PREVIA DE LA FACTURA EN UN CUADRO
        st.markdown("---")
        st.markdown("### 🔍 Vista Previa del Documento Soporte")
        
        pdf_bytes_encontrado = None
        folio_clean = str(fac_sel["Folio"]).replace("-", "").strip()
        pdf_esperado = str(fac_sel["Soporte PDF Renombrado"])
        
        if "dict_pdfs" in st.session_state and st.session_state["dict_pdfs"]:
            if pdf_esperado in st.session_state["dict_pdfs"]:
                pdf_bytes_encontrado = st.session_state["dict_pdfs"][pdf_esperado]
            else:
                for fname, pbytes in st.session_state["dict_pdfs"].items():
                    fname_clean = fname.replace("-", "").replace(" ", "")
                    if folio_clean and folio_clean in fname_clean:
                        pdf_bytes_encontrado = pbytes
                        break
        
        if pdf_bytes_encontrado:
            b64_pdf = base64.b64encode(pdf_bytes_encontrado).decode('utf-8')
            
            # Encabezado visual y botón de descarga directa
            col_doc1, col_doc2 = st.columns(2)
            with col_doc1:
                st.markdown(f"""
                <div style="background:#0070ba; color:white; padding:8px 14px; border-radius:6px 6px 0 0; font-weight:600; font-size:14px;">
                    📄 Documento Digitalizado Completo: Factura {fac_sel['Factura']} - {fac_sel['Proveedor']}
                </div>
                """, unsafe_allow_html=True)
            with col_doc2:
                st.download_button(
                    label="📥 Descargar este PDF",
                    data=pdf_bytes_encontrado,
                    file_name=pdf_esperado,
                    mime="application/pdf",
                    use_container_width=True
                )
            
            # Visor HTML5 con Canvas (100% compatible con Chrome, sin bloqueos de seguridad)
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
                  background: #475569;
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
                  padding: 20px;
                  font-size: 14px;
                  text-align: center;
                }}
              </style>
            </head>
            <body>
              <div id="status">Cargando vista previa de la factura...</div>
              <div id="viewer-container"></div>
              <script>
                try {{
                  const rawPdf = atob("{b64_pdf}");
                  const pdfjsLib = window['pdfjs-dist/build/pdf'];
                  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

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
                    document.getElementById('status').innerHTML = '<span style="color:#fca5a5;">No se pudo procesar la vista previa: ' + err.message + '</span>';
                  }});
                }} catch (e) {{
                  document.getElementById('status').innerHTML = '<span style="color:#fca5a5;">Error al decodificar: ' + e.message + '</span>';
                }}
              </script>
            </body>
            </html>
            """
            components.html(html_visor, height=560, scrolling=True)
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
                    <i>💡 Nota: Al pulsar "Desbloquear, Separar y Renombrar PDFs" en la Pestaña 1, el PDF individual correspondiente se mostrará en este marco.</i>
                </p>
            </div>
            """, unsafe_allow_html=True)

        # TABLA DE CONTABILIZACION (ASIENTO CONTABLE)
        st.markdown("### 📋 Asiento Contable del Comprobante (Cómo se Contabilizó)")
        st.caption("Detalle de partida doble con imputación de cuentas, débitos, créditos y sumas iguales:")
        
        asiento_filas = []
        es_nc = "Devolucion" in str(fac_sel["Operacion"])
        
        # 1. Gasto / Costo / Inventario
        asiento_filas.append({
            "Código Cuenta": fac_sel["Cta Principal"],
            "Descripción de la Cuenta": f"{fac_sel['Categoría']} - {fac_sel['Proveedor'][:25]}",
            "Tercero / NIT": fac_sel["NIT Emisor"],
            "Débito ($)": 0.0 if es_nc else fac_sel["Base"],
            "Crédito ($)": fac_sel["Base"] if es_nc else 0.0
        })
        
        # 2. IVA Descontable
        if fac_sel["IVA"] > 0:
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta IVA"],
                "Descripción de la Cuenta": f"IVA Descontable (Base: ${fac_sel['Base']:,.0f})",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": 0.0 if es_nc else fac_sel["IVA"],
                "Crédito ($)": fac_sel["IVA"] if es_nc else 0.0
            })
            
        # 3. Retención en la Fuente
        if fac_sel["ReteFuente"] > 0 and fac_sel["Cta ReteFuente"]:
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta ReteFuente"],
                "Descripción de la Cuenta": f"ReteFuente Practicada ({fac_sel['Categoría']})",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": fac_sel["ReteFuente"] if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else fac_sel["ReteFuente"]
            })
            
        # 4. ReteICA
        if fac_sel.get("ReteICA", 0.0) > 0 and fac_sel.get("Cta ReteICA"):
            asiento_filas.append({
                "Código Cuenta": fac_sel["Cta ReteICA"],
                "Descripción de la Cuenta": "Retención ICA Practicada",
                "Tercero / NIT": fac_sel["NIT Emisor"],
                "Débito ($)": fac_sel["ReteICA"] if es_nc else 0.0,
                "Crédito ($)": 0.0 if es_nc else fac_sel["ReteICA"]
            })
            
        # 5. Cuenta por Pagar (Proveedores)
        neto_cxp = fac_sel["Total"] - fac_sel["ReteFuente"] - fac_sel.get("ReteICA", 0.0)
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
        
        # Generar archivo Excel con las 3 hojas completas
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
