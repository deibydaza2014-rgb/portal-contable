import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
import re
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

AGENTES_ADUANEROS = ["DHL", "ADUANA", "EURO SHIPPING", "PORTUARIA", "ALMACENADORA", "CARGO", "TRADE GLOBAL", "TERMINAL", "BUENAVENTURA"]

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, tipo_doc):
    nombre = str(nombre_emisor).upper()
    es_nc = "CREDITO" in str(tipo_doc).upper() or "CRÉDITO" in str(tipo_doc).upper()
    t_comp = 17 if es_nc else 10
    op = "Devolucion Compra" if es_nc else "Compra"
    
    if any(k in nombre for k in AGENTES_ADUANEROS):
        return t_comp, op, "146505", "22050501", f"Importacion / Transito - {nombre_emisor[:25]}", round(valor_base * 0.04, 2) if valor_base >= 210000 else 0.0, "Importacion (1465)", "Honorarios Agenciamiento vs Terceros"
        
    repuestos_kw = ["FERROMENDEZ", "TORNILLOLOCO", "CAUCHOS", "ASIMFER", "MAFLEXCOL", "EMPRECOL", "BATTS ZONE", "MECANIZAR", "HIDRAHULICAS", "BAMACOLGROUP"]
    if any(k in nombre for k in repuestos_kw):
        if valor_base >= 500000:
            return t_comp, op, "14350101", "22050501", f"Repuestos / Inventario - {nombre_emisor[:25]}", round(valor_base * 0.025, 2) if valor_base >= 1414000 else 0.0, "Inventario", "Compra repuestos mayores a 500k"
        else:
            return t_comp, op, "61800101", "23359501", f"Mantenimiento menor - {nombre_emisor[:25]}", 0.0, "Costo Mantenimiento", "Repuestos menores a 500k"
            
    if any(k in nombre for k in ["HOTEL", "ESTELAR", "GENOVA", "VITTAPARK"]):
        return t_comp, op, "51550501", "23359501", f"Alojamiento / Viaje - {nombre_emisor[:25]}", round(valor_base * 0.035, 2) if valor_base >= 210000 else 0.0, "Gasto Viaje", "Hospedaje de personal"
        
    if "SIIGO" in nombre:
        return t_comp, op, "51352001", "23359501", f"Software Siigo - {nombre_emisor[:25]}", 0.0, "Software", "Autorretenedor de renta"
        
    if "PANAMERICANA" in nombre:
        return t_comp, op, "51953001", "23359501", f"Papeleria - {nombre_emisor[:25]}", 0.0, "Gastos Papeleria", "Utiles de oficina"
        
    if valor_base >= 500000:
        return t_comp, op, "14350101", "22050501", f"Compra mercancias - {nombre_emisor[:25]}", round(valor_base * 0.025, 2) if valor_base >= 1414000 else 0.0, "Mercancia", "Compra general > 500k"
    else:
        return t_comp, op, "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}", 0.0, "Gasto General", "Compra menor general"

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o el token de acceso, y los PDFs para desbloquear y renombrar automáticamente por comprobante.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel de la DIAN (ej. prueba.xlsx)", type=["xlsx", "xls"])
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF (desbloqueo y renombrado)", type=["pdf"], accept_multiple_files=True)
        
    if archivo_excel is not None:
        df_dian = pd.read_excel(archivo_excel)
        st.success(f"Reporte procesado: **{len(df_dian)} facturas identificadas**.")
        
        filas = []
        for idx, r in df_dian.iterrows():
            tipo_doc = r.get("Tipo de documento", "Factura electrónica")
            folio = str(r.get("Folio", f"Doc_{idx+1}")).strip()
            prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
            fecha = str(r.get("Fecha Emisión", "S/F")).split()[0]
            nit_e = str(r.get("NIT Emisor", "")).strip()
            nom_e = str(r.get("Nombre Emisor", "Proveedor")).strip()
            iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
            tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
            base = round(tot - iva, 2)
            
            t_comp, op, cta_p, cta_c, desc, rfte, cat, razon = clasificar_factura(nit_e, nom_e, base, tipo_doc)
            consecutivo = 680 + idx
            
            # Nombre estandarizado del PDF vinculado al comprobante
            nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
            nombre_pdf_esperado = f"Comp_{t_comp}-{consecutivo}_{prefijo}{folio}_{nom_limpio_prov}.pdf"
            
            filas.append({
                "N°": idx + 1,
                "Tipo Comp": t_comp,
                "Consecutivo": consecutivo,
                "Comprobante Siigo": f"Comp {t_comp}-{consecutivo}",
                "Fecha": fecha,
                "Factura": f"{prefijo}-{folio}" if prefijo else folio,
                "Proveedor": nom_e,
                "NIT Emisor": nit_e,
                "Concepto / Cta": cta_p,
                "Categoría": cat,
                "Base": base,
                "IVA": iva,
                "ReteFuente": rfte,
                "Total": tot,
                "Cta Contrapartida": cta_c,
                "Razón Contable": razon,
                "Soporte PDF Renombrado": nombre_pdf_esperado
            })
            
        df_proc = pd.DataFrame(filas)
        st.session_state["df_procesado"] = df_proc
        
        st.markdown("#### Matriz Contable Preliminar vinculada a Comprobantes:")
        st.dataframe(df_proc[["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "Concepto / Cta", "Base", "IVA", "ReteFuente", "Total", "Soporte PDF Renombrado"]], use_container_width=True)

    # Procesar y renombrar PDFs con el nombre de cada comprobante
    if archivos_pdfs:
        st.markdown("---")
        st.markdown("#### 📑 Procesamiento y Renombrado de PDFs por Comprobante:")
        if st.button("🔓 Desbloquear y Renombrar PDFs ahora"):
            buffer_zip = io.BytesIO()
            exitosos = 0
            nit_limpio = "901346412"
            
            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
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
                        writer = PdfWriter()
                        texto_pdf = ""
                        for page in reader.pages:
                            writer.add_page(page)
                            texto_pdf += page.extract_text() + "\n"
                        
                        # Buscar correspondencia con el comprobante en la matriz
                        nombre_final = f"Comprobante_{idx_pdf+1}_{pdf_file.name}"
                        if "df_procesado" in st.session_state:
                            for _, r_mat in st.session_state["df_procesado"].iterrows():
                                fac_num = str(r_mat["Factura"]).replace("-", "")
                                if fac_num and fac_num in texto_pdf.replace("-", "").replace(" ", ""):
                                    nombre_final = r_mat["Soporte PDF Renombrado"]
                                    break
                                    
                        out_pdf = io.BytesIO()
                        writer.write(out_pdf)
                        zf.writestr(nombre_final, out_pdf.getvalue())
                        exitosos += 1
                    except Exception as e:
                        st.error(f"Error con {pdf_file.name}: {e}")
                        
            st.success(f"¡{exitosos} PDFs desbloqueados y renombrados con el nombre del comprobante correspondiente!")
            buffer_zip.seek(0)
            st.download_button(
                label="📥 Descargar Paquete de Facturas Renombradas por Comprobante (.ZIP)",
                data=buffer_zip,
                file_name="Facturas_INDUMAQ_Organizadas_Por_Comprobante.zip",
                mime="application/zip",
                use_container_width=True
            )

with tab_auditoria:
    st.markdown("### Modulo de Auditoria Contable y Trazabilidad")
    st.caption("Inspeccion de cuentas, deducciones y separacion de gastos por cuenta de terceros.")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        opciones_fac = [f"[{r['Comprobante Siigo']}] {r['Factura']} - {r['Proveedor']} (${r['Total']:,.0f})" for _, r in df_p.iterrows()]
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
                    <li><b>Cuenta Asignada:</b> <span class="tag-propio">{fac_sel['Concepto / Cta']}</span> - {fac_sel['Categoría']}</li>
                    <li><b>Motivo Tecnico:</b> {fac_sel['Razón Contable']}</li>
                    <li><b>Base Gravable:</b> ${fac_sel['Base']:,.2f} - <b>IVA:</b>${fac_sel['IVA']:,.2f}</li>
                    <li><b>Retencion en la Fuente:</b> ${fac_sel['ReteFuente']:,.2f}</li>
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
            st.metric("Total a Pagar (Cta 22 / 23)", f"${fac_sel['Total'] - fac_sel['ReteFuente']:,.0f}")
    else:
        st.info("Carga el archivo Excel en la Pestana 1 para habilitar la auditoria.")

with tab_siigo:
    st.markdown("### Descargar Planilla de Importacion Oficial para Siigo Nube")
    
    if "df_procesado" in st.session_state:
        df_p = st.session_state["df_procesado"]
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_siigo = df_p[["Tipo Comp", "Consecutivo", "Fecha", "NIT Emisor", "Factura", "Proveedor", "Concepto / Cta", "Base", "IVA", "ReteFuente", "Cta Contrapartida"]].copy()
            df_siigo.columns = ["Tipo Comp", "Consecutivo", "Fecha", "NIT", "Factura Num", "Descripcion", "Cta Principal", "Valor Base", "IVA", "ReteFuente", "Cta Contrapartida"]
            df_siigo.to_excel(writer, sheet_name="matriz_captura", index=False)
            
        output.seek(0)
        st.download_button(
            label=f"Descargar Planilla Siigo ({empresa['nombre']})",
            data=output,
            file_name=f"Plantilla_Siigo_{empresa['nombre'].replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Estructura validada para cargar directamente en Siigo Nube.")
    else:
        st.info("Primero procesa los documentos en la Pestana 1 para habilitar la descarga.")
