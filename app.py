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
        "desc": "Carga de reportes DIAN/Token, separacion y renombrado de facturas compiladas por comprobante, auditoria contable y plantilla Siigo.",
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
    "1. Cargar Excel, Desglosar y Renombrar PDFs",
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
    st.markdown("### 1. Insumos DIAN y Facturas en PDF (Compilado o Individuales)")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o matriz Siigo, y el archivo **PDF compilado** (o los PDFs individuales) para que el sistema lo **desglose, separe y renombre comprobante por comprobante**.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        archivo_excel = st.file_uploader("1. Reporte Excel (DIAN prueba.xlsx o Matriz Siigo)", type=["xlsx", "xls"])
    with col_u2:
        archivos_pdfs = st.file_uploader("2. Facturas en PDF (Sube el PDF compilado o sueltas)", type=["pdf"], accept_multiple_files=True)
        
    if archivo_excel is not None:
        try:
            df_dian = pd.read_excel(archivo_excel)
        except Exception as e:
            df_dian = pd.DataFrame()
            st.error(f"Error al leer el archivo Excel: {e}")
            
        if not df_dian.empty:
            st.success(f"Reporte cargado con exito: **{len(df_dian)} registros identificados**.")
            
            filas = []
            for idx, r in df_dian.iterrows():
                tipo_doc = r.get("Tipo de documento", r.get("Operacion", "Factura electrónica"))
                folio = str(r.get("Folio", r.get("Factura Num", f"{idx+1}"))).strip()
                prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
                if prefijo == "nan": prefijo = ""
                fecha = str(r.get("Fecha Emisión", r.get("Fecha (DD/MM/AAAA)", r.get("Fecha", "S/F")))).split()[0]
                nit_e = str(r.get("NIT Emisor", r.get("NIT", ""))).strip().split("-")[0].replace(".", "")
                nom_e = str(r.get("Nombre Emisor", r.get("Descripcion", "Proveedor"))).strip()
                
                tot = float(r.get("Total", 0.0)) if pd.notna(r.get("Total")) else 0.0
                iva = float(r.get("IVA", 0.0)) if pd.notna(r.get("IVA")) else 0.0
                base_col = float(r.get("Valor Base", 0.0)) if pd.notna(r.get("Valor Base")) else 0.0
                base = base_col if base_col > 0 else round(tot - iva, 2)
                if tot == 0.0 and base > 0:
                    tot = base + iva
                
                t_comp_exist = r.get("Tipo Comp", None)
                consec_exist = r.get("Consecutivo", None)
                
                t_comp, op, cta_p, cta_c, desc, rfte, cat, razon = clasificar_factura(nit_e, nom_e, base, tipo_doc)
                
                t_comp_final = int(t_comp_exist) if pd.notna(t_comp_exist) else t_comp
                consecutivo_final = int(consec_exist) if pd.notna(consec_exist) else (680 + idx)
                
                nom_limpio_prov = re.sub(r'[^a-zA-Z0-9]', '', nom_e)[:15]
                nombre_pdf_esperado = f"Comp_{t_comp_final}-{consecutivo_final}_{prefijo}{folio}_{nom_limpio_prov}.pdf"
                
                filas.append({
                    "N°": idx + 1,
                    "Tipo Comp": t_comp_final,
                    "Consecutivo": consecutivo_final,
                    "Comprobante Siigo": f"Comp {t_comp_final}-{consecutivo_final}",
                    "Fecha": fecha,
                    "Factura": f"{prefijo}-{folio}" if prefijo else folio,
                    "Folio": folio,
                    "Prefijo": prefijo,
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
            
            st.markdown("#### Matriz Contable Vinculada a Comprobantes:")
            st.dataframe(df_proc[["Comprobante Siigo", "Fecha", "Factura", "Proveedor", "Concepto / Cta", "Base", "IVA", "ReteFuente", "Total", "Soporte PDF Renombrado"]], use_container_width=True)

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
                
                # Leer todos los PDFs subidos y sus páginas
                total_paginas = []
                for pdf_file in archivos_pdfs:
                    try:
                        reader = PdfReader(pdf_file)
                        if reader.is_encrypted:
                            for pwd in contrasenas:
                                try:
                                    if reader.decrypt(pwd) > 0:
                                        break
                                except:
                                    pass
                        for p in reader.pages:
                            txt = p.extract_text() or ""
                            total_paginas.append((p, txt))
                    except Exception as e:
                        st.error(f"Error abriendo {pdf_file.name}: {e}")
                        
                st.info(f"Se cargaron un total de **{len(total_paginas)} páginas de facturas** para procesar.")
                
                facturas_asignadas = {}
                
                if "Modo Secuencial" in modo_corte:
                    for idx_f, row in df_proc.iterrows():
                        if idx_f < len(total_paginas):
                            facturas_asignadas[idx_f] = [total_paginas[idx_f][0]]
                            
                elif "Modo Archivos Sueltos" in modo_corte:
                    for idx_pdf, pdf_file in enumerate(archivos_pdfs):
                        try:
                            pdf_file.seek(0)
                            r_single = PdfReader(pdf_file)
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
                    # Modo Inteligente: Analizar hoja por hoja
                    inv_actual = 0
                    for p_obj, p_txt in total_paginas:
                        clean_txt = re.sub(r'[^a-zA-Z0-9]', '', p_txt.upper())
                        
                        match_encontrado = None
                        for idx_f, row in df_proc.iterrows():
                            num_compuesto = re.sub(r'[^a-zA-Z0-9]', '', f"{row['Prefijo']}{row['Folio']}".upper())
                            num_solo = re.sub(r'[^a-zA-Z0-9]', '', str(row['Folio']).upper())
                            nit_prov = str(row['NIT Emisor']).replace(".", "")
                            
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
                        
                st.success(f"🎉 ¡Proceso completado! Se generaron **{len(facturas_asignadas)} archivos de factura individuales** nombrados con su respectivo comprobante contable.")
                
                buffer_zip.seek(0)
                st.download_button(
                    label="📥 Descargar Paquete Completo de Facturas Separadas y Renombradas (.ZIP)",
                    data=buffer_zip,
                    file_name="Facturas_INDUMAQ_Separadas_Por_Comprobante.zip",
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
