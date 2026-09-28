import streamlit as st
import pandas as pd
import openpyxl
import io
import time
import zipfile
import re
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Sistema ERP y Auditoria Contable DIAN", layout="wide", page_icon="🏢")

# TEMPORIZADOR DE INACTIVIDAD (30 MINUTOS = 1800 SEG)
TIEMPO_MAX_INACTIVIDAD = 30 * 60

# Script inactividad navegador
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
        "actividad": "Comercio y Reparacion de Maquinaria / Importaciones",
        "regimen": "Responsable de IVA",
        "estado": "ACTIVA"
    },
    {
        "nombre": "ASMINCOL S.A.S.",
        "nit": "900.467.519-1",
        "actividad": "Servicios Mineros y Construccion",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    },
    {
        "nombre": "CONSTRUDISENO CT SAS",
        "nit": "900.524.356-8",
        "actividad": "Construccion y Obras Civiles",
        "regimen": "Responsable de IVA",
        "estado": "PROXIMAMENTE"
    },
    {
        "nombre": "SCG TRANSPORTES",
        "nit": "901.700.731-8",
        "actividad": "Transporte de Carga y Logistica",
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

# Inicializacion de variables de sesion
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "empresa_activa" not in st.session_state:
    st.session_state["empresa_activa"] = None
if "proceso_activo" not in st.session_state:
    st.session_state["proceso_activo"] = None
if "ultima_actividad" not in st.session_state:
    st.session_state["ultima_actividad"] = time.time()

# Alerta de sesion expirada por URL
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
    # Si recargo la pagina (F5) y estaba autenticado dentro de los 30 min
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

# PANTALLA 1: LOGIN
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

# Barra superior con boton para Cerrar Sesion manual
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
    "3. Exportar Planilla Oficial a Siigo"
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

def clasificar_factura(nit_emisor, nombre_emisor, valor_base, tipo_doc):
    nombre = str(nombre_emisor).upper()
    es_nc = "CREDITO" in str(tipo_doc).upper() or "CRÉDITO" in str(tipo_doc).upper()
    t_comp = 17 if es_nc else 10
    op = "Devolucion Compra" if es_nc else "Compra"
    
    if any(k in nombre for k in AGENTES_ADUANEROS):
        return t_comp, op, "146505", "22050501", f"Importación / Tránsito - {nombre_emisor[:25]}", round(valor_base * 0.04, 2) if valor_base >= 210000 else 0.0, "Importación (1465)", "Honorarios Agenciamiento vs Terceros"
        
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
        return t_comp, op, "51953001", "23359501", f"Papelería - {nombre_emisor[:25]}", 0.0, "Gastos Papelería", "Útiles de oficina"
        
    if valor_base >= 500000:
        return t_comp, op, "14350101", "22050501", f"Compra mercancías - {nombre_emisor[:25]}", round(valor_base * 0.025, 2) if valor_base >= 1414000 else 0.0, "Mercancía", "Compra general > 500k"
    else:
        return t_comp, op, "51959501", "23359501", f"Gastos generales - {nombre_emisor[:25]}", 0.0, "Gasto General", "Compra menor general"

with tab_compras:
    st.markdown("### 1. Insumos DIAN y Facturas en PDF (Compilado o Individuales)")
    st.write("Sube el archivo Excel de la DIAN (`prueba.xlsx`) o matriz Siigo, y el archivo **PDF compilado** para desglose y renombrado automático por comprobante.")
    
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
            st.success(f"Reporte cargado con éxito: **{len(df_dian)} registros identificados**.")
            
            filas = []
            for idx, r in df_dian.iterrows():
                tipo_doc = r.get("Tipo de documento", r.get("Operacion", "Factura electrónica"))
                folio = str(r.get("Folio", r.get("Factura Num", f"{idx+1}"))).strip()
                prefijo = str(r.get("Prefijo", "")).strip() if pd.notna(r.get("Prefijo")) else ""
                if prefijo.lower() == "nan": prefijo = ""
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
                st.write(f"**Factura:** {fac_sel['Factura']} — **Proveedor:** {fac_sel['Proveedor']} (NIT: {fac_sel['NIT Emisor']})")
                st.write(f"**Fecha de Emisión:** {fac_sel['Fecha']} | **Total:** ${fac_sel['Total']:,.2f}")
                st.markdown("---")
                st.info(f"**Cuenta Asignada:** {fac_sel['Concepto / Cta']} ({fac_sel['Categoría']})\n\n**Motivo Técnico:** {fac_sel['Razón Contable']}")
                st.write(f"• **Base Gravable:** ${fac_sel['Base']:,.2f}")
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
                st.metric("Base Gravable", f"${fac_sel['Base']:,.0f}")
                st.metric("IVA Liquidado", f"${fac_sel['IVA']:,.0f}")
                st.metric("ReteFuente", f"${fac_sel['ReteFuente']:,.0f}")
                st.metric("Total Neto a Pagar", f"${fac_sel['Total'] - fac_sel['ReteFuente']:,.0f}")
    else:
        st.info("Carga el archivo Excel en la Pestaña 1 para habilitar la auditoría.")

with tab_siigo:
    st.markdown("### Descargar Planilla de Importación Oficial para Siigo Nube")
    
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
        st.info("Primero procesa los documentos en la Pestaña 1 para habilitar la descarga.")
