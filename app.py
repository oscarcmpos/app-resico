import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import pdfplumber
import re
from supabase import create_client, Client

# Inicializar variable de sesión para el botón de Limpiar/Regresar
if "reset_key" not in st.session_state:
    st.session_state.reset_key = 0

# ==========================================
# 0. CONEXIÓN A SUPABASE
# ==========================================
@st.cache_resource
def init_connection():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        return None

supabase = init_connection()

# ==========================================
# 1. FUNCIONES DE PROCESAMIENTO (CÁLCULO Y REDONDEO)
# ==========================================
def procesar_ingreso_resico(archivo_subido):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        
        emisor = root.find('cfdi:Emisor', ns)
        if emisor is None or emisor.attrib.get('RegimenFiscal') != '626': return None 
            
        folio = root.attrib.get('Folio', 'Sin Folio')
        subtotal = float(root.attrib.get('SubTotal', 0.0))
        
        iva_trasladado, isr_retenido, iva_retenido = 0.0, 0.0, 0.0
        
        impuestos = root.find('cfdi:Impuestos', ns)
        if impuestos is not None:
            traslados = impuestos.find('cfdi:Traslados', ns)
            if traslados is not None:
                for t in traslados.findall('cfdi:Traslado', ns):
                    if t.attrib.get('Impuesto') == '002': iva_trasladado += float(t.attrib.get('Importe', 0.0))
            
            retenciones = impuestos.find('cfdi:Retenciones', ns)
            if retenciones is not None:
                for r in retenciones.findall('cfdi:Retencion', ns):
                    if r.attrib.get('Impuesto') == '001': isr_retenido += float(r.attrib.get('Importe', 0.0))
                    elif r.attrib.get('Impuesto') == '002': iva_retenido += float(r.attrib.get('Importe', 0.0))
                    
        return {"Archivo": archivo_subido.name, "Folio": folio, "Subtotal": subtotal, 
                "IVA Trasladado": iva_trasladado, "ISR Retenido": isr_retenido, "IVA Retenido": iva_retenido}
    except:
        return None

def procesar_gasto_inteligente(archivo_subido):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        
        emisor = root.find('cfdi:Emisor', ns)
        rfc_emisor = emisor.attrib.get('Rfc', '') if emisor is not None else ''
        nombre_emisor = emisor.attrib.get('Nombre', 'Desconocido') if emisor is not None else 'Desconocido'
        regimen_emisor = emisor.attrib.get('RegimenFiscal', '') if emisor is not None else ''

        receptor = root.find('cfdi:Receptor', ns)
        uso_cfdi = receptor.attrib.get('UsoCFDI', '') if receptor is not None else ''

        subtotal = float(root.attrib.get('SubTotal', 0.0))
        iva_acreditable = 0.0
        
        impuestos = root.find('cfdi:Impuestos', ns)
        if impuestos is not None:
            traslados = impuestos.find('cfdi:Traslados', ns)
            if traslados is not None:
                for t in traslados.findall('cfdi:Traslado', ns):
                    if t.attrib.get('Impuesto') == '002': 
                        iva_acreditable += float(t.attrib.get('Importe', 0.0))

        estado = "Acreditable"
        motivo = "Válido para RESICO"

        if regimen_emisor == '605':
            estado = "Excluido"
            motivo = "Régimen 605 (Sueldos y Salarios no acreditable)"

        usos_personales = ['D01', 'D02', 'D03', 'D04', 'D05', 'D06', 'D07', 'D08', 'D09', 'D10']
        if uso_cfdi in usos_personales:
            estado = "Excluido"
            motivo = f"Uso de CFDI ({uso_cfdi}) es deducción personal"

        return {
            "Archivo": archivo_subido.name,
            "Proveedor": nombre_emisor,
            "RFC Emisor": rfc_emisor,
            "Régimen": regimen_emisor,
            "Uso CFDI": uso_cfdi,
            "Subtotal Gasto": subtotal,
            "IVA Acreditable": iva_acreditable if estado == "Acreditable" else 0.0,
            "Estado": estado,
            "Motivo Rechazo": motivo
        }
    except:
        return None

def calcular_isr_resico(ingresos_totales):
    if ingresos_totales <= 25000.00: tasa = 0.01
    elif ingresos_totales <= 50000.00: tasa = 0.011
    elif ingresos_totales <= 83333.33: tasa = 0.015
    elif ingresos_totales <= 208333.33: tasa = 0.02
    else: tasa = 0.025
    return ingresos_totales * tasa, tasa

def generar_excel_formulado(df_ingresos, df_gastos, df_excluidos, totales):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Pestaña de Resumen con Fórmulas nativas de Excel
        wb_resumen_data = [
            ["Concepto", "Importe"],
            ["Ingresos Base", totales['ingresos']],
            ["Tasa ISR (%)", totales['tasa'] * 100],
            ["ISR Determinado", "=B2*B3"],
            ["ISR Retenido", totales['isr_retenido']],
            ["ISR a Pagar", "=MAX(0, B4-B5)"],
            ["---", 0],
            ["IVA Trasladado", totales['iva_trasladado']],
            ["IVA Acreditable (Filtrado)", totales['iva_acreditable']],
            ["IVA Retenido", totales['iva_retenido']],
            [f"IVA a {totales['estatus_iva']}", "=B8-B9-B10"]
        ]
        df_resumen = pd.DataFrame(wb_resumen_data[1:], columns=wb_resumen_data[0])
        df_resumen.to_excel(writer, sheet_name='Resumen Fiscal', index=False)
        
        if not df_ingresos.empty: df_ingresos.to_excel(writer, sheet_name='Detalle Ingresos', index=False)
        if df_gastos is not None and not df_gastos.empty: df_gastos.to_excel(writer, sheet_name='Gastos Acreditables', index=False)
        if df_excluidos is not None and not df_excluidos.empty: df_excluidos.to_excel(writer, sheet_name='Gastos Excluidos', index=False)
    return buffer.getvalue()

# ==========================================
# 2. FUNCIONES DE EXTRACCIÓN AVANZADA (VISOR)
# ==========================================
def obtener_metadatos_basicos(archivo_subido, tipo_archivo):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        
        fecha = root.attrib.get('Fecha', '')[:10]
        total = root.attrib.get('Total', '0.00')
        tipo = root.attrib.get('TipoDeComprobante', '')
        metodo = root.attrib.get('MetodoPago', 'N/A')
        
        if tipo == 'P': metodo = 'Pago (Complemento)'
            
        if tipo_archivo == 'ingreso':
            nodo = root.find('cfdi:Receptor', ns)
            nombre = nodo.attrib.get('Nombre', 'Cliente Desconocido') if nodo is not None else 'Cliente Desconocido'
        else:
            nodo = root.find('cfdi:Emisor', ns)
            nombre = nodo.attrib.get('Nombre', 'Proveedor Desconocido') if nodo is not None else 'Proveedor Desconocido'
        
        return {
            "nombre_archivo": archivo_subido.name,
            "fecha": fecha,
            "entidad": nombre,
            "total": float(total),
            "metodo": metodo,
            "display": f"{fecha} | {nombre} | ${float(total):,.2f} | {metodo}"
        }
    except:
        return None

def extraer_cfdi_completo(archivo_subido):
    archivo_subido.seek(0)
    tree = ET.parse(archivo_subido)
    root = tree.getroot()
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4', 'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}
    
    datos = {
        "fecha": root.attrib.get('Fecha', ''),
        "folio": root.attrib.get('Folio', 'S/F'),
        "serie": root.attrib.get('Serie', ''),
        "subtotal": float(root.attrib.get('SubTotal', '0.00')),
        "descuento": float(root.attrib.get('Descuento', '0.00')),
        "total": float(root.attrib.get('Total', '0.00')),
        "moneda": root.attrib.get('Moneda', 'MXN'),
        "uuid": "No encontrado",
        "conceptos": [],
        "iva_trasladado": 0.0,
        "isr_retenido": 0.0,
        "iva_retenido": 0.0
    }
    
    emisor = root.find('cfdi:Emisor', ns)
    datos["emi_rfc"] = emisor.attrib.get('Rfc', '') if emisor is not None else ''
    datos["emi_nombre"] = emisor.attrib.get('Nombre', '') if emisor is not None else ''
    datos["emi_regimen"] = emisor.attrib.get('RegimenFiscal', '') if emisor is not None else ''
    
    receptor = root.find('cfdi:Receptor', ns)
    datos["rec_rfc"] = receptor.attrib.get('Rfc', '') if receptor is not None else ''
    datos["rec_nombre"] = receptor.attrib.get('Nombre', '') if receptor is not None else ''
    datos["rec_uso"] = receptor.attrib.get('UsoCFDI', '') if receptor is not None else ''
    
    complemento = root.find('cfdi:Complemento', ns)
    if complemento is not None:
        tfd = complemento.find('tfd:TimbreFiscalDigital', ns)
        if tfd is not None: datos["uuid"] = tfd.attrib.get('UUID', '')
            
    conceptos = root.find('cfdi:Conceptos', ns)
    if conceptos is not None:
        for c in conceptos.findall('cfdi:Concepto', ns):
            datos["conceptos"].append({
                "cantidad": c.attrib.get('Cantidad', '0'),
                "descripcion": c.attrib.get('Descripcion', 'Sin descripción'),
                "valor_unitario": float(c.attrib.get('ValorUnitario', '0.00')),
                "importe": float(c.attrib.get('Importe', '0.00'))
            })
            
    impuestos = root.find('cfdi:Impuestos', ns)
    if impuestos is not None:
        traslados = impuestos.find('cfdi:Traslados', ns)
        if traslados is not None:
            for t in traslados.findall('cfdi:Traslado', ns):
                if t.attrib.get('Impuesto') == '002': datos["iva_trasladado"] += float(t.attrib.get('Importe', '0.00'))
                    
        retenciones = impuestos.find('cfdi:Retenciones', ns)
        if retenciones is not None:
            for r in retenciones.findall('cfdi:Retencion', ns):
                if r.attrib.get('Impuesto') == '001': datos["isr_retenido"] += float(r.attrib.get('Importe', '0.00'))
                elif r.attrib.get('Impuesto') == '002': datos["iva_retenido"] += float(r.attrib.get('Importe', '0.00'))
                    
    return datos

# ==========================================
# 3. INTERFAZ WEB PRINCIPAL
# ==========================================
st.set_page_config(page_title="App RESICO", layout="wide")

if supabase is None:
    st.error("⚠️ No se pudo conectar a la base de datos.")
    st.stop()

col_titulo, col_boton = st.columns([4, 1])
with col_titulo:
    st.title("🧮 Sistema Contable Automatizado RESICO")
with col_boton:
    st.write("") 
    if st.button("🔄 Limpiar Todo / Regresar", use_container_width=True):
        st.session_state.reset_key += 1
        st.rerun()

tab_calc, tab_visor, tab_historial = st.tabs(["📊 Calculadora Mensual", "📄 Visor Avanzado de CFDI", "📑 Historial y Acuses"])

# --- PESTAÑA 1: CALCULADORA ---
with tab_calc:
    st.subheader("Configuración del Periodo")
    col_cli, col_mes, col_anio = st.columns(3)
    
    with col_cli:
        respuesta_clientes = supabase.table("clientes").select("*").execute()
        lista_clientes = respuesta_clientes.data if respuesta_clientes.data else []
        nombres_clientes = [c['nombre'] for c in lista_clientes]
        
        cliente_sel = st.selectbox("Cliente Actual", nombres_clientes) if nombres_clientes else None
            
        with st.expander("➕ Agregar Nuevo Cliente"):
            with st.form("form_cliente", clear_on_submit=True):
                nuevo_cli = st.text_input("Nombre / Razón Social")
                nuevo_rfc = st.text_input("RFC (Debe ser único)")
                if st.form_submit_button("Registrar Cliente"):
                    if not nuevo_cli.strip() or not nuevo_rfc.strip():
                        st.error("⚠️ El Nombre y el RFC son obligatorios.")
                    else:
                        try:
                            supabase.table("clientes").insert({"nombre": nuevo_cli.strip(), "rfc": nuevo_rfc.strip()}).execute()
                            st.success("✅ Cliente registrado exitosamente.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"⚠️ Error al registrar (Es posible que el RFC ya exista).")
    
    with col_mes:
        mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    with col_anio:
        anio_sel = st.selectbox("Año", [2024, 2025, 2026, 2027])
        
    st.divider()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Ingresos (XML Emitidos)")
        xml_ingresos = st.file_uploader("Sube facturas emitidas", type=['xml'], accept_multiple_files=True, key=f"ing_{st.session_state.reset_key}")
    with col2:
        st.subheader("2. Gastos (XML Recibidos)")
        xml_gastos = st.file_uploader("Sube facturas de gastos", type=['xml'], accept_multiple_files=True, key=f"gas_{st.session_state.reset_key}")

    if xml_ingresos:
        datos_ingresos = [procesar_ingreso_resico(x) for x in xml_ingresos]
        datos_ingresos = [d for d in datos_ingresos if d is not None] 
        
        if not datos_ingresos:
            st.error("⚠️ Ninguno de los XML subidos corresponde al Régimen RESICO (626).")
        else:
            df_ingresos = pd.DataFrame(datos_ingresos)
            
            # Totales con decimales para el detalle
            total_ingresos = df_ingresos["Subtotal"].sum()
            total_iva_trasladado = df_ingresos["IVA Trasladado"].sum()
            total_isr_retenido = df_ingresos["ISR Retenido"].sum()
            total_iva_retenido = df_ingresos["IVA Retenido"].sum()
            
            total_iva_acreditable = 0.0
            df_gastos = None
            df_excluidos = None
            
            if xml_gastos:
                datos_gastos = [procesar_gasto_inteligente(x) for x in xml_gastos]
                datos_gastos = [d for d in datos_gastos if d is not None]
                if datos_gastos:
                    df_todos_gastos = pd.DataFrame(datos_gastos)
                    df_gastos = df_todos_gastos[df_todos_gastos["Estado"] == "Acreditable"]
                    df_excluidos = df_todos_gastos[df_todos_gastos["Estado"] == "Excluido"]
                    total_iva_acreditable = df_gastos["IVA Acreditable"].sum()
            
            # --- REDONDEO FISCAL PARA EL SAT (Cifras enteras) ---
            ingresos_redondeados = round(total_ingresos)
            isr_determinado, tasa_aplicada = calcular_isr_resico(ingresos_redondeados)
            isr_determinado = round(isr_determinado)
            isr_retenido_redondeado = round(total_isr_retenido)
            isr_a_pagar = max(0, isr_determinado - isr_retenido_redondeado)
            
            iva_trasladado_redondeado = round(total_iva_trasladado)
            iva_acreditable_redondeado = round(total_iva_acreditable)
            iva_retenido_redondeado = round(total_iva_retenido)
            iva_neto = iva_trasladado_redondeado - iva_acreditable_redondeado - iva_retenido_redondeado
            estatus_iva = "Cargo" if iva_neto > 0 else "Favor"
            
            st.divider()
            st.header(f"📊 Resumen del Cálculo: {mes_sel} {anio_sel}")
            
            if df_excluidos is not None and not df_excluidos.empty:
                st.warning(f"🛡️ **Filtro Inteligente RESICO:** Se excluyeron **{len(df_excluidos)} factura(s)** (Régimen 605 o personales).")

            st.subheader("Determinación de ISR (Redondeado oficial SAT)")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Ingresos Base", f"${ingresos_redondeados:,.0f}")
            c2.metric(f"ISR Determinado ({(tasa_aplicada*100):.2f}%)", f"${isr_determinado:,.0f}")
            c3.metric("ISR Retenido", f"- ${isr_retenido_redondeado:,.0f}")
            c4.metric("ISR a Pagar", f"${isr_a_pagar:,.0f}", delta="Pago requerido", delta_color="inverse")
            
            st.subheader("Determinación de IVA (Redondeado oficial SAT)")
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("IVA Trasladado", f"${iva_trasladado_redondeado:,.0f}")
            c6.metric("IVA Acreditable", f"- ${iva_acreditable_redondeado:,.0f}")
            c7.metric("IVA Retenido", f"- ${iva_retenido_redondeado:,.0f}")
            
            color_iva = "normal" if iva_neto < 0 else "inverse"
            c8.metric(f"IVA a {estatus_iva}", f"${abs(iva_neto):,.0f}", delta=f"Saldo a {estatus_iva}", delta_color=color_iva)

            with st.expander("🔎 Ver detalle completo con decimales exactos"):
                st.write(f"• Ingresos exactos: ${total_ingresos:,.2f}")
                st.write(f"• IVA trasladado exacto: ${total_iva_trasladado:,.2f}")
                st.write(f"• IVA acreditable exacto: ${total_iva_acreditable:,.2f}")

            st.divider()
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if cliente_sel and st.button("💾 Guardar Cálculo en Base de Datos", use_container_width=True):
                    cliente_id = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_sel)
                    viejo = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id).eq("mes", mes_sel).eq("anio", anio_sel).execute()
                    if viejo.data:
                        supabase.table("historial_calculos").delete().eq("id", viejo.data[0]["id"]).execute()
                    
                    datos_insertar = {
                        "cliente_id": cliente_id, "mes": mes_sel, "anio": anio_sel,
                        "ingresos_base": float(ingresos_redondeados), "isr_determinado": float(isr_a_pagar),
                        "iva_cargo_favor": float(iva_neto), "estatus": "Guardado / Pendiente de Acuse"
                    }
                    try:
                        supabase.table("historial_calculos").insert(datos_insertar).execute()
                        st.success("✅ Cálculo guardado con éxito. Ya puedes consultarlo en la pestaña 'Historial y Acuses'.")
                    except Exception as e:
                        st.error(f"Error al guardar: {e}")

            with col_btn2:
                diccionario_totales = {
                    'ingresos': ingresos_redondeados, 'tasa': tasa_aplicada, 'isr_determinado': isr_determinado,
                    'isr_retenido': isr_retenido_redondeado, 'isr_a_pagar': isr_a_pagar, 'iva_trasladado': iva_trasladado_redondeado,
                    'iva_acreditable': iva_acreditable_redondeado, 'iva_retenido': iva_retenido_redondeado,
                    'iva_neto': abs(iva_neto), 'estatus_iva': estatus_iva
                }
                archivo_excel = generar_excel_formulado(df_ingresos, df_gastos, df_excluidos, diccionario_totales)
                st.download_button("📥 Descargar Papeles de Trabajo Formulados (.xlsx)", data=archivo_excel, file_name=f"Papel_Trabajo_{mes_sel}_{anio_sel}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- PESTAÑA 2: VISOR DE XML ---
with tab_visor:
    st.header("🔍 Visor de XML con Filtros Inteligentes")
    tipo_visor = st.radio("¿Qué facturas deseas consultar?", ["Emitidas (Ingresos)", "Recibidas (Gastos)"], horizontal=True)
    archivos_to_show = xml_ingresos if tipo_visor == "Emitidas (Ingresos)" else xml_gastos
    
    if archivos_to_show:
        metadatos = [obtener_metadatos_basicos(f, 'ingreso' if tipo_visor.startswith("Emitidas") else 'gasto') for f in archivos_to_show]
        df_meta = pd.DataFrame([m for m in metadatos if m])
        if not df_meta.empty:
            seleccion = st.selectbox("Selecciona CFDI:", df_meta['display'].tolist())
            archivo_activo = next(f for f in archivos_to_show if f.name == df_meta[df_meta['display'] == seleccion]['nombre_archivo'].iloc[0])
            st.json(extraer_cfdi_completo(archivo_activo))
    else:
        st.info("Sube XMLs en la pestaña 1.")

# --- PESTAÑA 3: HISTORIAL Y ACUSES (CORREGIDA) ---
with tab_historial:
    st.header("🗂️ Historial Mensual y Verificación de Acuses")
    
    if lista_clientes:
        cliente_hist = st.selectbox("Selecciona Cliente para ver su Historial", nombres_clientes, key="cli_hist")
        cliente_id_hist = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_hist)
        
        try:
            historial_db = supabase.table("historial_calculos").select("*").eq("cliente_id", cliente_id_hist).execute()
            
            if historial_db.data:
                df_historial = pd.DataFrame(historial_db.data)
                st.dataframe(df_historial[["mes", "anio", "estatus", "ingresos_base", "isr_determinado", "iva_cargo_favor"]], use_container_width=True)
                
                st.divider()
                st.subheader("Subir Acuse de Declaración (PDF)")
                mes_acuse = st.selectbox("Mes a verificar", df_historial["mes"].unique(), key="mes_ac")
                anio_acuse = st.selectbox("Año a verificar", df_historial["anio"].unique(), key="anio_ac")
                
                acuse_pdf = st.file_uploader("Sube el acuse del SAT en PDF", type=['pdf'], key=f"pdf_{st.session_state.reset_key}")
                
                if acuse_pdf:
                    with st.spinner("Leyendo acuse..."):
                        texto_pdf = ""
                        with pdfplumber.open(acuse_pdf) as pdf:
                            for p in pdf.pages: texto_pdf += p.extract_text() + "\n"
                        cantidades = re.findall(r"\$\s*([\d,]+\.\d{2})", texto_pdf)
                        if cantidades:
                            st.success(f"Monto detectado en acuse: ${cantidades[-1]}")
                            if st.button("✅ Cuadra perfecto. Marcar como Declarado"):
                                registro = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id_hist).eq("mes", mes_acuse).eq("anio", anio_acuse).execute()
                                if registro.data:
                                    supabase.table("historial_calculos").update({"estatus": "Declarado y Cuadrado"}).eq("id", registro.data[0]["id"]).execute()
                                    st.balloons()
                                    st.success("¡Estatus actualizado a Declarado!")
                        else:
                            st.warning("No se pudo extraer el monto exacto, revisa el texto:")
                            st.text(texto_pdf)
            else:
                st.info("No hay cálculos guardados para este cliente todavía. Realiza un cálculo en la Pestaña 1 y guárdalo.")
        except Exception as e:
            st.error(f"Error al conectar con la base de datos de historial: {e}")
    else:
        st.info("Registra un cliente primero.")
