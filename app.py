import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import pdfplumber
import re
from supabase import create_client, Client

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
# 1. FUNCIONES DE PROCESAMIENTO (CALCULADORA)
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

def procesar_gasto(archivo_subido):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        
        subtotal = float(root.attrib.get('SubTotal', 0.0))
        iva_acreditable = 0.0
        impuestos = root.find('cfdi:Impuestos', ns)
        if impuestos is not None:
            traslados = impuestos.find('cfdi:Traslados', ns)
            if traslados is not None:
                for t in traslados.findall('cfdi:Traslado', ns):
                    if t.attrib.get('Impuesto') == '002': iva_acreditable += float(t.attrib.get('Importe', 0.0))
        return {"Archivo": archivo_subido.name, "Subtotal Gasto": subtotal, "IVA Acreditable": iva_acreditable}
    except:
        return None

def calcular_isr_resico(ingresos_totales):
    if ingresos_totales <= 25000.00: tasa = 0.01
    elif ingresos_totales <= 50000.00: tasa = 0.011
    elif ingresos_totales <= 83333.33: tasa = 0.015
    elif ingresos_totales <= 208333.33: tasa = 0.02
    else: tasa = 0.025
    return ingresos_totales * tasa, tasa

def generar_excel(df_ingresos, df_gastos, totales):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_resumen = pd.DataFrame({
            "Concepto": ["Ingresos Base", "Tasa ISR (%)", "ISR Determinado", "ISR Retenido", "ISR a Pagar", "---",
                         "IVA Trasladado", "IVA Acreditable", "IVA Retenido", f"IVA a {totales['estatus_iva']}"],
            "Importe": [totales['ingresos'], totales['tasa'] * 100, totales['isr_determinado'], totales['isr_retenido'], 
                        totales['isr_a_pagar'], 0.0, totales['iva_trasladado'], totales['iva_acreditable'], 
                        totales['iva_retenido'], totales['iva_neto']]
        })
        df_resumen.to_excel(writer, sheet_name='Resumen Fiscal', index=False)
        if not df_ingresos.empty: df_ingresos.to_excel(writer, sheet_name='Detalle Ingresos', index=False)
        if df_gastos is not None and not df_gastos.empty: df_gastos.to_excel(writer, sheet_name='Detalle Gastos', index=False)
    return buffer.getvalue()

# ==========================================
# 2. FUNCIONES DE EXTRACCIÓN (VISOR DE XML)
# ==========================================
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
        if tfd is not None:
            datos["uuid"] = tfd.attrib.get('UUID', '')
            
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
                if t.attrib.get('Impuesto') == '002':
                    datos["iva_trasladado"] += float(t.attrib.get('Importe', '0.00'))
                    
        retenciones = impuestos.find('cfdi:Retenciones', ns)
        if retenciones is not None:
            for r in retenciones.findall('cfdi:Retencion', ns):
                if r.attrib.get('Impuesto') == '001':
                    datos["isr_retenido"] += float(r.attrib.get('Importe', '0.00'))
                elif r.attrib.get('Impuesto') == '002':
                    datos["iva_retenido"] += float(r.attrib.get('Importe', '0.00'))
                    
    return datos

# ==========================================
# 3. INTERFAZ WEB PRINCIPAL
# ==========================================
st.set_page_config(page_title="App RESICO", layout="wide")

if supabase is None:
    st.error("⚠️ No se pudo conectar a la base de datos. Revisa que tu archivo .streamlit/secrets.toml esté bien configurado.")
    st.stop()

st.title("🧮 Sistema Contable Automatizado RESICO")

tab_calc, tab_visor, tab_historial = st.tabs(["📊 Calculadora Mensual", "📄 Visor de CFDI", "📑 Historial y Acuses"])

# --- PESTAÑA 1: CALCULADORA ---
with tab_calc:
    # 1. Configuración del mes a calcular
    st.subheader("Configuración del Periodo")
    col_cli, col_mes, col_anio = st.columns(3)
    
    with col_cli:
        respuesta_clientes = supabase.table("clientes").select("*").execute()
        lista_clientes = respuesta_clientes.data if respuesta_clientes.data else []
        nombres_clientes = [c['nombre'] for c in lista_clientes]
        cliente_sel = st.selectbox("Cliente", nombres_clientes) if nombres_clientes else st.selectbox("Cliente", ["Sin clientes"])
    
    with col_mes:
        mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    with col_anio:
        anio_sel = st.selectbox("Año", [2024, 2025, 2026, 2027])
        
    st.divider()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Ingresos (XML Emitidos)")
        xml_ingresos = st.file_uploader("Sube facturas emitidas", type=['xml'], accept_multiple_files=True, key="ing")
    with col2:
        st.subheader("2. Gastos (XML Recibidos)")
        xml_gastos = st.file_uploader("Sube facturas de gastos para el IVA", type=['xml'], accept_multiple_files=True, key="gas")

    if xml_ingresos:
        datos_ingresos = [procesar_ingreso_resico(x) for x in xml_ingresos]
        datos_ingresos = [d for d in datos_ingresos if d is not None] 
        
        if datos_ingresos:
            df_ingresos = pd.DataFrame(datos_ingresos)
            total_ingresos = df_ingresos["Subtotal"].sum()
            total_iva_trasladado = df_ingresos["IVA Trasladado"].sum()
            total_isr_retenido = df_ingresos["ISR Retenido"].sum()
            total_iva_retenido = df_ingresos["IVA Retenido"].sum()
            
            total_iva_acreditable = 0.0
            df_gastos = None
            
            if xml_gastos:
                datos_gastos = [procesar_gasto(x) for x in xml_gastos]
                datos_gastos = [d for d in datos_gastos if d is not None]
                if datos_gastos:
                    df_gastos = pd.DataFrame(datos_gastos)
                    total_iva_acreditable = df_gastos["IVA Acreditable"].sum()
            
            isr_determinado, tasa_aplicada = calcular_isr_resico(total_ingresos)
            isr_a_pagar = isr_determinado - total_isr_retenido
            if isr_a_pagar < 0: isr_a_pagar = 0.0
            iva_neto = total_iva_trasladado - total_iva_acreditable - total_iva_retenido
            estatus_iva = "Cargo" if iva_neto > 0 else "Favor"
            
            st.divider()
            st.header(f"📊 Resumen del Cálculo: {mes_sel} {anio_sel}")
            
            st.subheader("Determinación de ISR")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Ingresos Base", f"${total_ingresos:,.2f}")
            c2.metric(f"ISR Determinado ({(tasa_aplicada*100):.2f}%)", f"${isr_determinado:,.2f}")
            c3.metric("ISR Retenido", f"- ${total_isr_retenido:,.2f}")
            c4.metric("ISR a Pagar", f"${isr_a_pagar:,.2f}", delta="Pago requerido", delta_color="inverse")
            
            st.subheader("Determinación de IVA")
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("IVA Trasladado (Cobrado)", f"${total_iva_trasladado:,.2f}")
            c6.metric("IVA Acreditable (Gastos)", f"- ${total_iva_acreditable:,.2f}")
            c7.metric("IVA Retenido", f"- ${total_iva_retenido:,.2f}")
            
            color_iva = "normal" if iva_neto < 0 else "inverse"
            c8.metric(f"IVA a {estatus_iva}", f"${abs(iva_neto):,.2f}", delta=f"Saldo a {estatus_iva}", delta_color=color_iva)

            # Botones de Acción Final
            st.divider()
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if st.button("💾 Guardar Cálculo en Base de Datos", use_container_width=True):
                    if lista_clientes:
                        cliente_id = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_sel)
                        
                        # Borrar si ya existía el mismo mes para sobreescribir
                        viejo = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id).eq("mes", mes_sel).eq("anio", anio_sel).execute()
                        if viejo.data:
                            supabase.table("historial_calculos").delete().eq("id", viejo.data[0]["id"]).execute()
                        
                        # Insertar nuevo cálculo
                        datos_insertar = {
                            "cliente_id": cliente_id,
                            "mes": mes_sel,
                            "anio": anio_sel,
                            "ingresos_base": float(total_ingresos),
                            "isr_determinado": float(isr_a_pagar),
                            "iva_cargo_favor": float(iva_neto),
                            "estatus": "Calculando"
                        }
                        supabase.table("historial_calculos").insert(datos_insertar).execute()
                        st.success(f"Cálculo de {mes_sel} {anio_sel} guardado exitosamente para {cliente_sel}.")

            with col_btn2:
                diccionario_totales = {
                    'ingresos': total_ingresos, 'tasa': tasa_aplicada, 'isr_determinado': isr_determinado,
                    'isr_retenido': total_isr_retenido, 'isr_a_pagar': isr_a_pagar, 'iva_trasladado': total_iva_trasladado,
                    'iva_acreditable': total_iva_acreditable, 'iva_retenido': total_iva_retenido,
                    'iva_neto': abs(iva_neto), 'estatus_iva': estatus_iva
                }
                archivo_excel = generar_excel(df_ingresos, df_gastos, diccionario_totales)
                
                st.download_button(
                    label="📥 Descargar Papeles de Trabajo (.xlsx)",
                    data=archivo_excel,
                    file_name=f"Papel_Trabajo_{mes_sel}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# --- PESTAÑA 2: VISOR DE XML ---
with tab_visor:
    st.header("Visor de XML (Estilo Factura)")
    
    todos_los_archivos = []
    if 'xml_ingresos' in locals() and xml_ingresos: todos_los_archivos.extend(xml_ingresos)
    if 'xml_gastos' in locals() and xml_gastos: todos_los_archivos.extend(xml_gastos)
    
    if todos_los_archivos:
        nombres = [f.name for f in todos_los_archivos]
        seleccion = st.selectbox("Selecciona un archivo XML para visualizar:", nombres)
        
        archivo_activo = next(f for f in todos_los_archivos if f.name == seleccion)
        datos_visor = extraer_cfdi_completo(archivo_activo)
        
        html_cfdi = f"""<div style="font-family: Arial, sans-serif; border: 1px solid #ddd; border-radius: 8px; padding: 25px; background-color: #ffffff; color: #333;">
<div style="display: flex; justify-content: space-between; border-bottom: 2px solid #2e86c1; padding-bottom: 10px; margin-bottom: 20px;">
<div>
<h2 style="margin: 0; color: #2e86c1;">Factura Electrónica (CFDI)</h2>
<p style="margin: 5px 0 0 0;"><strong>UUID:</strong> {datos_visor['uuid']}</p>
</div>
<div style="text-align: right;">
<h3 style="margin: 0; color: #555;">Serie y Folio: {datos_visor['serie']}{datos_visor['folio']}</h3>
<p style="margin: 5px 0 0 0;">Fecha: {datos_visor['fecha']}</p>
</div>
</div>
<div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
<div style="width: 48%; padding: 15px; background-color: #f9f9f9; border-radius: 5px;">
<h4 style="margin-top: 0; border-bottom: 1px solid #ccc; padding-bottom: 5px;">Datos del Emisor</h4>
<p style="margin: 5px 0;"><strong>Nombre:</strong> {datos_visor['emi_nombre']}</p>
<p style="margin: 5px 0;"><strong>RFC:</strong> {datos_visor['emi_rfc']}</p>
<p style="margin: 5px 0;"><strong>Régimen:</strong> {datos_visor['emi_regimen']}</p>
</div>
<div style="width: 48%; padding: 15px; background-color: #f9f9f9; border-radius: 5px;">
<h4 style="margin-top: 0; border-bottom: 1px solid #ccc; padding-bottom: 5px;">Datos del Receptor</h4>
<p style="margin: 5px 0;"><strong>Nombre:</strong> {datos_visor['rec_nombre']}</p>
<p style="margin: 5px 0;"><strong>RFC:</strong> {datos_visor['rec_rfc']}</p>
<p style="margin: 5px 0;"><strong>Uso CFDI:</strong> {datos_visor['rec_uso']}</p>
</div>
</div>
<h4 style="margin-bottom: 10px; border-bottom: 1px solid #ccc; padding-bottom: 5px;">Conceptos</h4>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
<tr style="background-color: #2e86c1; color: white;">
<th style="padding: 8px; text-align: center; border: 1px solid #ddd;">Cant.</th>
<th style="padding: 8px; text-align: left; border: 1px solid #ddd;">Descripción</th>
<th style="padding: 8px; text-align: right; border: 1px solid #ddd;">Valor Unitario</th>
<th style="padding: 8px; text-align: right; border: 1px solid #ddd;">Importe</th>
</tr>"""
        for c in datos_visor['conceptos']:
            html_cfdi += f"""<tr>
<td style="padding: 8px; text-align: center; border: 1px solid #ddd;">{c['cantidad']}</td>
<td style="padding: 8px; text-align: left; border: 1px solid #ddd;">{c['descripcion']}</td>
<td style="padding: 8px; text-align: right; border: 1px solid #ddd;">${c['valor_unitario']:,.2f}</td>
<td style="padding: 8px; text-align: right; border: 1px solid #ddd;">${c['importe']:,.2f}</td>
</tr>"""
        html_cfdi += f"""</table>
<div style="display: flex; justify-content: flex-end;">
<div style="width: 320px; padding: 15px; background-color: #f0f8ff; border-radius: 5px; border: 1px solid #b0c4de;">
<div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
<strong>Subtotal:</strong> <span>${datos_visor['subtotal']:,.2f}</span>
</div>"""

        if datos_visor['descuento'] > 0:
            html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px; color: #d9534f;">
<strong>Descuento:</strong> <span>-${datos_visor['descuento']:,.2f}</span>
</div>"""
        if datos_visor['iva_trasladado'] > 0:
            html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
<strong>IVA (16%):</strong> <span>${datos_visor['iva_trasladado']:,.2f}</span>
</div>"""
        if datos_visor['isr_retenido'] > 0:
            html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px; color: #d9534f;">
<strong>Retención ISR:</strong> <span>-${datos_visor['isr_retenido']:,.2f}</span>
</div>"""
        if datos_visor['iva_retenido'] > 0:
            html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px; color: #d9534f;">
<strong>Retención IVA:</strong> <span>-${datos_visor['iva_retenido']:,.2f}</span>
</div>"""

        html_cfdi += f"""<div style="display: flex; justify-content: space-between; font-size: 1.2em; border-top: 1px solid #ccc; padding-top: 5px; margin-top: 5px;">
<strong>Total:</strong> <strong>${datos_visor['total']:,.2f} {datos_visor['moneda']}</strong>
</div>
</div>
</div>
</div>"""
        st.markdown(html_cfdi, unsafe_allow_html=True)
    else:
        st.info("Sube archivos XML en la pestaña de 'Calculadora Mensual' para poder visualizarlos aquí.")

# --- PESTAÑA 3: HISTORIAL Y ACUSES ---
with tab_historial:
    st.header("🗂️ Historial Mensual y Verificación de Acuses")
    
    if lista_clientes:
        cliente_hist = st.selectbox("Selecciona Cliente", nombres_clientes, key="cli_hist")
        cliente_id_hist = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_hist)
        
        historial_db = supabase.table("historial_calculos").select("*").eq("cliente_id", cliente_id_hist).order("anio", desc=True).order("mes", desc=True).execute()
        
        if historial_db.data:
            df_historial = pd.DataFrame(historial_db.data)
            df_historial = df_historial[["mes", "anio", "estatus", "ingresos_base", "isr_determinado", "iva_cargo_favor"]]
            
            # Formato moneda para visualización
            for col in ["ingresos_base", "isr_determinado", "iva_cargo_favor"]:
                df_historial[col] = df_historial[col].apply(lambda x: f"${x:,.2f}")
                
            st.dataframe(df_historial, use_container_width=True)
            
            st.divider()
            st.subheader("Subir Acuse de Declaración (PDF)")
            mes_acuse = st.selectbox("Mes a verificar", df_historial["mes"].unique())
            anio_acuse = st.selectbox("Año a verificar", df_historial["anio"].unique())
            
            acuse_pdf = st.file_uploader(f"Sube el acuse del SAT para {mes_acuse} {anio_acuse}", type=['pdf'])
            
            if acuse_pdf:
                with st.spinner("Leyendo documento del SAT..."):
                    try:
                        texto_pdf = ""
                        with pdfplumber.open(acuse_pdf) as pdf:
                            for pagina in pdf.pages:
                                texto_pdf += pagina.extract_text() + "\n"
                        
                        # Buscar cantidades de "Total a pagar" o "Cantidad a cargo"
                        cantidades_encontradas = re.findall(r"\$\s*([\d,]+\.\d{2})", texto_pdf)
                        
                        if cantidades_encontradas:
                            st.success(f"Se extrajeron montos del PDF. Último monto detectado: ${cantidades_encontradas[-1]}")
                            st.write("Verifica que cuadre con tu cálculo inicial.")
                            
                            if st.button("✅ Cuadra perfecto. Marcar como Declarado", use_container_width=True):
                                # Actualizar estatus en Supabase
                                registro = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id_hist).eq("mes", mes_acuse).eq("anio", anio_acuse).execute()
                                if registro.data:
                                    supabase.table("historial_calculos").update({"estatus": "Declarado y Cuadrado"}).eq("id", registro.data[0]["id"]).execute()
                                    st.balloons()
                                    st.success("¡Mes cerrado y actualizado en la base de datos! Refresca la página para ver el cambio.")
                        else:
                            st.warning("No se pudo extraer automáticamente una cantidad exacta con formato moneda. Aquí tienes el texto para revisión manual:")
                            with st.expander("Ver texto del acuse"):
                                st.text(texto_pdf)
                    except Exception as e:
                        st.error(f"Error al leer el PDF: {e}")
        else:
            st.info("Aún no has guardado cálculos para este cliente.")
    else:
        st.info("No hay clientes registrados en la base de datos.")
