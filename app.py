import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
import pdfplumber
import re
from supabase import create_client, Client

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

MESES_LISTA = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

def calcular_iva_acumulado_anterior(cliente_id, mes_actual, anio_actual):
    try:
        idx_actual = MESES_LISTA.index(mes_actual)
        res = supabase.table("historial_calculos").select("*").eq("cliente_id", cliente_id).eq("anio", int(anio_actual)).execute()
        if not res.data:
            return 0.0
        
        disponible = 0.0
        for m in MESES_LISTA[:idx_actual]:
            mes_encontrado = next((item for item in res.data if item["mes"] == m), None)
            if mes_encontrado:
                val = float(mes_encontrado.get("iva_cargo_favor", 0.0))
                if val < 0:
                    disponible += abs(val)
                elif val > 0:
                    disponible -= val
                    if disponible < 0:
                        disponible = 0.0
                        
        return disponible
    except:
        return 0.0

# ==========================================
# 1. FUNCIONES DE PROCESAMIENTO
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

def procesar_gasto_o_pago(archivo_subido):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {
            'cfdi': 'http://www.sat.gob.mx/cfd/4',
            'pago20': 'http://www.sat.gob.mx/Pagos20'
        }
        
        tipo_comprobante = root.attrib.get('TipoDeComprobante', '')
        
        if tipo_comprobante == 'P':
            emisor = root.find('cfdi:Emisor', ns)
            nombre_emisor = emisor.attrib.get('Nombre', 'Proveedor REP') if emisor is not None else 'Proveedor REP'
            rfc_emisor = emisor.attrib.get('Rfc', '') if emisor is not None else ''
            
            iva_pagado_rep = 0.0
            uuids_relacionados = []
            
            pagos = root.findall('.//pago20:Pago', ns)
            for pago in pagos:
                doctos = pago.findall('pago20:DoctoRelacionado', ns)
                for doc in doctos:
                    uuid_rel = doc.attrib.get('IdDocumento', '')
                    if uuid_rel:
                        uuids_relacionados.append(uuid_rel.upper())
                
                traslados_p = pago.findall('.//pago20:TrasladoP', ns)
                for t_p in traslados_p:
                    if t_p.attrib.get('ImpuestoP') == '002':
                        iva_pagado_rep += float(t_p.attrib.get('ImporteP', 0.0))

            return {
                "TipoRegistro": "REP",
                "Archivo": archivo_subido.name,
                "Proveedor": nombre_emisor,
                "RFC Emisor": rfc_emisor,
                "UUIDs Relacionados": uuids_relacionados,
                "IVA Acreditable": iva_pagado_rep,
                "Subtotal Gasto": 0.0,
                "Estado": "REP Válido",
                "Motivo Rechazo": "Complemento de Pago que libera IVA de PPD"
            }

        emisor = root.find('cfdi:Emisor', ns)
        rfc_emisor = emisor.attrib.get('Rfc', '') if emisor is not None else ''
        nombre_emisor = emisor.attrib.get('Nombre', 'Desconocido') if emisor is not None else 'Desconocido'
        regimen_emisor = emisor.attrib.get('RegimenFiscal', '') if emisor is not None else ''

        receptor = root.find('cfdi:Receptor', ns)
        regimen_receptor = receptor.attrib.get('RegimenFiscalReceptor', '') if receptor is not None else ''
        uso_cfdi = receptor.attrib.get('UsoCFDI', '') if receptor is not None else ''
        
        metodo_pago = root.attrib.get('MetodoPago', '')
        
        tfd = root.find('.//tfd:TimbreFiscalDigital', {'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'})
        uuid_factura = tfd.attrib.get('UUID', '').upper() if tfd is not None else ''

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

        if regimen_receptor != '626':
            estado = "Excluido"
            motivo = f"Régimen fiscal del receptor ({regimen_receptor}) no es RESICO (626)"
        elif metodo_pago == 'PPD':
            estado = "PPD Pendiente"
            motivo = "Método PPD (Esperando Complemento de Pago / REP)"
        elif uso_cfdi in ['D01', 'D02', 'D03', 'D04', 'D05', 'D06', 'D07', 'D08', 'D09', 'D10']:
            estado = "Excluido"
            motivo = f"Uso de CFDI ({uso_cfdi}) es deducción personal"

        return {
            "TipoRegistro": "Factura",
            "UUID": uuid_factura,
            "Archivo": archivo_subido.name,
            "Proveedor": nombre_emisor,
            "RFC Emisor": rfc_emisor,
            "Régimen": regimen_emisor,
            "Regimen Receptor": regimen_receptor,
            "Método Pago": metodo_pago,
            "Uso CFDI": uso_cfdi,
            "Subtotal Gasto": subtotal,
            "IVA Acreditable": iva_acreditable,
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
            ["IVA a Favor Acumulado", totales['iva_favor_anterior']],
            [f"IVA a {totales['estatus_iva']}", "=B8-B9-B10-B11"]
        ]
        df_resumen = pd.DataFrame(wb_resumen_data[1:], columns=wb_resumen_data[0])
        df_resumen.to_excel(writer, sheet_name='Resumen Fiscal', index=False)
        
        if df_ingresos is not None and not df_ingresos.empty: df_ingresos.to_excel(writer, sheet_name='Detalle Ingresos', index=False)
        if df_gastos is not None and not df_gastos.empty: df_gastos.to_excel(writer, sheet_name='Gastos Acreditables', index=False)
        if df_excluidos is not None and not df_excluidos.empty: df_excluidos.to_excel(writer, sheet_name='Gastos Excluidos', index=False)
    return buffer.getvalue()

def obtener_metadatos_basicos(archivo_subido, tipo_archivo):
    try:
        archivo_subido.seek(0)
        tree = ET.parse(archivo_subido)
        root = tree.getroot()
        ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
        
        fecha_completa = root.attrib.get('Fecha', '')
        fecha = fecha_completa[:10]
        tipo = root.attrib.get('TipoDeComprobante', '')
        metodo = root.attrib.get('MetodoPago', 'N/A')
        
        total = 0.0
        if tipo == 'P':
            metodo = 'Pago (Complemento)'
            pagos = root.findall('.//{http://www.sat.gob.mx/Pagos20}Pago')
            for p in pagos:
                total += float(p.attrib.get('Monto', 0.0))
        else:
            total = float(root.attrib.get('Total', '0.00'))
            
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
            "total": total,
            "metodo": metodo,
            "display": f"{fecha} | {nombre} | ${total:,.2f} | {metodo}"
        }
    except:
        return None

def extraer_cfdi_completo(archivo_subido):
    archivo_subido.seek(0)
    tree = ET.parse(archivo_subido)
    root = tree.getroot()
    ns = {'cfdi': 'http://www.sat.gob.mx/cfd/4', 'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital', 'pago20': 'http://www.sat.gob.mx/Pagos20'}
    
    tipo_comp = root.attrib.get('TipoDeComprobante', '')
    fecha_hora = root.attrib.get('Fecha', '')
    
    fecha = fecha_hora[:10] if len(fecha_hora) >= 10 else ''
    hora = fecha_hora[11:19] if len(fecha_hora) >= 19 else '00:00:00'
    
    datos = {
        "serie": root.attrib.get('Serie', ''),
        "folio": root.attrib.get('Folio', 'S/F'),
        "fecha": fecha,
        "hora": hora,
        "subtotal": float(root.attrib.get('SubTotal', '0.00')),
        "descuento": float(root.attrib.get('Descuento', '0.00')) if root.attrib.get('Descuento') else 0.0,
        "total": float(root.attrib.get('Total', '0.00')),
        "moneda": root.attrib.get('Moneda', 'MXN'),
        "uuid": "No encontrado",
        "conceptos": [],
        "iva_trasladado": 0.0,
        "iva_retenido": 0.0,
        "isr_retenido": 0.0
    }
    
    if tipo_comp == 'P':
        monto_total_pago = 0.0
        pagos = root.findall('.//pago20:Pago', ns)
        for p in pagos:
            monto_total_pago += float(p.attrib.get('Monto', 0.0))
            
        datos["subtotal"] = monto_total_pago
        datos["total"] = monto_total_pago
        datos["conceptos"].append({
            "cantidad": "1",
            "descripcion": "Complemento de Recepción de Pagos (REP)",
            "valor_unitario": monto_total_pago,
            "importe": monto_total_pago
        })
    
    emisor = root.find('cfdi:Emisor', ns)
    datos["emi_rfc"] = emisor.attrib.get('Rfc', '') if emisor is not None else ''
    datos["emi_nombre"] = emisor.attrib.get('Nombre', '') if emisor is not None else ''
    datos["emi_regimen"] = emisor.attrib.get('RegimenFiscal', '') if emisor is not None else ''
    
    receptor = root.find('cfdi:Receptor', ns)
    datos["rec_rfc"] = receptor.attrib.get('Rfc', '') if receptor is not None else ''
    datos["rec_nombre"] = receptor.attrib.get('Nombre', '') if receptor is not None else ''
    datos["rec_regimen"] = receptor.attrib.get('RegimenFiscalReceptor', '') if receptor is not None else ''
    datos["rec_uso"] = receptor.attrib.get('UsoCFDI', '') if receptor is not None else ''
    
    complemento = root.find('cfdi:Complemento', ns)
    if complemento is not None:
        tfd = complemento.find('tfd:TimbreFiscalDigital', ns)
        if tfd is not None: datos["uuid"] = tfd.attrib.get('UUID', '')
            
    conceptos = root.find('cfdi:Conceptos', ns)
    if conceptos is not None and tipo_comp != 'P':
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
                            existe = supabase.table("clientes").select("id").eq("rfc", nuevo_rfc.strip()).execute()
                            if existe.data:
                                st.warning(f"⚠️ El RFC '{nuevo_rfc.strip()}' ya está registrado.")
                            else:
                                supabase.table("clientes").insert({"nombre": nuevo_cli.strip(), "rfc": nuevo_rfc.strip()}).execute()
                                st.success("✅ Cliente registrado con éxito. Recargando...")
                                st.rerun()
                        except Exception as e:
                            st.error(f"⚠️ Error al registrar: {e}")
        
        if lista_clientes:
            with st.expander("🗑️ Eliminar Cliente (Duplicado o Inactivo)"):
                with st.form("form_eliminar_cliente"):
                    cliente_a_borrar = st.selectbox("Selecciona cliente a eliminar", nombres_clientes, key="del_cli")
                    if st.form_submit_button("🗑️ Borrar Cliente Definitivamente", use_container_width=True):
                        cli_id_borrar = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_a_borrar)
                        try:
                            supabase.table("historial_calculos").delete().eq("cliente_id", cli_id_borrar).execute()
                            supabase.table("clientes").delete().eq("id", cli_id_borrar).execute()
                            st.success(f"🗑️ Cliente '{cliente_a_borrar}' eliminado correctamente.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al eliminar: {e}")
    
    with col_mes:
        mes_sel = st.selectbox("Mes", MESES_LISTA)
    with col_anio:
        anio_sel = st.selectbox("Año", [2024, 2025, 2026, 2027])
        
    st.divider()
    
    iva_favor_acumulado = 0.0
    if cliente_sel and lista_clientes:
        cli_obj = next((c for c in lista_clientes if c['nombre'] == cliente_sel), None)
        if cli_obj:
            iva_favor_acumulado = calcular_iva_acumulado_anterior(cli_obj['id'], mes_sel, anio_sel)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Ingresos (XML Emitidos)")
        xml_ingresos = st.file_uploader("Sube facturas emitidas", type=['xml'], accept_multiple_files=True, key=f"ing_{st.session_state.reset_key}")
    with col2:
        st.subheader("2. Gastos y Pagos (XML Recibidos + REPs)", help="Sube tus facturas (PUE/PPD) y sus complementos de pago tipo P")
        xml_gastos = st.file_uploader("Sube facturas y complementos", type=['xml'], accept_multiple_files=True, key=f"gas_{st.session_state.reset_key}")

    if xml_ingresos or xml_gastos:
        total_ingresos = 0.0
        total_iva_trasladado = 0.0
        total_isr_retenido = 0.0
        total_iva_retenido = 0.0
        df_ingresos = pd.DataFrame()

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
        df_gastos = pd.DataFrame()
        df_excluidos = pd.DataFrame()
        
        if xml_gastos:
            resultados_brutos = [procesar_gasto_o_pago(x) for x in xml_gastos]
            resultados_brutos = [r for r in resultados_brutos if r is not None]
            
            if resultados_brutos:
                df_todos = pd.DataFrame(resultados_brutos)
                
                df_facturas = df_todos[df_todos["TipoRegistro"] == "Factura"].copy() if "TipoRegistro" in df_todos.columns else pd.DataFrame()
                df_reps = df_todos[df_todos["TipoRegistro"] == "REP"].copy() if "TipoRegistro" in df_todos.columns else pd.DataFrame()
                
                uuid_acreditados_por_rep = []
                if not df_reps.empty:
                    for _, rep in df_reps.iterrows():
                        for uuid_rel in rep["UUIDs Relacionados"]:
                            uuid_acreditados_por_rep.append(uuid_rel)
                
                if uuid_acreditados_por_rep and not df_facturas.empty:
                    mask_ppd_pagado = df_facturas["UUID"].isin(uuid_acreditados_por_rep) & (df_facturas["Estado"] == "PPD Pendiente")
                    df_facturas.loc[mask_ppd_pagado, "Estado"] = "Acreditable"
                    df_facturas.loc[mask_ppd_pagado, "Motivo Rechazo"] = "PPD Liberado mediante Complemento de Pago (REP)"

                iva_extra_reps = df_reps["IVA Acreditable"].sum() if not df_reps.empty else 0.0
                df_gastos = df_facturas[df_facturas["Estado"] == "Acreditable"].copy() if not df_facturas.empty else pd.DataFrame()
                
                df_excluidos = pd.concat([
                    df_facturas[df_facturas["Estado"] != "Acreditable"] if not df_facturas.empty else pd.DataFrame(),
                    df_reps
                ], ignore_index=True) if not df_reps.empty or not df_facturas.empty else pd.DataFrame()
                
                total_iva_acreditable = (df_gastos["IVA Acreditable"].sum() if not df_gastos.empty else 0.0) + iva_extra_reps
                
                excluir_manual = []
                if not df_gastos.empty:
                    with st.expander("⚙️ Exclusión Manual Adicional de Gastos (Opcional)"):
                        opciones_manuales = df_gastos["Archivo"].tolist()
                        excluir_manual = st.multiselect(
                            "Selecciona facturas que deseas omitir manualmente de la acreditación:",
                            options=opciones_manuales,
                            format_func=lambda x: f"{x} - {df_gastos[df_gastos['Archivo']==x]['Proveedor'].values[0]} (${df_gastos[df_gastos['Archivo']==x]['Subtotal Gasto'].values[0]:,.2f})"
                        )
                
                if excluir_manual:
                    df_gastos.loc[df_gastos["Archivo"].isin(excluir_manual), "Estado"] = "Excluido Manual"
                    df_gastos.loc[df_gastos["Archivo"].isin(excluir_manual), "Motivo Rechazo"] = "Exclusión manual por el contador"
                    
                    df_excluidos = pd.concat([df_excluidos, df_gastos[df_gastos["Archivo"].isin(excluir_manual)]], ignore_index=True)
                    df_gastos = df_gastos[~df_gastos["Archivo"].isin(excluir_manual)]
                    total_iva_acreditable = df_gastos["IVA Acreditable"].sum() + iva_extra_reps

        ingresos_redondeados = round(total_ingresos)
        isr_determinado, tasa_aplicada = calcular_isr_resico(ingresos_redondeados)
        isr_determinado = round(isr_determinado)
        isr_retenido_redondeado = round(total_isr_retenido)
        isr_a_pagar = max(0, isr_determinado - isr_retenido_redondeado)
        
        iva_trasladado_redondeado = round(total_iva_trasladado)
        iva_acreditable_redondeado = round(total_iva_acreditable)
        iva_retenido_redondeado = round(total_iva_retenido)
        
        iva_neto_bruto = iva_trasladado_redondeado - iva_acreditable_redondeado - iva_retenido_redondeado
        iva_neto = iva_neto_bruto - round(iva_favor_acumulado)
        estatus_iva = "Cargo" if iva_neto > 0 else "Favor"
        
        st.divider()
        st.header(f"📊 Resumen del Cálculo: {mes_sel} {anio_sel}")
        
        if iva_favor_acumulado > 0:
            st.info(f"💡 **Acumulado Histórico:** Se aplicó un IVA a favor acumulado de periodos anteriores por **${iva_favor_acumulado:,.0f}**.")

        if not df_excluidos.empty:
            st.warning(f"🛡️ **Filtro Inteligente RESICO:** Se procesaron complementos de pago y se filtraron registros no acreditables.")

        st.subheader("Determinación de ISR (Redondeado oficial SAT)")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ingresos Base", f"${ingresos_redondeados:,.0f}")
        c2.metric(f"ISR Determinado ({(tasa_aplicada*100):.2f}%)", f"${isr_determinado:,.0f}")
        c3.metric("ISR Retenido", f"- ${isr_retenido_redondeado:,.0f}")
        c4.metric("ISR a Pagar", f"${isr_a_pagar:,.0f}", delta="Pago requerido", delta_color="inverse")
        
        st.subheader("Determinación de IVA (Redondeado oficial SAT)")
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("IVA Trasladado", f"${iva_trasladado_redondeado:,.0f}")
        c6.metric("IVA Acreditable (Incluye PPD liberados por REP)", f"- ${iva_acreditable_redondeado:,.0f}")
        c7.metric("IVA Retenido", f"- ${iva_retenido_redondeado:,.0f}")
        
        color_iva = "normal" if iva_neto < 0 else "inverse"
        c8.metric(f"IVA a {estatus_iva}", f"${abs(iva_neto):,.0f}", delta=f"Saldo a {estatus_iva}", delta_color=color_iva)

        with st.expander("🔎 Ver detalle completo con decimales exactos"):
            st.write(f"• Ingresos exactos: ${total_ingresos:,.2f}")
            st.write(f"• IVA trasladado exacto: ${total_iva_trasladado:,.2f}")
            st.write(f"• IVA acreditable exacto (PUE + REP liberados): ${total_iva_acreditable:,.2f}")
            st.write(f"• IVA a favor acumulado de periodos anteriores: ${iva_favor_acumulado:,.2f}")

        if not df_excluidos.empty:
            with st.expander("🔍 Ver detalle de facturas excluidas / REPs informativos"):
                st.dataframe(df_excluidos[["Archivo", "Proveedor", "Método Pago", "Motivo Rechazo"]], use_container_width=True)

        st.divider()
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if cliente_sel and st.button("💾 Guardar Cálculo en Base de Datos", use_container_width=True):
                cliente_id = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_sel)
                
                viejo = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id).eq("mes", mes_sel).eq("anio", int(anio_sel)).execute()
                if viejo.data:
                    supabase.table("historial_calculos").delete().eq("id", viejo.data[0]["id"]).execute()
                
                datos_insertar = {
                    "cliente_id": cliente_id, 
                    "mes": mes_sel, 
                    "anio": int(anio_sel),
                    "ingresos_base": float(ingresos_redondeados), 
                    "isr_determinado": float(isr_a_pagar),
                    "iva_cargo_favor": float(iva_neto), 
                    "estatus": "Guardado / Pendiente de Acuse"
                }
                try:
                    supabase.table("historial_calculos").insert(datos_insertar).execute()
                    st.success("✅ Cálculo guardado y listo para arrastrar saldo al siguiente mes.")
                except Exception as e:
                    st.error(f"Error al guardar en Supabase: {e}")

        with col_btn2:
            diccionario_totales = {
                'ingresos': ingresos_redondeados, 'tasa': tasa_aplicada, 'isr_determinado': isr_determinado,
                'isr_retenido': isr_retenido_redondeado, 'isr_a_pagar': isr_a_pagar, 'iva_trasladado': iva_trasladado_redondeado,
                'iva_acreditable': iva_acreditable_redondeado, 'iva_retenido': iva_retenido_redondeado,
                'iva_favor_anterior': round(iva_favor_acumulado),
                'iva_neto': abs(iva_neto), 'estatus_iva': estatus_iva
            }
            archivo_excel = generar_excel_formulado(df_ingresos, df_gastos if not df_gastos.empty else None, df_excluidos if not df_excluidos.empty else None, diccionario_totales)
            st.download_button("📥 Descargar Papeles de Trabajo Formulados (.xlsx)", data=archivo_excel, file_name=f"Papel_Trabajo_{mes_sel}_{anio_sel}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- PESTAÑA 2: VISOR DE XML (ESTILO REPRESENTACIÓN IMPRESA) ---
with tab_visor:
    st.header("📄 Visor Avanzado de CFDI")
    tipo_visor = st.radio("¿Qué facturas deseas consultar?", ["Emitidas (Ingresos)", "Recibidas (Gastos)"], horizontal=True)
    
    archivos_a_mostrar = xml_ingresos if tipo_visor == "Emitidas (Ingresos)" else xml_gastos
    tipo_str = 'ingreso' if tipo_visor == "Emitidas (Ingresos)" else 'gasto'
    label_busqueda = "👤 Buscar Cliente:" if tipo_visor == "Emitidas (Ingresos)" else "🏢 Buscar Proveedor:"
    
    if archivos_a_mostrar:
        metadatos = [obtener_metadatos_basicos(f, tipo_str) for f in archivos_a_mostrar]
        df_meta = pd.DataFrame([m for m in metadatos if m is not None])
        
        if not df_meta.empty:
            nombres_unicos = sorted(df_meta['entidad'].unique().tolist())
            with st.expander("⚙️ Filtros de Búsqueda", expanded=True):
                f_col1, f_col2 = st.columns(2)
                filtro_entidad = f_col1.selectbox(label_busqueda, ["Todos"] + nombres_unicos)
                filtro_metodo = f_col2.selectbox("🏷️ Método de Pago:", ["Todos", "PUE", "PPD", "Pago (Complemento)"])
                
                if filtro_entidad != "Todos":
                    df_meta = df_meta[df_meta['entidad'] == filtro_entidad]
                if filtro_metodo != "Todos":
                    df_meta = df_meta[df_meta['metodo'] == filtro_metodo]
            
            st.divider()
            if not df_meta.empty:
                opciones_mostrar = df_meta['display'].tolist()
                seleccion_display = st.selectbox("📄 Selecciona un CFDI para ver el desglose visual:", opciones_mostrar)
                
                nombre_archivo_seleccionado = df_meta[df_meta['display'] == seleccion_display]['nombre_archivo'].iloc[0]
                archivo_activo = next(f for f in archivos_a_mostrar if f.name == nombre_archivo_seleccionado)
                
                d = extraer_cfdi_completo(archivo_activo)
                
                # Tarjeta HTML Profesional tipo Representación Impresa
                html_cfdi = f"""
                <div style="font-family: Arial, sans-serif; border: 1px solid #bdc3c7; border-radius: 8px; padding: 25px; background-color: #ffffff; color: #2c3e50;">
                    <div style="display: flex; justify-content: space-between; border-bottom: 2px solid #2980b9; padding-bottom: 12px; margin-bottom: 20px;">
                        <div>
                            <h2 style="margin: 0; color: #2980b9; font-size: 22px;">Representación Impresa de CFDI</h2>
                            <p style="margin: 5px 0 0 0; font-size: 13px; color: #7f8c8d;"><strong>UUID:</strong> {d['uuid']}</p>
                        </div>
                        <div style="text-align: right;">
                            <p style="margin: 0; font-size: 14px;"><strong>Serie y Folio:</strong> {d['serie']} {d['folio']}</p>
                            <p style="margin: 3px 0 0 0; font-size: 13px; color: #7f8c8d;">Fecha y Hora: {d['fecha']} {d['hora']}</p>
                        </div>
                    </div>

                    <div style="display: flex; justify-content: space-between; margin-bottom: 20px;">
                        <div style="width: 48%; padding: 15px; background-color: #f8f9fa; border-left: 4px solid #2980b9; border-radius: 4px;">
                            <h4 style="margin: 0 0 8px 0; color: #2980b9; font-size: 15px;">Datos del Emisor</h4>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>Nombre:</strong> {d['emi_nombre']}</p>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>RFC:</strong> {d['emi_rfc']}</p>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>Régimen Fiscal:</strong> {d['emi_regimen']}</p>
                        </div>
                        <div style="width: 48%; padding: 15px; background-color: #f8f9fa; border-left: 4px solid #27ae60; border-radius: 4px;">
                            <h4 style="margin: 0 0 8px 0; color: #27ae60; font-size: 15px;">Datos del Receptor</h4>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>Nombre:</strong> {d['rec_nombre']}</p>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>RFC:</strong> {d['rec_rfc']}</p>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>Régimen Receptor:</strong> {d['rec_regimen']}</p>
                            <p style="margin: 3px 0; font-size: 13px;"><strong>Uso CFDI:</strong> {d['rec_uso']}</p>
                        </div>
                    </div>

                    <h4 style="margin-bottom: 8px; color: #34495e; font-size: 15px; border-bottom: 1px solid #ecf0f1; padding-bottom: 5px;">Conceptos</h4>
                    <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px;">
                        <tr style="background-color: #2980b9; color: white;">
                            <th style="padding: 8px; text-align: center; border: 1px solid #bdc3c7;">Cant.</th>
                            <th style="padding: 8px; text-align: left; border: 1px solid #bdc3c7;">Descripción</th>
                            <th style="padding: 8px; text-align: right; border: 1px solid #bdc3c7;">Valor Unitario</th>
                            <th style="padding: 8px; text-align: right; border: 1px solid #bdc3c7;">Importe</th>
                        </tr>
                """
                for c in d['conceptos']:
                    html_cfdi += f"""
                        <tr>
                            <td style="padding: 8px; text-align: center; border: 1px solid #bdc3c7;">{c['cantidad']}</td>
                            <td style="padding: 8px; text-align: left; border: 1px solid #bdc3c7;">{c['descripcion']}</td>
                            <td style="padding: 8px; text-align: right; border: 1px solid #bdc3c7;">${c['valor_unitario']:,.2f}</td>
                            <td style="padding: 8px; text-align: right; border: 1px solid #bdc3c7;">${c['importe']:,.2f}</td>
                        </tr>
                    """
                
                html_cfdi += f"""
                    </table>

                    <div style="display: flex; justify-content: flex-end;">
                        <div style="width: 320px; padding: 15px; background-color: #f8f9fa; border-radius: 6px; border: 1px solid #bdc3c7; font-size: 13px;">
                            <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                                <strong>Subtotal:</strong> <span>${d['subtotal']:,.2f}</span>
                            </div>
                """
                if d['descuento'] > 0:
                    html_cfdi += f"""
                            <div style="display: flex; justify-content: space-between; margin-bottom: 4px; color: #c0392b;">
                                <strong>Descuento:</strong> <span>-${d['descuento']:,.2f}</span>
                            </div>
                    """
                if d['iva_trasladado'] > 0:
                    html_cfdi += f"""
                            <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                                <strong>I.V.A. (Trasladado):</strong> <span>${d['iva_trasladado']:,.2f}</span>
                            </div>
                    """
                if d['retencion_isr'] > 0 if 'retencion_isr' in d else False: # por compatibilidad
                    pass
                if d['isr_retenido'] > 0:
                    html_cfdi += f"""
                            <div style="display: flex; justify-content: space-between; margin-bottom: 4px; color: #c0392b;">
                                <strong>Retención I.S.R.:</strong> <span>-${d['isr_retenido']:,.2f}</span>
                            </div>
                    """
                if d['iva_retenido'] > 0:
                    html_cfdi += f"""
                            <div style="display: flex; justify-content: space-between; margin-bottom: 4px; color: #c0392b;">
                                <strong>Retención I.V.A.:</strong> <span>-${d['iva_retenido']:,.2f}</span>
                            </div>
                    """
                html_cfdi += f"""
                            <div style="display: flex; justify-content: space-between; font-size: 1.2em; border-top: 2px solid #bdc3c7; padding-top: 6px; margin-top: 6px; color: #2c3e50;">
                                <strong>Total:</strong> <strong>${d['total']:,.2f} {d['moneda']}</strong>
                            </div>
                        </div>
                    </div>
                </div>
                """
                st.markdown(html_cfdi, unsafe_allow_html=True)
    else:
        st.info("Sube archivos XML en la pestaña de 'Calculadora Mensual' para consultarlos aquí.")

# --- PESTAÑA 3: HISTORIAL Y ACUSES (ORDEN CRONOLÓGICO DE MESES) ---
with tab_historial:
    st.header("🗂️ Historial Mensual y Verificación de Acuses")
    
    if lista_clientes:
        cliente_hist = st.selectbox("Selecciona Cliente", nombres_clientes, key="cli_hist")
        cliente_id_hist = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_hist)
        
        try:
            historial_db = supabase.table("historial_calculos").select("*").eq("cliente_id", cliente_id_hist).execute()
            
            if historial_db.data:
                df_historial = pd.DataFrame(historial_db.data)
                
                # ORDEN CRONOLÓGICO DE MESES
                df_historial["mes_idx"] = df_historial["mes"].apply(lambda x: MESES_LISTA.index(x) if x in MESES_LISTA else 99)
                df_historial = df_historial.sort_values(by=["anio", "mes_idx"], ascending=[False, True]).drop(columns=["mes_idx"])
                
                st.dataframe(df_historial[["mes", "anio", "estatus", "ingresos_base", "isr_determinado", "iva_cargo_favor"]], use_container_width=True)
                
                st.divider()
                st.subheader("Subir Acuse de Declaración (PDF)")
                mes_acuse = st.selectbox("Mes a verificar", df_historial["mes"].unique(), key="mes_ac")
                anio_acuse = st.selectbox("Año a verificar", df_historial["anio"].unique(), key="anio_ac")
                
                acuse_pdf = st.file_uploader("Sube el acuse del SAT en PDF", type=['pdf'], key=f"pdf_{st.session_state.reset_key}")
                
                if acuse_pdf:
                    with st.spinner("Leyendo acuse del SAT..."):
                        texto_pdf = ""
                        with pdfplumber.open(acuse_pdf) as pdf:
                            for p in pdf.pages: texto_pdf += p.extract_text() + "\n"
                        
                        cantidades = re.findall(r"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b", texto_pdf)
                        
                        if cantidades:
                            cantidades_unicas = list(dict.fromkeys(cantidades))
                            st.success(f"¡Se detectaron cifras numéricas en el acuse del SAT!")
                            
                            monto_seleccionado = st.selectbox("Selecciona el importe clave de la declaración para validar:", cantidades_unicas)
                            
                            if st.button("✅ Confirmar y Marcar como Declarado", use_container_width=True):
                                registro = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id_hist).eq("mes", mes_acuse).eq("anio", int(anio_acuse)).execute()
                                if registro.data:
                                    supabase.table("historial_calculos").update({"estatus": "Declarado y Cuadrado"}).eq("id", registro.data[0]["id"]).execute()
                                    st.balloons()
                                    st.success(f"¡Estatus actualizado! Mes declarado correctamente.")
                                    st.rerun()
                        else:
                            st.warning("No se pudo extraer texto numérico automáticamente. Revisa el texto extraído:")
                            with st.expander("Ver texto completo del PDF"):
                                st.text(texto_pdf)
            else:
                st.info("No hay cálculos guardados para este cliente todavía. Realiza un cálculo en la Pestaña 1 y guárdalo.")
        except Exception as e:
            st.error(f"Error al cargar el historial desde Supabase: {e}")
    else:
        st.info("Registra un cliente en la Pestaña 1 primero.")
