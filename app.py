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

def obtener_mes_anterior(mes_actual, anio_actual):
    try:
        idx = MESES_LISTA.index(mes_actual)
        if idx == 0:
            return "Diciembre", int(anio_actual) - 1
        else:
            return MESES_LISTA[idx - 1], int(anio_actual)
    except:
        return None, anio_actual

# ==========================================
# 1. FUNCIONES DE PROCESAMIENTO (FILTRO ESTRICTO RESICO 626)
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
        regimen_receptor = receptor.attrib.get('RegimenFiscalReceptor', '') if receptor is not None else ''
        uso_cfdi = receptor.attrib.get('UsoCFDI', '') if receptor is not None else ''
        
        metodo_pago = root.attrib.get('MetodoPago', '')
        tipo_comprobante = root.attrib.get('TipoDeComprobante', '')

        subtotal = float(root.attrib.get('SubTotal', 0.0))
        iva_acreditable = 0.0
        
        impuestos = root.find('cfdi:Impuestos', ns)
        if impuestos is not None:
            traslados = impuestos.find('cfdi:Traslados', ns)
            if traslados is not None:
                for t in traslados.findall('cfdi:Traslado', ns):
                    if t.attrib.get('Impuesto') == '002': 
                        iva_acreditable += float(t.attrib.get('Importe', 0.0))

        # --- FILTROS ESTRICTOS RESICO (EXIGIR RECEPTOR 626) ---
        estado = "Acreditable"
        motivo = "Válido para RESICO"

        if regimen_receptor != '626':
            estado = "Excluido"
            motivo = f"Régimen fiscal del receptor ({regimen_receptor}) no es RESICO (626)"
        elif metodo_pago == 'PPD':
            estado = "Excluido"
            motivo = "Método PPD (Requiere REP para acreditar IVA)"
        elif tipo_comprobante in ['P', 'N']:
            estado = "Excluido"
            motivo = f"Comprobante tipo {tipo_comprobante} no genera IVA acreditable"
        elif uso_cfdi in ['D01', 'D02', 'D03', 'D04', 'D05', 'D06', 'D07', 'D08', 'D09', 'D10']:
            estado = "Excluido"
            motivo = f"Uso de CFDI ({uso_cfdi}) es deducción personal"

        return {
            "Archivo": archivo_subido.name,
            "Proveedor": nombre_emisor,
            "RFC Emisor": rfc_emisor,
            "Régimen": regimen_emisor,
            "Regimen Receptor": regimen_receptor,
            "Método Pago": metodo_pago,
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
            ["IVA a Favor Anterior", totales['iva_favor_anterior']],
            [f"IVA a {totales['estatus_iva']}", "=B8-B9-B10-B11"]
        ]
        df_resumen = pd.DataFrame(wb_resumen_data[1:], columns=wb_resumen_data[0])
        df_resumen.to_excel(writer, sheet_name='Resumen Fiscal', index=False)
        
        if not df_ingresos.empty: df_ingresos.to_excel(writer, sheet_name='Detalle Ingresos', index=False)
        if df_gastos is not None and not df_gastos.empty: df_gastos.to_excel(writer, sheet_name='Gastos Acreditables', index=False)
        if df_excluidos is not None and not df_excluidos.empty: df_excluidos.to_excel(writer, sheet_name='Gastos Excluidos', index=False)
    return buffer.getvalue()

# ==========================================
# 2. FUNCIONES DE EXTRACCIÓN AVANZADA (VISOR HTML)
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
    
    iva_favor_anterior = 0.0
    if cliente_sel and lista_clientes:
        cli_obj = next((c for c in lista_clientes if c['nombre'] == cliente_sel), None)
        if cli_obj:
            mes_ant, anio_ant = obtener_mes_anterior(mes_sel, anio_sel)
            if mes_ant:
                try:
                    res_ant = supabase.table("historial_calculos").select("*").eq("cliente_id", cli_obj['id']).eq("mes", mes_ant).eq("anio", int(anio_ant)).execute()
                    if res_ant.data:
                        val_cargo_favor = float(res_ant.data[0].get("iva_cargo_favor", 0.0))
                        if val_cargo_favor < 0:
                            iva_favor_anterior = abs(val_cargo_favor)
                except:
                    pass

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
                    
                    # Separar automáticos válidos
                    df_auto_acreditable = df_todos_gastos[df_todos_gastos["Estado"] == "Acreditable"]
                    
                    # --- NUEVA FUNCIÓN: EXCLUSIÓN MANUAL DE FACTURAS ACREDITABLES ---
                    excluir_manual = []
                    if not df_auto_acreditable.empty:
                        with st.expander("⚙️ Exclusión Manual Adicional de Gastos (Opcional)"):
                            opciones_manuales = df_auto_acreditable["Archivo"].tolist()
                            excluir_manual = st.multiselect(
                                "Selecciona facturas que deseas omitir manualmente de la acreditación:",
                                options=opciones_manuales,
                                format_func=lambda x: f"{x} - {df_auto_acreditable[df_auto_acreditable['Archivo']==x]['Proveedor'].values[0]} (${df_auto_acreditable[df_auto_acreditable['Archivo']==x]['Subtotal Gasto'].values[0]:,.2f})"
                            )
                    
                    # Aplicar exclusión manual si el usuario seleccionó alguna
                    if excluir_manual:
                        df_todos_gastos.loc[df_todos_gastos["Archivo"].isin(excluir_manual), "Estado"] = "Excluido"
                        df_todos_gastos.loc[df_todos_gastos["Archivo"].isin(excluir_manual), "Motivo Rechazo"] = "Exclusión manual por el contador"
                        df_todos_gastos.loc[df_todos_gastos["Archivo"].isin(excluir_manual), "IVA Acreditable"] = 0.0

                    df_gastos = df_todos_gastos[df_todos_gastos["Estado"] == "Acreditable"]
                    df_excluidos = df_todos_gastos[df_todos_gastos["Estado"] == "Excluido"]
                    total_iva_acreditable = df_gastos["IVA Acreditable"].sum()
            
            ingresos_redondeados = round(total_ingresos)
            isr_determinado, tasa_aplicada = calcular_isr_resico(ingresos_redondeados)
            isr_determinado = round(isr_determinado)
            isr_retenido_redondeado = round(total_isr_retenido)
            isr_a_pagar = max(0, isr_determinado - isr_retenido_redondeado)
            
            iva_trasladado_redondeado = round(total_iva_trasladado)
            iva_acreditable_redondeado = round(total_iva_acreditable)
            iva_retenido_redondeado = round(total_iva_retenido)
            
            iva_neto_bruto = iva_trasladado_redondeado - iva_acreditable_redondeado - iva_retenido_redondeado
            iva_neto = iva_neto_bruto - round(iva_favor_anterior)
            estatus_iva = "Cargo" if iva_neto > 0 else "Favor"
            
            st.divider()
            st.header(f"📊 Resumen del Cálculo: {mes_sel} {anio_sel}")
            
            if iva_favor_anterior > 0:
                st.info(f"💡 **Acumulado Automático:** Se aplicó un IVA a favor arrastrado del mes anterior por **${iva_favor_anterior:,.0f}**.")

            if df_excluidos is not None and not df_excluidos.empty:
                st.warning(f"🛡️ **Filtro Estricto RESICO (626):** Se excluyeron **{len(df_excluidos)} factura(s)** por régimen incorrecto, PPD sin REP, exclusión manual o uso personal.")

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
                st.write(f"• IVA a favor aplicado del periodo anterior: ${iva_favor_anterior:,.2f}")

            if df_excluidos is not None and not df_excluidos.empty:
                with st.expander("🔍 Ver detalle de facturas excluidas"):
                    st.dataframe(df_excluidos[["Archivo", "Proveedor", "Regimen Receptor", "Método Pago", "Motivo Rechazo"]], use_container_width=True)

            st.divider()
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if cliente_sel and st.button("💾 Guardar Cálculo en Base de Datos", use_container_width=True):
                    cliente_id = next(c['id'] for c in lista_clientes if c['nombre'] == cliente_sel)
                    
                    viejo = supabase.table("historial_calculos").select("id").eq("cliente_id", cliente_id).eq("mes", mes_sel).eq("anio", anio_sel).execute()
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
                    'iva_favor_anterior': round(iva_favor_anterior),
                    'iva_neto': abs(iva_neto), 'estatus_iva': estatus_iva
                }
                archivo_excel = generar_excel_formulado(df_ingresos, df_gastos, df_excluidos, diccionario_totales)
                st.download_button("📥 Descargar Papeles de Trabajo Formulados (.xlsx)", data=archivo_excel, file_name=f"Papel_Trabajo_{mes_sel}_{anio_sel}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- PESTAÑA 2: VISOR DE XML ---
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
                    <strong>Descuento:</strong> <span>-${datos_visor['descuento']:,.2f}</span></div>"""
                if datos_visor['iva_trasladado'] > 0:
                    html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                    <strong>IVA (16%):</strong> <span>${datos_visor['iva_trasladado']:,.2f}</span></div>"""
                if datos_visor['isr_retenido'] > 0:
                    html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px; color: #d9534f;">
                    <strong>Retención ISR:</strong> <span>-${datos_visor['isr_retenido']:,.2f}</span></div>"""
                if datos_visor['iva_retenido'] > 0:
                    html_cfdi += f"""<div style="display: flex; justify-content: space-between; margin-bottom: 5px; color: #d9534f;">
                    <strong>Retención IVA:</strong> <span>-${datos_visor['iva_retenido']:,.2f}</span></div>"""
                html_cfdi += f"""<div style="display: flex; justify-content: space-between; font-size: 1.2em; border-top: 1px solid #ccc; padding-top: 5px; margin-top: 5px;">
                <strong>Total:</strong> <strong>${datos_visor['total']:,.2f} {datos_visor['moneda']}</strong>
                </div></div></div></div>"""
                st.markdown(html_cfdi, unsafe_allow_html=True)
    else:
        st.info("Sube archivos XML en la pestaña de 'Calculadora Mensual' para consultarlos aquí.")

# --- PESTAÑA 3: HISTORIAL Y ACUSES ---
with tab_historial:
    st.header("🗂️ Historial Mensual y Verificación de Acuses")
    
    if lista_clientes:
        cliente_hist = st.selectbox("Selecciona Cliente", nombres_clientes, key="cli_hist")
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
