import os
import io
import xml.etree.ElementTree as ET
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
import openpyxl
from openpyxl.styles import numbers, PatternFill

app = Flask(__name__)

# Seguridad desde variables de entorno
app.secret_key = os.getenv("SECRET_KEY_APP_XML", "CAMBIA_ESTA_CLAVE_EN_RENDER")
CORRECT_PASSWORD = os.getenv("APP_PASSWORD", "AFC2024*")  # Define APP_PASSWORD en Render

def formatear_numero(valor):
    if valor is None:
        return ""
    return str(valor).replace(".", ",")

def formatear_fecha(fecha_str):
    if fecha_str:
        try:
            return datetime.fromisoformat(fecha_str.replace('Z', '+00:00')).strftime('%d-%m-%Y')
        except ValueError:
            return fecha_str
    return ""

def extraer_datos_xml_en_memoria(xml_files, numero_receptor_filtro):
    wb = openpyxl.Workbook()
    ws_detalladas = wb.active
    ws_detalladas.title = "facturas_detalladas"

    headers = [
        "Clave","Consecutivo","Fecha","Nombre Emisor","Número Emisor","Nombre Receptor","Número Receptor",
        "Código Cabys","Detalle","Cantidad","Precio Unitario","Monto Total","Monto Descuento","Subtotal",
        "Tarifa (%)","Monto Impuesto","Impuesto Neto","Código Moneda","Tipo Cambio",
        "Total Gravado","Total Exento","Total Exonerado","Total Venta","Total Descuentos",
        "Total Venta Neta","Total Impuesto","Total Comprobante","Otros Cargos","Archivo","Tipo de Documento"
    ]
    ws_detalladas.append(headers)

    # Diccionario para acumular resumen por factura
    resumen_por_factura = {}

    for uploaded_file in xml_files:
        filename = uploaded_file.filename
        try:
            tree = ET.parse(uploaded_file)
            root = tree.getroot()

            # Quitar namespaces
            for elem in root.iter():
                elem.tag = elem.tag.split('}', 1)[-1]

            if root.tag.endswith('FacturaElectronica') or root.tag == 'FacturaElectronica':
                clave = root.find('Clave').text if root.find('Clave') is not None else ""
                consecutivo = root.find('NumeroConsecutivo').text if root.find('NumeroConsecutivo') is not None else ""
                fecha = formatear_fecha(root.find('FechaEmision').text) if root.find('FechaEmision') is not None else ""
                nombre_emisor = root.find('Emisor/Nombre').text if root.find('Emisor/Nombre') is not None else ""
                numero_emisor = root.find('Emisor/Identificacion/Numero').text if root.find('Emisor/Identificacion/Numero') is not None else ""
                nombre_receptor = root.find('Receptor/Nombre').text if root.find('Receptor/Nombre') is not None else ""
                numero_receptor = root.find('Receptor/Identificacion/Numero').text if root.find('Receptor/Identificacion/Numero') is not None else ""

                detalles_servicio = root.find('DetalleServicio')
                lineas_detalle = detalles_servicio.findall('LineaDetalle') if detalles_servicio is not None else []

                # Variables para resumen
                total_venta_resumen = 0
                total_descuentos_resumen = 0
                total_venta_neta_resumen = 0
                total_impuesto_resumen = 0
                total_comprobante_resumen = 0

                for linea in lineas_detalle:
                    codigo_cabys = linea.find('Codigo').text if linea.find('Codigo') is not None else ""
                    detalle = linea.find('Detalle').text if linea.find('Detalle') is not None else ""
                    cantidad = formatear_numero(linea.find('Cantidad').text) if linea.find('Cantidad') is not None else ""
                    precio_unitario = formatear_numero(linea.find('PrecioUnitario').text) if linea.find('PrecioUnitario') is not None else ""
                    monto_total = formatear_numero(linea.find('MontoTotal').text) if linea.find('MontoTotal') is not None else ""
                    monto_descuento = formatear_numero(linea.find('Descuento/MontoDescuento').text) if linea.find('Descuento/MontoDescuento') is not None else "0,00"
                    subtotal = formatear_numero(linea.find('SubTotal').text) if linea.find('SubTotal') is not None else ""
                    impuesto = linea.find('Impuesto')
                    tarifa = formatear_numero(impuesto.find('Tarifa').text) if impuesto is not None and impuesto.find('Tarifa') is not None else "0,00"
                    monto_impuesto = formatear_numero(impuesto.find('Monto').text) if impuesto is not None and impuesto.find('Monto') is not None else "0,00"
                    impuesto_neto = formatear_numero(linea.find('ImpuestoNeto').text) if linea.find('ImpuestoNeto') is not None else ""

                    codigo_moneda = root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda').text if root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda') is not None else ""
                    tipo_cambio = formatear_numero(root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio').text) if root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio') is not None else ""
                    total_gravado = formatear_numero(root.find('ResumenFactura/TotalGravado').text) if root.find('ResumenFactura/TotalGravado') is not None else ""
                    total_exento = formatear_numero(root.find('ResumenFactura/TotalExento').text) if root.find('ResumenFactura/TotalExento') is not None else ""
                    total_exonerado = formatear_numero(root.find('ResumenFactura/TotalExonerado').text) if root.find('ResumenFactura/TotalExonerado') is not None else ""
                    total_venta = formatear_numero(root.find('ResumenFactura/TotalVenta').text) if root.find('ResumenFactura/TotalVenta') is not None else ""
                    total_descuentos = formatear_numero(root.find('ResumenFactura/TotalDescuentos').text) if root.find('ResumenFactura/TotalDescuentos') is not None else ""
                    total_venta_neta = formatear_numero(root.find('ResumenFactura/TotalVentaNeta').text) if root.find('ResumenFactura/TotalVentaNeta') is not None else ""
                    total_impuesto = formatear_numero(root.find('ResumenFactura/TotalImpuesto').text) if root.find('ResumenFactura/TotalImpuesto') is not None else ""
                    total_comprobante = formatear_numero(root.find('ResumenFactura/TotalComprobante').text) if root.find('ResumenFactura/TotalComprobante') is not None else ""
                    otros_cargos = formatear_numero(root.find('OtrosCargos/MontoCargo').text) if root.find('OtrosCargos/MontoCargo') is not None else ""

                    fila_excel = [
                        clave, consecutivo, fecha, nombre_emisor, numero_emisor, nombre_receptor, numero_receptor,
                        codigo_cabys, detalle, cantidad, precio_unitario, monto_total, monto_descuento, subtotal,
                        tarifa, monto_impuesto, impuesto_neto, codigo_moneda, tipo_cambio,
                        total_gravado, total_exento, total_exonerado, total_venta, total_descuentos,
                        total_venta_neta, total_impuesto, total_comprobante, otros_cargos, filename, "Factura Electronica"
                    ]
                    ws_detalladas.append(fila_excel)

                    # Acumular totales para resumen
                    try:
                        total_venta_resumen += float(total_venta.replace(",", "."))
                        total_descuentos_resumen += float(total_descuentos.replace(",", "."))
                        total_venta_neta_resumen += float(total_venta_neta.replace(",", "."))
                        total_impuesto_resumen += float(total_impuesto.replace(",", "."))
                        total_comprobante_resumen += float(total_comprobante.replace(",", "."))
                    except:
                        pass

                # Guardar resumen por factura
                resumen_por_factura[(clave, consecutivo)] = [
                    clave, consecutivo, fecha, nombre_emisor, numero_emisor, nombre_receptor, numero_receptor,
                    total_venta_resumen, total_descuentos_resumen, total_venta_neta_resumen, total_impuesto_resumen,
                    total_comprobante_resumen, filename, "Factura Electronica"
                ]
            else:
                ws_detalladas.append([""] * 29 + [filename, "Otro Documento"])

        except ET.ParseError as e:
            flash(f"Error al parsear '{filename}': {e}", 'error')
        except Exception as e:
            flash(f"Error al procesar '{filename}': {e}", 'error')

    # ---- Crear hoja resumida ----
    ws_resumen = wb.create_sheet(title="facturas_resumidas")
    headers_resumen = [
        "Clave","Consecutivo","Fecha","Nombre Emisor","Número Emisor","Nombre Receptor","Número Receptor",
        "Total Venta","Total Descuentos","Total Venta Neta","Total Impuesto","Total Comprobante","Archivo","Tipo de Documento"
    ]
    ws_resumen.append(headers_resumen)

    for resumen in resumen_por_factura.values():
        ws_resumen.append(resumen)

    # ---- Aplicar formato numérico y colores ----
    columnas_numericas_detalle = [10,11,12,13,14,15,16,17,19,20,21,22,23,24,25,26,27,28]
    columnas_a_resaltar_detalle = [2,3,9,16,21,24,25,29]
    fill_celeste = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_rojo = PatternFill(start_color="FFAAAA", end_color="FFAAAA", fill_type="solid")

    # Formato y colores en facturas_detalladas
    for fila in ws_detalladas.iter_rows(min_row=2):
        for idx_col in columnas_numericas_detalle:
            celda = fila[idx_col - 1]
            try:
                if isinstance(celda.value, str):
                    celda.value = float(celda.value.replace(",", "."))
                celda.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            except:
                pass
        for col_idx in columnas_a_resaltar_detalle:
            if 0 < col_idx <= ws_detalladas.max_column:
                fila[col_idx - 1].fill = fill_celeste
        # Rojo si receptor no coincide
        if fila[6].value and numero_receptor_filtro and str(fila[6].value) != str(numero_receptor_filtro):
            fila[6].fill = fill_rojo

    # Formato y colores en facturas_resumidas
    columnas_numericas_resumen = [8,9,10,11,12]
    columnas_a_resaltar_resumen = [2,3,8,11,12]
    for fila in ws_resumen.iter_rows(min_row=2):
        for idx_col in columnas_numericas_resumen:
            celda = fila[idx_col - 1]
            try:
                celda.value = float(celda.value)
                celda.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            except:
                pass
        for col_idx in columnas_a_resaltar_resumen:
            if 0 < col_idx <= ws_resumen.max_column:
                fila[col_idx - 1].fill = fill_celeste
        # Rojo si receptor no coincide
        if fila[6].value and numero_receptor_filtro and str(fila[6].value) != str(numero_receptor_filtro):
            fila[6].fill = fill_rojo

    # Eliminar filas vacías en ambas hojas
    for ws_iter in [ws_detalladas, ws_resumen]:
        vacias = []
        for i, row in enumerate(ws_iter.iter_rows(min_row=2), start=2):
            if all(c.value is None or (isinstance(c.value, str) and not c.value.strip()) for c in row):
                vacias.append(i)
        for i in reversed(vacias):
            ws_iter.delete_rows(i)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# --------- Rutas ---------

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == CORRECT_PASSWORD:
            session['logged_in'] = True
            flash('Inicio de sesión exitoso.', 'success')
            return redirect(url_for('index'))
        else:
            flash('Contraseña incorrecta. Inténtalo de nuevo.', 'error')
            return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    flash('Has cerrado sesión correctamente.', 'success')
    return redirect(url_for('login'))

@app.route('/upload', methods=['POST'])
def upload_files():
    if not session.get('logged_in'):
        flash('Por favor, inicia sesión para acceder a esta función.', 'error')
        return redirect(url_for('login'))

    if 'xml_files' not in request.files:
        flash('No se subieron archivos.', 'error')
        return redirect(url_for('index'))

    files = request.files.getlist('xml_files')
    if not files or files[0].filename == '':
        flash('No se seleccionó ningún archivo.', 'error')
        return redirect(url_for('index'))

    numero_receptor = request.form.get('numero_receptor')
    if not numero_receptor:
        flash('El número de identificación del receptor es obligatorio.', 'error')
        return redirect(url_for('index'))

    excel_stream = extraer_datos_xml_en_memoria(files, numero_receptor)
    return send_file(
        excel_stream,
        download_name='datos_facturas.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=False)
