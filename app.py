import os
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from openpyxl.styles import numbers, PatternFill
from datetime import datetime
import io

# Inicializa la aplicación Flask
app = Flask(__name__)
# ¡IMPORTANTE! Cambia esta clave secreta por una cadena larga y aleatoria
# Esto es crucial para la seguridad de tu aplicación en producción.
app.secret_key = 'tu_clave_secreta_aqui_CAMBIALA_por_algo_seguro_y_largo' # Reemplaza con una clave segura

# La contraseña correcta para el inicio de sesión
CORRECT_PASSWORD = "AFC2024*" 

# Configuración de la carpeta de subidas
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xml'}

# Asegurarse de que la carpeta de subidas existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    """Verifica si la extensión del archivo es XML."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def formatear_numero(valor):
    """Formatea un valor numérico, reemplazando puntos por comas."""
    if valor is None:
        return ""
    # Asegúrate de que el valor sea una cadena antes de intentar reemplazar
    return str(valor).replace(".", ",")

def formatear_fecha(fecha_str):
    """Formatea una cadena de fecha a 'dd-mm-yyyy'."""
    if fecha_str:
        try:
            # Reemplaza 'Z' para compatibilidad con datetime.fromisoformat
            fecha_obj = datetime.fromisoformat(fecha_str.replace('Z', '+00:00'))
            return fecha_obj.strftime('%d-%m-%Y')
        except ValueError:
            return fecha_str  # Devuelve la cadena original si no se puede formatear
    return ""

def extraer_datos_xml_en_memoria(xml_files, numero_receptor_filtro):
    """
    Extrae datos de archivos XML (en memoria) y los guarda en un objeto BytesIO de Excel.
    Aplica formateo y resaltado similar al script original.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    # Define los encabezados de las columnas del Excel
    headers = ["Clave", "Consecutivo", "Fecha", "Nombre Emisor", "Número Emisor", "Nombre Receptor", "Número Receptor",
                         "Código Cabys", "Detalle", "Cantidad", "Precio Unitario", "Monto Total", "Monto Descuento", "Subtotal",
                         "Tarifa (%)", "Monto Impuesto", "Impuesto Neto", "Código Moneda", "Tipo Cambio",
                         "Total Gravado", "Total Exento", "Total Exonerado", "Total Venta", "Total Descuentos",
                         "Total Venta Neta", "Total Impuesto", "Total Comprobante", "Otros Cargos", "Archivo", "Tipo de Documento"]
    ws.append(headers)

    # Itera sobre cada archivo XML recibido
    for uploaded_file in xml_files:
        filename = uploaded_file.filename # Obtiene el nombre original del archivo
        try:
            # Parsea el archivo XML directamente desde el objeto de archivo cargado
            tree = ET.parse(uploaded_file)
            root = tree.getroot()
            # Define el namespace para buscar elementos XML correctamente
            ns = {'cf': 'https://cdn.comprobanteselectronicos.go.cr/xml-schemas/v4.3/facturaElectronica'}
            
            # Limpiar namespaces para compatibilidad
            for elem in root.iter():
                elem.tag = elem.tag.split('}', 1)[-1]
            ns_clean = {}

            # Verifica si es una Factura Electrónica
            if root.tag.endswith('FacturaElectronica') or root.tag == 'FacturaElectronica':
                # Extrae todos los datos relevantes de la factura
                clave_element = root.find('Clave')
                clave = clave_element.text if clave_element is not None else ""
                consecutivo_element = root.find('NumeroConsecutivo')
                consecutivo = consecutivo_element.text if consecutivo_element is not None else ""
                fecha_element = root.find('FechaEmision')
                fecha = formatear_fecha(fecha_element.text) if fecha_element is not None else ""
                nombre_emisor_element = root.find('Emisor/Nombre')
                nombre_emisor = nombre_emisor_element.text if nombre_emisor_element is not None else ""
                numero_emisor_element = root.find('Emisor/Identificacion/Numero')
                numero_emisor = numero_emisor_element.text if numero_emisor_element is not None else ""
                nombre_receptor_element = root.find('Receptor/Nombre')
                nombre_receptor = nombre_receptor_element.text if nombre_receptor_element is not None else ""
                numero_receptor_element = root.find('Receptor/Identificacion/Numero')
                numero_receptor = numero_receptor_element.text if numero_receptor_element is not None else ""

                detalles_servicio = root.find('DetalleServicio')
                # Maneja el caso en que no haya detalles de servicio
                lineas_detalle = detalles_servicio.findall('LineaDetalle') if detalles_servicio is not None else []

                # Itera sobre cada línea de detalle de la factura
                for linea in lineas_detalle:
                    codigo_cabys_element = linea.find('Codigo')
                    codigo_cabys = codigo_cabys_element.text if codigo_cabys_element is not None else ""
                    detalle_element = linea.find('Detalle')
                    detalle = detalle_element.text if detalle_element is not None else ""
                    cantidad_element = linea.find('Cantidad')
                    cantidad = formatear_numero(cantidad_element.text) if cantidad_element is not None else ""
                    precio_unitario_element = linea.find('PrecioUnitario')
                    precio_unitario = formatear_numero(precio_unitario_element.text) if precio_unitario_element is not None else ""
                    monto_total_element = linea.find('MontoTotal')
                    monto_total = formatear_numero(monto_total_element.text) if monto_total_element is not None else ""
                    monto_descuento_element = linea.find('Descuento/MontoDescuento')
                    monto_descuento = formatear_numero(monto_descuento_element.text) if monto_descuento_element is not None else "0,00"
                    subtotal_element = linea.find('SubTotal')
                    subtotal = formatear_numero(subtotal_element.text) if subtotal_element is not None else ""
                    impuesto = linea.find('Impuesto')
                    tarifa = formatear_numero(impuesto.find('Tarifa').text) if impuesto is not None and impuesto.find('Tarifa') is not None else "0,00"
                    monto_impuesto = formatear_numero(impuesto.find('Monto').text) if impuesto is not None and impuesto.find('Monto') is not None else "0,00"
                    impuesto_neto_element = linea.find('ImpuestoNeto')
                    impuesto_neto = formatear_numero(impuesto_neto_element.text) if impuesto_neto_element is not None else ""

                    codigo_moneda_element = root.find('ResumenFactura/CodigoTipoMoneda/CodigoMoneda')
                    codigo_moneda = codigo_moneda_element.text if codigo_moneda_element is not None else ""
                    tipo_cambio_element = root.find('ResumenFactura/CodigoTipoMoneda/TipoCambio')
                    tipo_cambio = formatear_numero(tipo_cambio_element.text) if tipo_cambio_element is not None else ""
                    total_gravado_element = root.find('ResumenFactura/TotalGravado')
                    total_gravado = formatear_numero(total_gravado_element.text) if total_gravado_element is not None else ""
                    total_exento_element = root.find('ResumenFactura/TotalExento')
                    total_exento = formatear_numero(total_exento_element.text) if total_exento_element is not None else ""
                    total_exonerado_element = root.find('ResumenFactura/TotalExonerado')
                    total_exonerado = formatear_numero(total_exonerado_element.text) if total_exonerado_element is not None else ""
                    total_venta_element = root.find('ResumenFactura/TotalVenta')
                    total_venta = formatear_numero(total_venta_element.text) if total_venta_element is not None else ""
                    total_descuentos_element = root.find('ResumenFactura/TotalDescuentos')
                    total_descuentos = formatear_numero(total_descuentos_element.text) if total_descuentos_element is not None else ""
                    total_venta_neta_element = root.find('ResumenFactura/TotalVentaNeta')
                    total_venta_neta = formatear_numero(total_venta_neta_element.text) if total_venta_neta_element is not None else ""
                    total_impuesto_element = root.find('ResumenFactura/TotalImpuesto')
                    total_impuesto = formatear_numero(total_impuesto_element.text) if total_impuesto_element is not None else ""
                    total_comprobante_element = root.find('ResumenFactura/TotalComprobante')
                    total_comprobante = formatear_numero(total_comprobante_element.text) if total_comprobante_element is not None else ""
                    otros_cargos_element = root.find('OtrosCargos/MontoCargo')
                    otros_cargos = formatear_numero(otros_cargos_element.text) if otros_cargos_element is not None else ""

                    # Construye la fila de datos para el Excel
                    fila_excel = [clave, consecutivo, fecha, nombre_emisor, numero_emisor, nombre_receptor, numero_receptor,
                                  codigo_cabys, detalle, cantidad, precio_unitario, monto_total, monto_descuento, subtotal,
                                  tarifa, monto_impuesto, impuesto_neto, codigo_moneda, tipo_cambio,
                                  total_gravado, total_exento, total_exonerado, total_venta, total_descuentos,
                                  total_venta_neta, total_impuesto, total_comprobante, otros_cargos, filename, "Factura Electronica"]
                    ws.append(fila_excel)
            else:
                # Si no es una Factura Electrónica, añade una fila indicando que es otro tipo de documento
                fila_excel = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", filename, "Otro Documento"]
                ws.append(fila_excel)

        except ET.ParseError as e:
            flash(f"Error al parsear el archivo XML '{filename}': {e}", 'error')
        except Exception as e:
            flash(f"Error al procesar el archivo '{filename}': {e}", 'error')

    # --- Aplicar formato y resaltado al Excel ---
    # Convertir columnas a formato numérico (para los valores monetarios)
    columnas_numericas = [10, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]
    for fila in ws.iter_rows(min_row=2): # Comienza desde la segunda fila (después de los encabezados)
        for indice_columna in columnas_numericas:
            celda = fila[indice_columna - 1] # openpyxl usa índices base 0
            try:
                # Reemplaza la coma por punto para que Python lo reconozca como float
                if isinstance(celda.value, str):
                    celda.value = float(celda.value.replace(",", "."))
                # Aplica formato numérico de Excel (miles con coma, decimales con punto)
                celda.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            except (ValueError, AttributeError, TypeError):
                # Ignora errores si la celda no contiene un número válido
                pass

    # Resaltar las columnas especificadas con color azul claro
    columnas_a_resaltar = [2, 3, 9, 16, 21, 24, 25, 29] # Índices de columnas a resaltar (basados en 1)
    fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Color azul claro
    for col_idx in columnas_a_resaltar:
        if 0 < col_idx <= ws.max_column: # Asegura que el índice de la columna sea válido
            # Accede a la columna completa y aplica el relleno a cada celda
            columna = list(ws.columns)[col_idx - 1]
            for cell in columna:
                cell.fill = fill

    # Resaltar números de receptor diferentes al filtro en color rojo
    # La columna del "Número Receptor" es la 7ma (índice 6 en base 0)
    if 6 < ws.max_column:
        columna_numero_receptor = list(ws.columns)[6] # Accede a la columna G
        fill_rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid") # Color rojo
        # Itera desde la segunda celda para evitar el encabezado
        for cell in columna_numero_receptor[1:]:
            if cell.value != numero_receptor_filtro and cell.value is not None and numero_receptor_filtro is not None:
                cell.fill = fill_rojo

    # Eliminar filas completamente vacías
    filas_a_eliminar = []
    # Itera desde la segunda fila (saltando encabezados)
    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        # Comprueba si todas las celdas en la fila están vacías o contienen solo espacios en blanco
        if all(cell.value is None or (isinstance(cell.value, str) and str(cell.value).strip() == '') for cell in row):
            filas_a_eliminar.append(idx)

    # Eliminar las filas en orden inverso para no afectar los índices durante la eliminación
    for row_idx in reversed(filas_a_eliminar):
        ws.delete_rows(row_idx)

    # Guarda el libro de Excel en un objeto BytesIO (buffer en memoria)
    # Esto es crucial para enviar el archivo directamente al navegador sin guardarlo en disco.
    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0) # Mueve el puntero al inicio del stream para que se pueda leer desde el principio

    return excel_stream

# --- Rutas de la aplicación web Flask ---

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

    # Llama a la función para extraer datos y generar el Excel en memoria
    excel_stream = extraer_datos_xml_en_memoria(files, numero_receptor)
    
    # Envía el archivo Excel generado al navegador para su descarga
    flash('Excel generado y listo para descargar.', 'success')
    return send_file(
        excel_stream,
        download_name='datos_facturas.xlsx', # Nombre del archivo que se descargará
        as_attachment=True, # Indica que es un archivo adjunto para descargar
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' # Tipo MIME para archivos .xlsx
    )

if __name__ == '__main__':
    app.run(debug=True)
