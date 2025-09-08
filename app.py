import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl.styles import numbers, PatternFill
from datetime import datetime
import io
import uuid

# Inicializa la aplicación Flask
app = Flask(__name__)
# Usamos una clave secreta segura y generada dinámicamente
app.secret_key = os.environ.get("SECRET_KEY_APP_XML", str(uuid.uuid4()))

# --- Funciones de formateo de datos ---
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

# --- Lógica principal para extraer datos XML y generar el Excel ---
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

            # Verifica si es una Factura Electrónica
            if root.tag.endswith('FacturaElectronica'):
                # Extrae todos los datos relevantes de la factura
                clave_element = root.find('cf:Clave', ns)
                clave = clave_element.text if clave_element is not None else ""
                consecutivo_element = root.find('cf:NumeroConsecutivo', ns)
                consecutivo = consecutivo_element.text if consecutivo_element is not None else ""
                fecha_element = root.find('cf:FechaEmision', ns)
                fecha = formatear_fecha(fecha_element.text) if fecha_element is not None else ""
                nombre_emisor_element = root.find('cf:Emisor/cf:Nombre', ns)
                nombre_emisor = nombre_emisor_element.text if nombre_emisor_element is not None else ""
                numero_emisor_element = root.find('cf:Emisor/cf:Identificacion/cf:Numero', ns)
                numero_emisor = numero_emisor_element.text if numero_emisor_element is not None else ""
                nombre_receptor_element = root.find('cf:Receptor/cf:Nombre', ns)
                nombre_receptor = nombre_receptor_element.text if nombre_receptor_element is not None else ""
                numero_receptor_element = root.find('cf:Receptor/cf:Identificacion/cf:Numero', ns)
                numero_receptor = numero_receptor_element.text if numero_receptor_element is not None else ""

                detalles_servicio = root.find('cf:DetalleServicio', ns)
                # Maneja el caso en que no haya detalles de servicio
                lineas_detalle = detalles_servicio.findall('cf:LineaDetalle', ns) if detalles_servicio is not None else []

                # Itera sobre cada línea de detalle de la factura
                for linea in lineas_detalle:
                    codigo_cabys_element = linea.find('cf:Codigo', ns)
                    codigo_cabys = codigo_cabys_element.text if codigo_cabys_element is not None else ""
                    detalle_element = linea.find('cf:Detalle', ns)
                    detalle = detalle_element.text if detalle_element is not None else ""
                    cantidad_element = linea.find('cf:Cantidad', ns)
                    cantidad = formatear_numero(cantidad_element.text) if cantidad_element is not None else ""
                    precio_unitario_element = linea.find('cf:PrecioUnitario', ns)
                    precio_unitario = formatear_numero(precio_unitario_element.text) if precio_unitario_element is not None else ""
                    monto_total_element = linea.find('cf:MontoTotal', ns)
                    monto_total = formatear_numero(monto_total_element.text) if monto_total_element is not None else ""
                    monto_descuento_element = linea.find('cf:Descuento/cf:MontoDescuento', ns)
                    monto_descuento = formatear_numero(monto_descuento_element.text) if monto_descuento_element is not None else "0,00"
                    subtotal_element = linea.find('cf:SubTotal', ns)
                    subtotal = formatear_numero(subtotal_element.text) if subtotal_element is not None else ""
                    impuesto = linea.find('cf:Impuesto', ns)
                    tarifa = formatear_numero(impuesto.find('cf:Tarifa', ns).text) if impuesto is not None and impuesto.find('cf:Tarifa', ns) is not None else "0,00"
                    monto_impuesto = formatear_numero(impuesto.find('cf:Monto', ns).text) if impuesto is not None and impuesto.find('cf:Monto', ns) is not None else "0,00"
                    impuesto_neto_element = linea.find('cf:ImpuestoNeto', ns)
                    impuesto_neto = formatear_numero(impuesto_neto_element.text) if impuesto_neto_element is not None else ""

                    codigo_moneda_element = root.find('cf:ResumenFactura/cf:CodigoTipoMoneda/cf:CodigoMoneda', ns)
                    codigo_moneda = codigo_moneda_element.text if codigo_moneda_element is not None else ""
                    tipo_cambio_element = root.find('cf:ResumenFactura/cf:CodigoTipoMoneda/cf:TipoCambio', ns)
                    tipo_cambio = formatear_numero(tipo_cambio_element.text) if tipo_cambio_element is not None else ""
                    total_gravado_element = root.find('cf:ResumenFactura/cf:TotalGravado', ns)
                    total_gravado = formatear_numero(total_gravado_element.text) if total_gravado_element is not None else ""
                    total_exento_element = root.find('cf:ResumenFactura/cf:TotalExento', ns)
                    total_exento = formatear_numero(total_exento_element.text) if total_exento_element is not None else ""
                    total_exonerado_element = root.find('cf:ResumenFactura/cf:TotalExonerado', ns)
                    total_exonerado = formatear_numero(total_exonerado_element.text) if total_exonerado_element is not None else ""
                    total_venta_element = root.find('cf:ResumenFactura/cf:TotalVenta', ns)
                    total_venta = formatear_numero(total_venta_element.text) if total_venta_element is not None else ""
                    total_descuentos_element = root.find('cf:ResumenFactura/cf:TotalDescuentos', ns)
                    total_descuentos = formatear_numero(total_descuentos_element.text) if total_descuentos_element is not None else ""
                    total_venta_neta_element = root.find('cf:ResumenFactura/cf:TotalVentaNeta', ns)
                    total_venta_neta = formatear_numero(total_venta_neta_element.text) if total_venta_neta_element is not None else ""
                    total_impuesto_element = root.find('cf:ResumenFactura/cf:TotalImpuesto', ns)
                    total_impuesto = formatear_numero(total_impuesto_element.text) if total_impuesto_element is not None else ""
                    total_comprobante_element = root.find('cf:ResumenFactura/cf:TotalComprobante', ns)
                    total_comprobante = formatear_numero(total_comprobante_element.text) if total_comprobante_element is not None else ""
                    otros_cargos_element = root.find('cf:OtrosCargos/cf:MontoCargo', ns)
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
            print(f"Error al parsear el archivo XML '{filename}': {e}")
        except Exception as e:
            print(f"Error al procesar el archivo '{filename}': {e}")

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
            # Accede a la columna completa y aplica el relleno al encabezado y las celdas
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
    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0) # Mueve el puntero al inicio del stream para que se pueda leer desde el principio

    return excel_stream

# --- Rutas de la aplicación web Flask ---
@app.route("/")
def index():
    if "logged_in" not in session:
        return redirect(url_for("login"))
    # Renderiza el index.html solo si el usuario está logueado
    return render_template("index.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password")
        # Aquí se corrige la contraseña para que sea la clave correcta
        if password == "AFC2024*":
            session["logged_in"] = True
            flash("Inicio de sesión exitoso.", "success")
            return redirect(url_for("index"))
        else:
            flash("Contraseña incorrecta. Inténtalo de nuevo.", "danger")
            return redirect(url_for("login"))
    return render_template("login.html")

@app.route('/upload', methods=['POST'])
def upload_files():
    """Ruta para manejar la carga de archivos XML y generar el Excel."""
    # Verifica que el usuario haya iniciado sesión antes de procesar los archivos
    if "logged_in" not in session:
        flash("Necesitas iniciar sesión para acceder a esta función.", "warning")
        return redirect(url_for("login"))

    # Verifica si se enviaron archivos XML en la solicitud
    if 'files[]' not in request.files:
        flash('No se encontraron archivos XML. Por favor, selecciona al menos uno.')
        return redirect(url_for('index'))

    # Obtiene la lista de archivos XML subidos
    xml_files = request.files.getlist('files[]')
    if not xml_files or all(f.filename == '' for f in xml_files):
        flash('No se seleccionó ningún archivo XML. Por favor, arrastra o selecciona archivos.')
        return redirect(url_for('index'))

    # Obtiene el número de receptor ingresado por el usuario
    # Ahora el campo de contraseña está eliminado, solo se usa el campo del número de identificación
    numero_receptor_filtro = request.form.get('idReceptor', '').strip()
    if not numero_receptor_filtro:
        flash('Por favor, ingrese el número de identificación del receptor.')
        return redirect(url_for('index'))

    try:
        # Llama a la función para extraer datos y generar el Excel en memoria
        excel_stream = extraer_datos_xml_en_memoria(xml_files, numero_receptor_filtro)
        # Envía el archivo Excel generado al navegador para su descarga
        return send_file(
            excel_stream,
            download_name='datos_facturas.xlsx', # Nombre del archivo que se descargará
            as_attachment=True, # Indica que es un archivo adjunto para descargar
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' # Tipo MIME para archivos .xlsx
        )
    except Exception as e:
        # Manejo de errores durante el procesamiento
        flash(f'Ocurrió un error al procesar los archivos: {e}')
        # Redirige de vuelta a la página de inicio con el mensaje de error
        return redirect(url_for('index'))

@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    flash("Has cerrado sesión.", "info")
    return redirect(url_for("login"))

# Bloque principal para ejecutar la aplicación Flask
if __name__ == '__main__':
    # Asegura que la carpeta 'templates' exista al iniciar la app
    if not os.path.exists('templates'):
        os.makedirs('templates')
    app.run(debug=True, host='0.0.0.0', port=5000)
