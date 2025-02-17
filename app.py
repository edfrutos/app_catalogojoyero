import os
from flask import Flask, request, render_template, redirect, url_for, send_from_directory
import openpyxl
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)

EXCEL_FILE = "datos.xlsx"

# Carpeta para subir imágenes
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "imagenes_subidas")
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}

def allowed_file(filename):
    """Comprueba si la extensión del archivo está permitida."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# ----------------------------------
#  RUTA PRINCIPAL: LISTAR Y AÑADIR
# ----------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # AÑADIR NUEVO REGISTRO
        numero = request.form.get("numero")
        descripcion = request.form.get("descripcion")
        peso = request.form.get("peso")
        valor = request.form.get("valor")

        # Obtener la lista de ficheros subidos (hasta 3)
        files = request.files.getlist("imagenes")
        # Guardaremos hasta 3 rutas de imágenes
        rutas_imagenes = [None, None, None]

        # Procesar cada archivo (máx. 3)
        for i, file in enumerate(files[:3]):  # por si suben más de 3
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

        # Crear o cargar Excel
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            hoja = wb.active
        else:
            wb = openpyxl.Workbook()
            hoja = wb.active
            hoja.title = "Datos"
            # Cabeceras
            hoja["A1"] = "Número"
            hoja["B1"] = "Descripción"
            hoja["C1"] = "Peso"
            hoja["D1"] = "Valor"
            hoja["E1"] = "Imagen1 (Ruta)"
            hoja["F1"] = "Imagen2 (Ruta)"
            hoja["G1"] = "Imagen3 (Ruta)"

        # Insertar nueva fila al final
        nueva_fila = hoja.max_row + 1
        hoja.cell(row=nueva_fila, column=1, value=numero)
        hoja.cell(row=nueva_fila, column=2, value=descripcion)
        hoja.cell(row=nueva_fila, column=3, value=peso)
        hoja.cell(row=nueva_fila, column=4, value=valor)

        # Columnas E, F, G => imagen1, imagen2, imagen3
        for idx, col in enumerate([5, 6, 7]):  # E=5, F=6, G=7
            if rutas_imagenes[idx]:
                celda = hoja.cell(row=nueva_fila, column=col)
                celda.value = f"Ver Imagen {idx+1}"
                celda.hyperlink = rutas_imagenes[idx]
                celda.style = "Hyperlink"

        wb.save(EXCEL_FILE)
        wb.close()

        return redirect(url_for("index"))

    # GET => MOSTRAR LA TABLA
    data = []
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE, read_only=False)
        hoja = wb.active

        # Leer filas (asumiendo fila 1 = cabeceras)
        for row in hoja.iter_rows(min_row=2, max_col=7):
            numero = str(row[0].value)
            descripcion = row[1].value
            peso = row[2].value
            valor = row[3].value

            # Leer hipervínculos en columnas E, F, G => row[4], row[5], row[6]
            imagenes = []
            for celda in [row[4], row[5], row[6]]:
                if celda and celda.hyperlink:
                    # Extraer la ruta real
                    imagenes.append(celda.hyperlink.target)
                else:
                    imagenes.append(None)

            data.append({
                "numero": numero,
                "descripcion": descripcion,
                "peso": peso,
                "valor": valor,
                "imagenes": imagenes  # lista de 3 posibles rutas
            })
        wb.close()

    return render_template("index.html", data=data)

# ----------------------------------
#  RUTA PARA EDITAR
# ----------------------------------
@app.route("/editar/<numero>", methods=["GET", "POST"])
def editar(numero):
    """
    - GET: Mostrar formulario con datos actuales.  
    - POST: Actualiza la fila (descripción, peso, valor) y/o sustituir imágenes.
    """
    if request.method == "GET":
        if not os.path.exists(EXCEL_FILE):
            return "No existe el Excel.", 404

        wb = load_workbook(EXCEL_FILE, read_only=False)
        hoja = wb.active

        registro = None
        for row in hoja.iter_rows(min_row=2, max_col=7):
            valor_numero = str(row[0].value)
            if valor_numero == str(numero):
                registro = {
                    "numero": valor_numero,
                    "descripcion": row[1].value,
                    "peso": row[2].value,
                    "valor": row[3].value,
                }
                break
        wb.close()

        if not registro:
            return f"No existe el número {numero} en el Excel.", 404

        return render_template("editar.html", registro=registro)
    else:
        # POST => Actualizar
        nueva_descripcion = request.form.get("descripcion")
        nuevo_peso = request.form.get("peso")
        nuevo_valor = request.form.get("valor")

        # Capturar hasta 3 imágenes nuevas
        files = request.files.getlist("imagenes")
        rutas_imagenes = [None, None, None]
        for i, file in enumerate(files[:3]):
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

        if not os.path.exists(EXCEL_FILE):
            return "No existe el Excel.", 404

        wb = load_workbook(EXCEL_FILE)
        hoja = wb.active

        fila_encontrada = None
        for row in hoja.iter_rows(min_row=2, max_col=7):
            valor_numero = str(row[0].value)
            if valor_numero == str(numero):
                fila_encontrada = row
                break

        if not fila_encontrada:
            wb.close()
            return f"No existe el número {numero} en el Excel.", 404

        # Sobrescribimos datos en las celdas B, C, D
        fila_encontrada[1].value = nueva_descripcion
        fila_encontrada[2].value = nuevo_peso
        fila_encontrada[3].value = nuevo_valor

        # Actualizar las imágenes si se subieron
        # Columnas E=4, F=5, G=6 (en zero-based indexing) => (row[4], row[5], row[6])
        for idx, celda in enumerate([fila_encontrada[4], fila_encontrada[5], fila_encontrada[6]]):
            if rutas_imagenes[idx]:
                # Sustituimos la imagen en la columna correspondiente
                celda.value = f"Ver Imagen {idx+1}"
                celda.hyperlink = rutas_imagenes[idx]
                celda.style = "Hyperlink"
            # Si no subes nada para esa columna, se deja la anterior

        wb.save(EXCEL_FILE)
        wb.close()

        return redirect(url_for("index"))

# ----------------------------------
#  DESCARGAR EXCEL
# ----------------------------------
@app.route("/descargar-excel")
def descargar_excel():
    if not os.path.exists(EXCEL_FILE):
        return "El Excel no existe aún.", 404
    return send_from_directory(
        directory=app.root_path,
        path=EXCEL_FILE,
        as_attachment=True
    )

# ----------------------------------
#  SERVIR IMÁGENES
# ----------------------------------
@app.route("/imagenes_subidas/<path:filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

if __name__ == "__main__":
    app.run(debug=True)