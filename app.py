import os
from flask import Flask, request, render_template, redirect, url_for, send_from_directory
import openpyxl
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename

app = Flask(__name__)

EXCEL_FILE = "datos.xlsx"

# Carpeta donde guardamos las imágenes subidas
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "imagenes_subidas")
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# ----------------------------------------------------------
# FUNCIONES AUXILIARES PARA LEER Y ESCRIBIR DATOS ORDENADOS
# ----------------------------------------------------------
def leer_datos_excel():
    """
    Lee todos los registros de 'datos.xlsx' y retorna una lista de diccionarios.
    Cada diccionario tiene:
      {
        "numero":  (str),
        "descripcion": (str),
        "peso": (str) o float,
        "valor": (str) o float,
        "imagenes": [ruta_imagen1, ruta_imagen2, ruta_imagen3] (puede tener None)
      }
    Si el archivo no existe, retorna lista vacía.
    """
    if not os.path.exists(EXCEL_FILE):
        return []

    wb = load_workbook(EXCEL_FILE, read_only=False)
    hoja = wb.active

    data = []
    # Asumimos fila 1 = cabeceras, por lo que empezamos en 2
    for row in hoja.iter_rows(min_row=2, max_col=7):
        numero = row[0].value  # A
        if numero is None:
            continue  # Fila vacía
        numero = str(numero)   # lo tratamos siempre como string para evitar problemas

        descripcion = row[1].value  # B
        peso = row[2].value         # C
        valor = row[3].value        # D

        # E, F, G => imágenes
        imagenes = []
        for celda_imagen in row[4:7]:
            if celda_imagen and celda_imagen.hyperlink:
                imagenes.append(celda_imagen.hyperlink.target)
            else:
                imagenes.append(None)

        data.append({
            "numero": numero,
            "descripcion": descripcion,
            "peso": peso,
            "valor": valor,
            "imagenes": imagenes
        })

    wb.close()
    return data

def escribir_datos_excel(data):
    """
    Sobrescribe por completo el archivo 'datos.xlsx' con la información
    de la lista 'data', ASUMIENDO que 'data' ya está ORDENADA por 'numero' asc.
    """
    # Creamos un libro nuevo
    wb = Workbook()
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

    fila = 2
    for item in data:
        hoja.cell(row=fila, column=1, value=item["numero"])
        hoja.cell(row=fila, column=2, value=item["descripcion"])
        hoja.cell(row=fila, column=3, value=item["peso"])
        hoja.cell(row=fila, column=4, value=item["valor"])

        # E, F, G => imágenes
        for idx, col in enumerate([5, 6, 7]):  # columnas E, F, G
            ruta = item["imagenes"][idx]
            if ruta:
                celda = hoja.cell(row=fila, column=col)
                celda.value = f"Ver Imagen {idx+1}"
                celda.hyperlink = ruta
                celda.style = "Hyperlink"
        fila += 1

    wb.save(EXCEL_FILE)
    wb.close()

# -------------------------------------------
#  RUTA PRINCIPAL (listar y crear)
# -------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Creamos o leemos la data existente
        data = leer_datos_excel()

        # Recoger datos del formulario
        numero = request.form.get("numero")
        descripcion = request.form.get("descripcion")
        peso = request.form.get("peso")
        valor = request.form.get("valor")

        # Comprobar si ya existe ese 'Numero'
        existe = any(item["numero"] == str(numero) for item in data)
        if existe:
            return f"Error: El número '{numero}' ya existe en el Excel.", 400

        # Procesar hasta 3 imágenes
        files = request.files.getlist("imagenes")
        rutas_imagenes = [None, None, None]
        for i, file in enumerate(files[:3]):
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

        # Insertar el nuevo registro en 'data'
        nuevo_registro = {
            "numero": str(numero),
            "descripcion": descripcion,
            "peso": peso,
            "valor": valor,
            "imagenes": rutas_imagenes
        }
        data.append(nuevo_registro)

        # ORDENAR por 'numero' ascendente físicamente
        # Asumimos que 'numero' se puede convertir en int. Haz try/except si necesario.
        data.sort(key=lambda x: int(x["numero"]))

        # Reescribir Excel con la data ordenada
        escribir_datos_excel(data)

        return redirect(url_for("index"))

    else:
        # GET => Listar todo lo que tengamos en Excel
        data = leer_datos_excel()
        return render_template("index.html", data=data)

# -------------------------------------------
#  RUTA PARA EDITAR (y eliminar fila)
# -------------------------------------------
@app.route("/editar/<numero>", methods=["GET", "POST"])
def editar(numero):
    if request.method == "GET":
        data = leer_datos_excel()
        # Buscar en 'data' donde item["numero"] == numero
        registro = next((item for item in data if item["numero"] == str(numero)), None)
        if not registro:
            return f"No existe el número {numero} en el Excel.", 404

        return render_template("editar.html", registro=registro)

    else:
        # POST => Guardar cambios o eliminar fila
        delete_flag = request.form.get("delete_record")
        data = leer_datos_excel()

        # Buscar el item en 'data'
        idx_encontrado = None
        for i, item in enumerate(data):
            if item["numero"] == str(numero):
                idx_encontrado = i
                break

        if idx_encontrado is None:
            return f"No existe el número {numero} en el Excel.", 404

        if delete_flag == "on":
            # ELIMINAR EL OBJETO de 'data'
            data.pop(idx_encontrado)
            # Reordenar y reescribir
            data.sort(key=lambda x: int(x["numero"]))
            escribir_datos_excel(data)
            return redirect(url_for("index"))
        else:
            # ACTUALIZAR
            nueva_descripcion = request.form.get("descripcion")
            nuevo_peso = request.form.get("peso")
            nuevo_valor = request.form.get("valor")

            # Subir hasta 3 imágenes nuevas
            files = request.files.getlist("imagenes")
            rutas_imagenes = [None, None, None]
            for i, file in enumerate(files[:3]):
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                    file.save(filepath)
                    rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

            # Checkboxes de eliminar imágenes
            remove_flags = [
                request.form.get("remove_img1"),
                request.form.get("remove_img2"),
                request.form.get("remove_img3")
            ]

            # Actualizamos los campos
            data[idx_encontrado]["descripcion"] = nueva_descripcion
            data[idx_encontrado]["peso"] = nuevo_peso
            data[idx_encontrado]["valor"] = nuevo_valor

            # Actualizar/borrar imágenes
            # Si subiste nueva imagen => sobrescribes
            for i in range(3):
                if rutas_imagenes[i]:
                    data[idx_encontrado]["imagenes"][i] = rutas_imagenes[i]
                elif remove_flags[i] == "on":
                    data[idx_encontrado]["imagenes"][i] = None

            # Reordenar y reescribir
            data.sort(key=lambda x: int(x["numero"]))
            escribir_datos_excel(data)

            return redirect(url_for("index"))

# -------------------------------------------
#  DESCARGAR EXCEL
# -------------------------------------------
@app.route("/descargar-excel")
def descargar_excel():
    if not os.path.exists(EXCEL_FILE):
        return "El Excel no existe aún.", 404
    return send_from_directory(
        directory=app.root_path,
        path=EXCEL_FILE,
        as_attachment=True
    )

# -------------------------------------------
#  SERVIR IMÁGENES
# -------------------------------------------
@app.route("/imagenes_subidas/<path:filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

if __name__ == "__main__":
    app.run(debug=True)