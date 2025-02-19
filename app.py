import os
import certifi
import ssl
from flask import Flask, request, render_template, redirect, url_for, send_from_directory, session
import openpyxl
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pymongo import MongoClient

# -------------------------------------------
#  CONFIGURACIÓN FLASK
# -------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("0d3393c333bcfe04b73c219f7c153df4d16027d2ba242d61", "fallback_key_in_dev") # Pon algo más seguro en producción

EXCEL_FILE = "datos.xlsx"

# Carpeta donde guardamos las imágenes subidas (para el catálogo de joyas)
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "imagenes_subidas")
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}

def allowed_file(filename):
    """Verifica si la extensión es válida."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# -------------------------------------------
#  CONEXIÓN A MONGODB ATLAS
# -------------------------------------------
# Según indicas, tu MONGO_URI y DB son:
MONGO_URI = "mongodb+srv://edfrutos:rYjwUC6pUNrLtbaI@cluster0.pmokh.mongodb.net/"
# Conexión usando certifi para resolver el CERTIFICATE_VERIFY_FAILED
client = MongoClient(MONGO_URI, tlsCAFile=certifi.where())

db = client["app_catalogojoyero"]  # Tu base de datos
users_collection = db["users"]     # Colección 'users'

# -------------------------------------------
#  MANEJO EXCEL - FUNCIONES AUXILIARES
# -------------------------------------------

def leer_datos_excel():
    """Lee todos los registros de 'datos.xlsx' y retorna una lista de diccionarios."""
    if not os.path.exists(EXCEL_FILE):
        return []

    wb = load_workbook(EXCEL_FILE, read_only=False)
    hoja = wb.active

    data = []
    # Asumimos fila 1 = cabeceras
    for row in hoja.iter_rows(min_row=2, max_col=7):
        numero = row[0].value
        if numero is None:
            continue
        numero = str(numero)

        descripcion = row[1].value
        peso = row[2].value
        valor = row[3].value

        # Columnas E, F, G => imágenes
        imagenes = []
        for celda_imagen in row[4:7]:  # E,F,G
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
    """Sobrescribe 'datos.xlsx' con la info en 'data', manteniendo columnas E, F, G para imágenes."""
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

        # Imágenes
        for idx, col in enumerate([5, 6, 7]):  # E,F,G
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
#  RUTAS DE AUTENTICACIÓN
# -------------------------------------------

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        nombre = request.form.get("nombre")
        email = request.form.get("email")
        password = request.form.get("password")

        # Verificar si ya existe el email
        user_existente = users_collection.find_one({"email": email})
        if user_existente:
            return "Error: Ese email ya está registrado. <a href='/register'>Volver</a>"

        # Insertar nuevo usuario
        hashed_pass = generate_password_hash(password)
        nuevo_usuario = {
            "nombre": nombre,
            "email": email,
            "password": hashed_pass
        }
        users_collection.insert_one(nuevo_usuario)

        return "Registro exitoso. <a href='/login'>Iniciar Sesión</a>"
    else:
        return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        usuario = users_collection.find_one({"email": email})
        if not usuario:
            return "Error: No existe ese usuario. <a href='/login'>Intentar de nuevo</a>"

        if check_password_hash(usuario["password"], password):
            session["email"] = email
            return redirect(url_for("index"))
        else:
            return "Error: Contraseña incorrecta. <a href='/login'>Intentar de nuevo</a>"
    else:
        return render_template("login.html")

@app.route("/recover", methods=["GET", "POST"])
def recover():
    if request.method == "POST":
        # Recuperación por nombre o email
        usuario = request.form.get("usuario")   # Introducen email o nombre
        nueva_password = request.form.get("password")

        # Buscar si coincide con email o con nombre
        encontrado = users_collection.find_one({
            "$or": [
                {"email": usuario},
                {"nombre": usuario}
            ]
        })
        if not encontrado:
            return "Error: No existe un usuario con ese nombre o email. <a href='/recover'>Intentar de nuevo</a>"

        # Actualizar la contraseña
        new_hashed_pass = generate_password_hash(nueva_password)
        users_collection.update_one(
            {"_id": encontrado["_id"]},
            {"$set": {"password": new_hashed_pass}}
        )

        return "Tu contraseña ha sido actualizada. <a href='/login'>Inicia Sesión</a>"
    else:
        return render_template("recover.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


# -------------------------------------------
#  RUTAS PRINCIPALES DEL CATÁLOGO (EXCEL)
# -------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    """Muestra tabla + formulario de alta si estás logueado."""
    if "email" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        data = leer_datos_excel()

        numero = request.form.get("numero")
        descripcion = request.form.get("descripcion")
        peso = request.form.get("peso")
        valor = request.form.get("valor")

        # Ver si el numero ya existe
        existe = any(item["numero"] == numero for item in data)
        if existe:
            return "Error: Ese Número ya existe. <a href='/'>Volver</a>"

        files = request.files.getlist("imagenes")
        rutas_imagenes = [None, None, None]
        for i, file in enumerate(files[:3]):
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

        nuevo_registro = {
            "numero": numero,
            "descripcion": descripcion,
            "peso": peso,
            "valor": valor,
            "imagenes": rutas_imagenes
        }
        data.append(nuevo_registro)
        escribir_datos_excel(data)

        return redirect(url_for("index"))
    else:
        data = leer_datos_excel()
        return render_template("index.html", data=data)

@app.route("/editar/<numero>", methods=["GET", "POST"])
def editar(numero):
    """Editar o eliminar un registro (fila) en el Excel. Requiere login."""
    if "email" not in session:
        return redirect(url_for("login"))

    if request.method == "GET":
        data = leer_datos_excel()
        registro = next((item for item in data if item["numero"] == str(numero)), None)
        if not registro:
            return f"No existe el número {numero} en el Excel."

        return render_template("editar.html", registro=registro)
    else:
        # POST
        delete_flag = request.form.get("delete_record")
        data = leer_datos_excel()

        idx_encontrado = None
        for i, item in enumerate(data):
            if item["numero"] == str(numero):
                idx_encontrado = i
                break

        if idx_encontrado is None:
            return f"No existe el número {numero} en el Excel."

        if delete_flag == "on":
            data.pop(idx_encontrado)
            escribir_datos_excel(data)
            return redirect(url_for("index"))
        else:
            nueva_descripcion = request.form.get("descripcion")
            nuevo_peso = request.form.get("peso")
            nuevo_valor = request.form.get("valor")

            files = request.files.getlist("imagenes")
            rutas_imagenes = [None, None, None]
            for i, file in enumerate(files[:3]):
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                    file.save(filepath)
                    rutas_imagenes[i] = os.path.join("imagenes_subidas", filename)

            remove_flags = [
                request.form.get("remove_img1"),
                request.form.get("remove_img2"),
                request.form.get("remove_img3")
            ]

            data[idx_encontrado]["descripcion"] = nueva_descripcion
            data[idx_encontrado]["peso"] = nuevo_peso
            data[idx_encontrado]["valor"] = nuevo_valor

            # Actualizar/borrar imágenes
            for i in range(3):
                if rutas_imagenes[i]:
                    data[idx_encontrado]["imagenes"][i] = rutas_imagenes[i]
                elif remove_flags[i] == "on":
                    data[idx_encontrado]["imagenes"][i] = None

            escribir_datos_excel(data)
            return redirect(url_for("index"))

@app.route("/descargar-excel")
def descargar_excel():
    if "email" not in session:
        return redirect(url_for("login"))

    if not os.path.exists(EXCEL_FILE):
        return "El Excel no existe aún."
    return send_from_directory(
        directory=app.root_path,
        path=EXCEL_FILE,
        as_attachment=True
    )

@app.route("/imagenes_subidas/<path:filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

# -------------------------------------------
#  MAIN
# -------------------------------------------
if __name__ == "__main__":
    # En producción, no uses debug=True, y preferiblemente un servidor WSGI
    app.run(debug=True)