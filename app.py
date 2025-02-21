import os
import certifi
import secrets
from datetime import datetime, timedelta
import tempfile
import zipfile

from flask import (
    Flask, request, render_template, redirect, url_for,
    send_from_directory, session, flash, send_file
)
import openpyxl
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pymongo import MongoClient
from flask_mail import Mail, Message
from bson import ObjectId

# Forzar que se use el bundle de certificados de certifi
os.environ['SSL_CERT_FILE'] = certifi.where()

# -------------------------------------------
# CONFIGURACIÓN FLASK
# -------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "CAMBIA_ESTA_CLAVE_EN_PRODUCCION")

# Carpeta para imágenes del catálogo
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "imagenes_subidas")
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])

# Carpeta para hojas de cálculo (tablas)
SPREADSHEET_FOLDER = os.path.join(app.root_path, "spreadsheets")
if not os.path.exists(SPREADSHEET_FOLDER):
    os.makedirs(SPREADSHEET_FOLDER)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# -------------------------------------------
# CONFIGURACIÓN DE EMAIL (Flask-Mail)
# -------------------------------------------
app.config["MAIL_SERVER"] = "smtp-relay.brevo.com"
app.config["MAIL_PORT"] = 587
app.config["MAIL_USE_TLS"] = True
app.config["MAIL_USERNAME"] = "admin@edefrutos.me"
app.config["MAIL_PASSWORD"] = "Rmp3UXwsIkvA0c1d"
app.config["MAIL_DEFAULT_SENDER"] = ("Administrador", "admin@edefrutos.me")
app.config["MAIL_DEBUG"] = True
mail = Mail(app)

# -------------------------------------------
# CONEXIÓN A MONGODB ATLAS
# -------------------------------------------
# Conexión a MongoDB Atlas

from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi

MONGO_URI = "mongodb+srv://edfrutos:rYjwUC6pUNrLtbaI@cluster0.pmokh.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"

# Create a new client and connect to the server
client = MongoClient(MONGO_URI, server_api=ServerApi('1'))

# Send a ping to confirm a successful connection
try:
    client.admin.command('ping')
    print("Pinged your deployment. You successfully connected to MongoDB!")
except Exception as e:
    print(e)

db = client["app_catalogojoyero"]
users_collection = db["users"]
resets_collection = db["password_resets"]
spreadsheets_collection = db["spreadsheets"]

# -------------------------------------------
# FUNCIONES AUXILIARES PARA HOJAS DE CÁLCULO
# -------------------------------------------
def leer_datos_excel(filename):
    if not os.path.exists(filename):
        return []
    wb = load_workbook(filename, read_only=False)
    hoja = wb.active
    data = []
    for row in hoja.iter_rows(min_row=2, max_col=7):
        numero = row[0].value
        if numero is None:
            continue
        numero = str(numero)
        descripcion = row[1].value
        peso = row[2].value
        valor = row[3].value
        imagenes = []
        for celda in row[4:7]:
            if celda and celda.hyperlink:
                imagenes.append(celda.hyperlink.target)
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

def escribir_datos_excel(data, filename):
    wb = Workbook()
    hoja = wb.active
    hoja.title = "Datos"
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
        for idx, col in enumerate([5, 6, 7]):
            ruta = item["imagenes"][idx]
            if ruta:
                celda = hoja.cell(row=fila, column=col)
                celda.value = f"Ver Imagen {idx+1}"
                celda.hyperlink = ruta
                celda.style = "Hyperlink"
        fila += 1
    wb.save(filename)
    wb.close()

def get_current_spreadsheet():
    filename = session.get("selected_table")
    if not filename:
        return None
    return os.path.join(SPREADSHEET_FOLDER, filename)

# -------------------------------------------
# RUTAS DE AUTENTICACIÓN
# -------------------------------------------
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        nombre = request.form.get("nombre").strip()
        email = request.form.get("email").strip().lower()
        password = request.form.get("password").strip()
        if users_collection.find_one({"email": email}):
            return "Error: Ese email ya está registrado. <a href='/register'>Volver</a>"
        hashed = generate_password_hash(password)
        nuevo_usuario = {"nombre": nombre, "email": email, "password": hashed}
        users_collection.insert_one(nuevo_usuario)
        return "Registro exitoso. <a href='/login'>Iniciar Sesión</a>"
    else:
        return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        login_input = request.form.get("login_input").strip()
        password = request.form.get("password").strip()
        usuario = users_collection.find_one({
            "$or": [
                {"nombre": {"$regex": f"^{login_input}$", "$options": "i"}},
                {"email": {"$regex": f"^{login_input}$", "$options": "i"}}
            ]
        })
        if not usuario:
            return "Error: Usuario no encontrado. <a href='/login'>Reintentar</a>"
        if check_password_hash(usuario["password"], password):
            session["usuario"] = usuario["nombre"]
            return redirect(url_for("home"))
        else:
            return "Error: Contraseña incorrecta. <a href='/login'>Reintentar</a>"
    else:
        return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/recover")
def recover_redirect():
    return redirect(url_for("forgot_password"))

@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        usuario_input = request.form.get("usuario").strip()
        user = users_collection.find_one({
            "$or": [
                {"email": {"$regex": f"^{usuario_input}$", "$options": "i"}},
                {"nombre": {"$regex": f"^{usuario_input}$", "$options": "i"}}
            ]
        })
        if not user:
            return "No se encontró ningún usuario con ese nombre o email. <a href='/forgot-password'>Volver</a>"
        token = secrets.token_urlsafe(32)
        expires_at = datetime.utcnow() + timedelta(minutes=30)
        resets_collection.insert_one({
            "user_id": user["_id"],
            "token": token,
            "expires_at": expires_at,
            "used": False
        })
        reset_link = url_for("reset_password", token=token, _external=True)
        msg = Message("Recuperación de contraseña", recipients=[user["email"]])
        msg.body = (f"Hola {user['nombre']},\n\nPara restablecer tu contraseña, haz clic en el siguiente enlace:\n"
                    f"{reset_link}\n\nEste enlace caduca en 30 minutos.")
        mail.send(msg)
        return "Se ha enviado un enlace de recuperación a tu email. <a href='/login'>Inicia Sesión</a>"
    else:
        return render_template("forgot_password.html")

@app.route("/reset-password", methods=["GET", "POST"])
def reset_password():
    token = request.args.get("token") or request.form.get("token")
    if not token:
        return "Token no proporcionado."
    reset_info = resets_collection.find_one({"token": token})
    if not reset_info:
        return "Token inválido o inexistente."
    if reset_info["used"]:
        return "Este token ya ha sido utilizado."
    if datetime.utcnow() > reset_info["expires_at"]:
        return "Token caducado."
    if request.method == "POST":
        new_pass = request.form.get("password").strip()
        hashed = generate_password_hash(new_pass)
        user_id = reset_info["user_id"]
        users_collection.update_one({"_id": user_id}, {"$set": {"password": hashed}})
        resets_collection.update_one({"_id": reset_info["_id"]}, {"$set": {"used": True}})
        return "Contraseña actualizada con éxito. <a href='/login'>Inicia Sesión</a>"
    else:
        return render_template("reset_password_form.html", token=token)

# -------------------------------------------
# RUTAS PARA GESTIÓN DE TABLAS (SpreadSheets)
# -------------------------------------------
@app.route("/")
def home():
    if "usuario" not in session:
        return redirect(url_for("login"))
    if "selected_table" in session:
        return redirect(url_for("catalog"))
    else:
        return redirect(url_for("tables"))

@app.route("/tables", methods=["GET", "POST"])
def tables():
    if "usuario" not in session:
        return redirect(url_for("login"))
    owner = session["usuario"]
    if request.method == "POST":
        table_name = request.form.get("table_name").strip()
        import_file = request.files.get("import_table")
        if import_file and import_file.filename != "":
            filename = secure_filename(import_file.filename)
            filepath = os.path.join(SPREADSHEET_FOLDER, filename)
            import_file.save(filepath)
        else:
            file_id = secrets.token_hex(8)
            filename = f"table_{file_id}.xlsx"
            filepath = os.path.join(SPREADSHEET_FOLDER, filename)
            wb = Workbook()
            hoja = wb.active
            hoja.title = "Datos"
            hoja["A1"] = "Número"
            hoja["B1"] = "Descripción"
            hoja["C1"] = "Peso"
            hoja["D1"] = "Valor"
            hoja["E1"] = "Imagen1 (Ruta)"
            hoja["F1"] = "Imagen2 (Ruta)"
            hoja["G1"] = "Imagen3 (Ruta)"
            wb.save(filepath)
            wb.close()
        spreadsheets_collection.insert_one({
            "owner": owner,
            "name": table_name,
            "filename": filename,
            "created_at": datetime.utcnow()
        })
        return redirect(url_for("tables"))
    else:
        tables = list(spreadsheets_collection.find({"owner": session["usuario"]}))
        return render_template("tables.html", tables=tables)

@app.route("/select_table/<table_id>")
def select_table(table_id):
    if "usuario" not in session:
        return redirect(url_for("login"))
    table = spreadsheets_collection.find_one({"_id": ObjectId(table_id)})
    if not table:
        return "Tabla no encontrada."
    session["selected_table"] = table["filename"]
    return redirect(url_for("catalog"))

# -------------------------------------------
# RUTAS DEL CATÁLOGO (Excel e imágenes) para la tabla seleccionada
# -------------------------------------------
@app.route("/catalog", methods=["GET", "POST"])
def catalog():
    if "usuario" not in session:
        return redirect(url_for("login"))
    spreadsheet_path = get_current_spreadsheet()
    if not spreadsheet_path or not os.path.exists(spreadsheet_path):
        return redirect(url_for("tables"))
    if request.method == "POST":
        data = leer_datos_excel(spreadsheet_path)
        numero = request.form.get("numero").strip()
        descripcion = request.form.get("descripcion").strip()
        peso = request.form.get("peso")
        valor = request.form.get("valor")
        if any(item["numero"] == numero for item in data):
            return render_template("index.html", data=data, error_message="Error: Ese número ya existe.")
        files = request.files.getlist("imagenes")
        rutas_imagenes = [None, None, None]
        for i, file in enumerate(files[:3]):
            if file and allowed_file(file.filename):
                fname = secure_filename(file.filename)
                fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
                file.save(fpath)
                rutas_imagenes[i] = os.path.join("imagenes_subidas", fname)
        nuevo_registro = {
            "numero": numero,
            "descripcion": descripcion,
            "peso": peso,
            "valor": valor,
            "imagenes": rutas_imagenes
        }
        data.append(nuevo_registro)
        escribir_datos_excel(data, spreadsheet_path)
        return render_template("index.html", data=data, error_message="Registro agregado con éxito.")
    else:
        data = leer_datos_excel(get_current_spreadsheet())
        return render_template("index.html", data=data)

@app.route("/editar/<numero>", methods=["GET", "POST"])
def editar(numero):
    if "usuario" not in session:
        return redirect(url_for("login"))
    spreadsheet_path = get_current_spreadsheet()
    if not spreadsheet_path or not os.path.exists(spreadsheet_path):
        return redirect(url_for("tables"))
    if request.method == "GET":
        data = leer_datos_excel(spreadsheet_path)
        registro = next((item for item in data if item["numero"] == str(numero)), None)
        if not registro:
            return f"No existe el número {numero}."
        return render_template("editar.html", registro=registro)
    else:
        data = leer_datos_excel(spreadsheet_path)
        idx_encontrado = None
        for i, item in enumerate(data):
            if item["numero"] == str(numero):
                idx_encontrado = i
                break
        if idx_encontrado is None:
            return f"No existe el número {numero}."
        delete_flag = request.form.get("delete_record")
        if delete_flag == "on":
            data.pop(idx_encontrado)
            escribir_datos_excel(data, spreadsheet_path)
            return redirect(url_for("catalog"))
        else:
            nueva_descripcion = request.form.get("descripcion").strip()
            nuevo_peso = request.form.get("peso")
            nuevo_valor = request.form.get("valor")
            files = request.files.getlist("imagenes")
            rutas_imagenes = [None, None, None]
            for i, file in enumerate(files[:3]):
                if file and allowed_file(file.filename):
                    fname = secure_filename(file.filename)
                    fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
                    file.save(fpath)
                    rutas_imagenes[i] = os.path.join("imagenes_subidas", fname)
            remove_flags = [
                request.form.get("remove_img1"),
                request.form.get("remove_img2"),
                request.form.get("remove_img3")
            ]
            data[idx_encontrado]["descripcion"] = nueva_descripcion
            data[idx_encontrado]["peso"] = nuevo_peso
            data[idx_encontrado]["valor"] = nuevo_valor
            for i in range(3):
                if rutas_imagenes[i]:
                    data[idx_encontrado]["imagenes"][i] = rutas_imagenes[i]
                elif remove_flags[i] == "on":
                    data[idx_encontrado]["imagenes"][i] = None
            escribir_datos_excel(data, spreadsheet_path)
            return redirect(url_for("catalog"))

@app.route("/descargar-excel")
def descargar_excel():
    if "usuario" not in session:
        return redirect(url_for("login"))
    spreadsheet_path = get_current_spreadsheet()
    if not spreadsheet_path or not os.path.exists(spreadsheet_path):
        return "El Excel no existe aún."
    
    # Crear un archivo ZIP temporal que incluya el Excel y las imágenes referenciadas
    temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(temp_zip.name, "w") as zf:
        # Agregar el archivo Excel
        zf.write(spreadsheet_path, arcname=os.path.basename(spreadsheet_path))
        # Recopilar rutas únicas de imágenes
        data = leer_datos_excel(spreadsheet_path)
        image_paths = set()
        for row in data:
            for ruta in row["imagenes"]:
                if ruta:
                    absolute_path = os.path.join(app.root_path, ruta)
                    if os.path.exists(absolute_path):
                        image_paths.add(absolute_path)
        # Agregar las imágenes al ZIP, en la subcarpeta "imagenes"
        for img_path in image_paths:
            arcname = os.path.join("imagenes", os.path.basename(img_path))
            zf.write(img_path, arcname=arcname)
    return send_from_directory(directory=os.path.dirname(temp_zip.name),
                               path=os.path.basename(temp_zip.name),
                               as_attachment=True,
                               download_name="catalogo.zip")

@app.route("/imagenes_subidas/<path:filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

# -------------------------------------------
# MAIN
# -------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)