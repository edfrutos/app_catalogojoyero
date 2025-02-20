import os
import certifi
import secrets
from datetime import datetime, timedelta

from flask import (
    Flask, request, render_template, redirect, url_for,
    send_from_directory, session, flash
)
import openpyxl
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from pymongo import MongoClient
from flask_mail import Mail, Message

# -------------------------------------------
# CONFIGURACIÓN FLASK
# -------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "CAMBIA_ESTA_CLAVE_EN_PRODUCCION")

EXCEL_FILE = "datos.xlsx"

# Carpeta para imágenes subidas (para el catálogo de joyas)
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "imagenes_subidas")
if not os.path.exists(app.config["UPLOAD_FOLDER"]):
    os.makedirs(app.config["UPLOAD_FOLDER"])

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
MONGO_URI = "mongodb+srv://edfrutos:rYjwUC6pUNrLtbaI@cluster0.pmokh.mongodb.net/"
client = MongoClient(MONGO_URI, tlsCAFile=certifi.where())
db = client["app_catalogojoyero"]
users_collection = db["users"]
resets_collection = db["password_resets"]

# -------------------------------------------
# MANEJO DE EXCEL - FUNCIONES AUXILIARES
# -------------------------------------------
def leer_datos_excel():
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE, read_only=False)
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

def escribir_datos_excel(data):
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
    wb.save(EXCEL_FILE)
    wb.close()

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
        # Usamos regex para comparación case-insensitive
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
            return redirect(url_for("index"))
        else:
            return "Error: Contraseña incorrecta. <a href='/login'>Reintentar</a>"
    else:
        return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# Redirección de /recover a /forgot-password
@app.route("/recover")
def recover_redirect():
    return redirect(url_for("forgot_password"))

@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        usuario_input = request.form.get("usuario").strip()
        # Buscar por email o nombre de forma case-insensitive
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
# RUTAS DEL CATÁLOGO (Excel e imágenes)
# -------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if "usuario" not in session:
        return redirect(url_for("login"))
    if request.method == "POST":
        data = leer_datos_excel()
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
        return render_template("index.html", data=data, error_message="Registro agregado con éxito.")
    else:
        data = leer_datos_excel()
        return render_template("index.html", data=data)

@app.route("/editar/<numero>", methods=["GET", "POST"])
def editar(numero):
    if "usuario" not in session:
        return redirect(url_for("login"))
    if request.method == "GET":
        data = leer_datos_excel()
        registro = next((item for item in data if item["numero"] == str(numero)), None)
        if not registro:
            return f"No existe el número {numero}."
        return render_template("editar.html", registro=registro)
    else:
        data = leer_datos_excel()
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
            escribir_datos_excel(data)
            return redirect(url_for("index"))
        else:
            nueva_descripcion = request.form.get("descripcion").strip()
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
            for i in range(3):
                if rutas_imagenes[i]:
                    data[idx_encontrado]["imagenes"][i] = rutas_imagenes[i]
                elif remove_flags[i] == "on":
                    data[idx_encontrado]["imagenes"][i] = None
            escribir_datos_excel(data)
            return redirect(url_for("index"))

@app.route("/descargar-excel")
def descargar_excel():
    if "usuario" not in session:
        return redirect(url_for("login"))
    if not os.path.exists(EXCEL_FILE):
        return "El Excel no existe aún."
    return send_from_directory(app.root_path, EXCEL_FILE, as_attachment=True)

@app.route("/imagenes_subidas/<path:filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

if __name__ == "__main__":
    app.run(debug=True)