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

# Forzar el uso del bundle de certificados de certifi
os.environ['SSL_CERT_FILE'] = certifi.where()

# -------------------------------------------
# CONFIGURACIÓN FLASK
# -------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("MBZl1W45ute3UEMCXPlL9JzcR7XsTeUi-4ZI6KCd79M", "CAMBIA_ESTA_CLAVE_EN_PRODUCCION")

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
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi

MONGO_URI = "mongodb+srv://edfrutos:rYjwUC6pUNrLtbaI@cluster0.pmokh.mongodb.net/?retryWrites=true&w=majority"

# Crear la conexión a MongoDB Atlas
client = MongoClient(
    MONGO_URI,
    tls=True,
    tlsCAFile=certifi.where(),
    server_api=ServerApi('1')
)

try:
    client.admin.command('ping')
    print("✅ Pingó su implementación. ¡Te conectaste con éxito a MongoDB!")
except Exception as e:
    print("❌ Error al conectar con MongoDB:", e)

# Conectar a la base de datos
db = client["app_catalogojoyero"]

# Mostrar colecciones disponibles
print("📌 Colecciones disponibles en MongoDB:", db.list_collection_names())

# Definir la colección específica del catálogo
catalog_collection = db["67b8c24a7fdc72dd4d8703cf"]  # Asegúrate de que el nombre es el correcto
users_collection = db["users"]
resets_collection = db["password_resets"]
spreadsheets_collection = db["spreadsheets"]

# Verificar si la colección tiene datos
registros = list(catalog_collection.find())

print("📌 Documentos en la colección:")
for doc in registros:
    print(doc)

print(f"📌 Total de registros obtenidos: {len(registros)}")

# Insertar un documento de prueba si la colección está vacía
if len(registros) == 0:
    doc_prueba = {"test": "Conexión funcionando"}
    catalog_collection.insert_one(doc_prueba)
    print("✅ Se insertó un documento de prueba.")

# Verificar que el documento de prueba se insertó correctamente
print("📌 Registros después de la prueba:", list(catalog_collection.find()))

# -------------------------------------------
# FUNCIONES AUXILIARES PARA HOJAS DE CÁLCULO
# -------------------------------------------
def leer_datos_excel(filename):
    if not os.path.exists(filename):
        return []

    wb = load_workbook(filename, read_only=True)
    hoja = wb.active
    data = []

    headers = [cell.value for cell in hoja[1]]

    for row in hoja.iter_rows(min_row=2, values_only=True):
        registro = {headers[i]: row[i] for i in range(len(headers))}

        # Asegurar que 'Imagenes' sea una lista
        if "Imagenes" in registro and isinstance(registro["Imagenes"], str):
            registro["Imagenes"] = registro["Imagenes"].split(", ")
        elif "Imagenes" in registro:
            registro["Imagenes"] = []

        data.append(registro)

    wb.close()
    return data

def escribir_datos_excel(data, filename):
    wb = Workbook()
    hoja = wb.active
    hoja.title = "Datos"

    headers = session.get("selected_headers", ["Número", "Descripción", "Peso", "Valor", "Imagenes"])

    if "Número" not in headers:
        headers.insert(0, "Número")

    hoja.append(headers)

    for item in data:
        fila = []
        for header in headers:
            valor = item.get(header, "")

            if header == "Imagenes" and isinstance(valor, list):
                valor = ", ".join(valor)

            fila.append(valor)

        hoja.append(fila)

    wb.save(filename)
    wb.close()

def get_current_spreadsheet():
    filename = session.get("selected_table")
    if not filename:
        return None
    return os.path.join(SPREADSHEET_FOLDER, filename)

# -------------------------------------------
# NUEVA RUTA: PÁGINA DE BIENVENIDA (MANUAL DE USO)
# -------------------------------------------
@app.route("/welcome")
def welcome():
    # Si el usuario no ha iniciado sesión, redirige al login
    if "usuario" not in session:
        return redirect(url_for("login"))
    return render_template("welcome.html")

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
        return render_template("welcome.html")  # Mostrar página de bienvenida si no hay sesión iniciada

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
        table_name = request.form.get("table_name", "").strip()
        import_file = request.files.get("import_table")

        if import_file and import_file.filename != "":
            # Importar el Excel
            filename = secure_filename(import_file.filename)
            filepath = os.path.join(SPREADSHEET_FOLDER, filename)
            import_file.save(filepath)
            
            # Leer los encabezados del Excel importado
            wb = openpyxl.load_workbook(filepath)
            hoja = wb.active
            headers = next(hoja.iter_rows(min_row=1, max_row=1, values_only=True))
            wb.close()
            
            # Verificar que los encabezados no estén vacíos
            if not headers or all(header is None for header in headers):
                flash("El archivo Excel importado no contiene encabezados válidos.", "error")
                return redirect(url_for("tables"))

            # Guardar info en MongoDB
            spreadsheets_collection.insert_one({
                "owner": owner,
                "name": table_name,
                "filename": filename,
                "headers": headers,
                "created_at": datetime.utcnow()
            })
        else:
            # Caso en que no sube archivo (creas uno nuevo con encabezados)
            headers_str = request.form.get("table_headers", "").strip()
            if not headers_str:
                headers = ["Número", "Descripción", "Peso", "Valor"]  # Por defecto
            else:
                headers = [h.strip() for h in headers_str.split(",") if h.strip()]

            # Verificar que los encabezados no estén vacíos
            if not headers:
                flash("Debe proporcionar al menos un encabezado válido.", "error")
                return redirect(url_for("tables"))
            
            file_id = secrets.token_hex(8)
            filename = f"table_{file_id}.xlsx"
            filepath = os.path.join(SPREADSHEET_FOLDER, filename)
            
            wb = Workbook()
            hoja = wb.active
            hoja.append(headers)
            wb.save(filepath)
            wb.close()

            # Guardar info en MongoDB
            spreadsheets_collection.insert_one({
                "owner": owner,
                "name": table_name,
                "filename": filename,
                "headers": headers,
                "created_at": datetime.utcnow()
            })
        
        # Almacenar los encabezados en la sesión para su uso posterior
        session["selected_headers"] = headers
        return redirect(url_for("tables"))

    # GET: mostrar las tablas existentes
    todas_las_tablas = list(spreadsheets_collection.find({"owner": session["usuario"]}))
    return render_template("tables.html", tables=todas_las_tablas)

@app.route("/select_table/<table_id>")
def select_table(table_id):
    if "usuario" not in session:
        return redirect(url_for("login"))

    table = spreadsheets_collection.find_one({"_id": ObjectId(table_id)})

    if not table:
        flash("Tabla no encontrada.", "error")
        return redirect(url_for("tables"))

    session["selected_table"] = table["filename"]
    session["selected_table_id"] = str(table["_id"])  # Almacenar el ID de la tabla
    session["selected_table_name"] = table["name"]  # Almacenar el nombre de la tabla

    return redirect(url_for("catalog"))


# -------------------------------------------
# RUTAS DEL CATÁLOGO (Excel e imágenes) para la tabla seleccionada
# -------------------------------------------

@app.route("/catalog", methods=["GET", "POST"])
def catalog():
    if "usuario" not in session:
        return redirect(url_for("welcome"))

    if "selected_table" not in session:
        flash("Por favor, seleccione una tabla primero.", "warning")
        return redirect(url_for("tables"))

    selected_table = session["selected_table"]
    table_info = spreadsheets_collection.find_one({"filename": selected_table})

    if not table_info:
        flash("La tabla seleccionada no existe.", "error")
        return redirect(url_for("tables"))

    headers = table_info.get("headers", [])
    if not headers:
        flash("La tabla seleccionada no tiene encabezados definidos.", "error")
        return redirect(url_for("tables"))

    # Guardar los encabezados en la sesión
    session["selected_headers"] = headers

    # Obtener registros de MongoDB para la tabla seleccionada
    registros = list(catalog_collection.find({"table": selected_table}))

    if request.method == "POST":
        form_data = {k.strip(): v.strip() for k, v in request.form.items()}

        # Verificar si existe "Número" o el primer encabezado como identificador
        id_field = headers[0]
        if id_field not in form_data or not form_data[id_field]:
            return render_template("index.html", data=registros, headers=headers, error_message=f"Error: Sin {id_field}.")

        # Verificar si el identificador ya existe en esta tabla
        if any(item.get(id_field) == form_data[id_field] for item in registros):
            return render_template("index.html", data=registros, headers=headers, error_message=f"Error: Ese {id_field} ya existe.")

        # Construir el nuevo registro
        nuevo_registro = {
            "Número": len(registros) + 1,
            "table": selected_table
        }
        
        # Agregar los campos del formulario
        for header in headers:
            nuevo_registro[header] = form_data.get(header, "").strip()

        # Manejo de imágenes
        files = request.files.getlist("imagenes")
        rutas_imagenes = [None, None, None]

        for i, file in enumerate(files[:3]):
            if file and allowed_file(file.filename):
                fname = secure_filename(file.filename)
                fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
                file.save(fpath)
                rutas_imagenes[i] = f"/imagenes_subidas/{fname}"

        nuevo_registro["Imagenes"] = rutas_imagenes

        # Insertar en MongoDB
        catalog_collection.insert_one(nuevo_registro)

        return redirect(url_for("catalog"))

    # GET: mostrar los registros existentes
    return render_template("index.html", data=registros, headers=headers)

@app.route("/editar/<id>", methods=["GET", "POST"])
def editar(id):
    if "usuario" not in session:
        return redirect(url_for("login"))

    if "selected_table" not in session:
        flash("Por favor, seleccione una tabla primero.", "warning")
        return redirect(url_for("tables"))

    selected_table = session["selected_table"]
    table_info = spreadsheets_collection.find_one({"filename": selected_table})

    if not table_info:
        flash("La tabla seleccionada no existe.", "error")
        return redirect(url_for("tables"))

    headers = table_info.get("headers", [])
    if not headers:
        flash("La tabla seleccionada no tiene encabezados definidos.", "error")
        return redirect(url_for("tables"))

    id_field = headers[0]

    # Obtenemos el registro desde MongoDB
    registro = catalog_collection.find_one({id_field: id, "table": selected_table})
    if not registro:
        flash(f"No existe el {id_field} {id} en la tabla seleccionada.", "error")
        return redirect(url_for("catalog"))

    if request.method == "GET":
        # Filtrar los headers para excluir 'Imagenes' del formulario principal
        headers_form = [h for h in headers if h != "Imagenes"]
        return render_template("editar.html", 
                             registro=registro, 
                             headers=headers_form,
                             imagenes_actuales=registro.get("Imagenes", [None, None, None]))

    # POST: Guardar cambios
    if request.form.get("delete_record") == "on":
        result = catalog_collection.delete_one({id_field: id, "table": selected_table})
        if result.deleted_count > 0:
            flash("Registro eliminado exitosamente.", "success")
        else:
            flash("No se pudo eliminar el registro.", "error")
        return redirect(url_for("catalog"))

    try:
        # Actualizar campos (excluyendo imágenes)
        update_data = {}
        
        # Manejar campos especiales primero
        update_data["Número"] = registro["Número"]  # Mantener el número original
        update_data["table"] = selected_table

        # Actualizar el resto de campos desde el formulario
        for header in headers:
            if header != "Imagenes" and header != "Número":
                form_value = request.form.get(header, "").strip()
                # Usar punto para campos con espacios en MongoDB
                safe_header = header.replace(" ", "_").replace(".", "_")
                update_data[safe_header] = form_value

        # Manejo de imágenes
        rutas_imagenes = registro.get("Imagenes", [None, None, None])
        files = request.files.getlist("imagenes")

        for i, file in enumerate(files[:3]):
            if file and file.filename and allowed_file(file.filename):
                fname = secure_filename(file.filename)
                fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
                file.save(fpath)
                rutas_imagenes[i] = f"/imagenes_subidas/{fname}"

        # Manejar eliminación de imágenes
        for i in range(3):
            if request.form.get(f"remove_img{i+1}") == "on":
                rutas_imagenes[i] = None

        update_data["Imagenes"] = rutas_imagenes

        # Actualizar en MongoDB usando replace_one en lugar de update_one
        result = catalog_collection.replace_one(
            {id_field: id, "table": selected_table},
            update_data
        )

        if result.modified_count > 0:
            flash("Registro actualizado exitosamente.", "success")
        else:
            flash("No se detectaron cambios en el registro.", "info")

    except Exception as e:
        flash(f"Error al actualizar el registro: {str(e)}", "error")
        print(f"Error en la actualización: {str(e)}")  # Para debugging

    return redirect(url_for("catalog"))

@app.route("/delete_table/<table_id>", methods=["POST"])
def delete_table(table_id):
    if "usuario" not in session:
        return redirect(url_for("login"))
    # Buscamos la tabla en la colección, asegurándonos que el usuario actual es el propietario.
    table = spreadsheets_collection.find_one({"_id": ObjectId(table_id), "owner": session["usuario"]})
    if not table:
        flash("Tabla no encontrada o no tienes permiso para eliminarla.", "error")
        return redirect(url_for("tables"))
    
    # Construir la ruta absoluta del archivo Excel
    file_path = os.path.join(SPREADSHEET_FOLDER, table["filename"])
    # Si el archivo existe, lo eliminamos del sistema de archivos.
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except Exception as e:
            flash(f"Error al eliminar el archivo: {e}", "error")
            return redirect(url_for("tables"))
    
    # Eliminamos el documento de la colección en MongoDB
    spreadsheets_collection.delete_one({"_id": ObjectId(table_id)})
    
    # Si la tabla eliminada era la seleccionada en sesión, la removemos de la sesión.
    if session.get("selected_table") == table["filename"]:
        session.pop("selected_table", None)
    
    flash("Tabla eliminada exitosamente.", "success")
    return redirect(url_for("tables"))

@app.route("/descargar-excel")
def descargar_excel():
    if "usuario" not in session:
        return redirect(url_for("login"))
    spreadsheet_path = get_current_spreadsheet()
    if not spreadsheet_path or not os.path.exists(spreadsheet_path):
        return "El Excel no existe aún."
    temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(temp_zip.name, "w") as zf:
        zf.write(spreadsheet_path, arcname=os.path.basename(spreadsheet_path))
        data = leer_datos_excel(spreadsheet_path)
        image_paths = set()
        for row in data:
            for ruta in row["imagenes"]:
                if ruta:
                    absolute_path = os.path.join(app.root_path, ruta)
                    if os.path.exists(absolute_path):
                        image_paths.add(absolute_path)
        for img_path in image_paths:
            arcname = os.path.join("imagenes", os.path.basename(img_path))
            zf.write(img_path, arcname=arcname)
    return send_from_directory(directory=os.path.dirname(temp_zip.name),
                               path=os.path.basename(temp_zip.name),
                               as_attachment=True,
                               download_name="catalogo.zip")

from flask import send_from_directory

@app.route("/imagenes_subidas/<filename>")
def uploaded_images(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

from bson import ObjectId

def convertir_registros(registros):
    for r in registros:
        if "_id" in r and isinstance(r["_id"], ObjectId):
            r["_id"] = str(r["_id"])
    return registros

@app.route("/debug_mongo")
def debug_mongo():
    print("📌 Verificando conexión con MongoDB en Flask")

    colecciones = db.list_collection_names()
    print("📌 Colecciones disponibles en MongoDB:", colecciones)

    documentos = list(catalog_collection.find())

    print("📌 Documentos en la colección correcta:")
    for doc in documentos:
        print(doc)

    # Convertir ObjectId a string para evitar error de serialización
    for doc in documentos:
        doc["_id"] = str(doc["_id"])

    return {"colecciones": colecciones, "documentos": documentos}

@app.route("/insert_test")
def insert_test():
    nuevo_registro = {
        "Número": 2,
        "Descripción": "Collar de prueba",
        "Peso": 10,
        "Valor": 2500,
        "Imagenes": ["/imagenes_subidas/collar_prueba.jpg"]
    }
    
    catalog_collection.insert_one(nuevo_registro)
    return "✅ Registro de prueba insertado en MongoDB"
# -------------------------------------------
# MAIN
# -------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
