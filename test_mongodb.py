# -------------------------------------------
# CONEXIÓN A MONGODB ATLAS
# -------------------------------------------
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
import certifi  # Agregar esta línea para evitar errores de SSL/TLS

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