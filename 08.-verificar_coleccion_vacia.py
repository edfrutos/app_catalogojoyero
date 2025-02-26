from pymongo import MongoClient

# Conectar a MongoDB
client = MongoClient("mongodb://localhost:27017/")
db = client["app_catalogojoyero"]

# Verificar si la colección existe
print("Colecciones disponibles en MongoDB:", db.list_collection_names())

# Intentar leer documentos
collection_name = "67b8c24a7fdc72dd4d8703cf"  # Reemplaza con la colección correcta
if collection_name in db.list_collection_names():
    catalog_collection = db[collection_name]
    registros = list(catalog_collection.find())
    print(f"📌 Documentos en {collection_name}:", registros)
else:
    print(f"❌ La colección {collection_name} no existe en MongoDB.")