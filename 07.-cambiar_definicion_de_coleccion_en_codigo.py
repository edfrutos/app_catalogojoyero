from pymongo import MongoClient

# Conectar a MongoDB
client = MongoClient("mongodb://localhost:27017/")
db = client["app_catalogojoyero"]  # Asegurar que la base de datos es correcta

# Definir la colección del catálogo
catalog_collection = db["67b8c24a7fdc72dd4d8703cf"]  # Nombre correcto de la colección

# Imprimir las colecciones para verificar
print("Colecciones disponibles en MongoDB:", db.list_collection_names())

# Imprimir el nombre de la colección usada
print(f"Nombre de la colección utilizada: {catalog_collection.name}")