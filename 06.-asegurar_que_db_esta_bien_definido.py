from pymongo import MongoClient

# Conectar a MongoDB
client = MongoClient("mongodb://localhost:27017/")
db = client["app_catalogojoyero"]  # Asegurar que la base de datos es correcta

# Imprimir colecciones disponibles
print("Colecciones disponibles en MongoDB:", db.list_collection_names())