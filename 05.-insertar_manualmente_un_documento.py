from pymongo import MongoClient

client = MongoClient("mongodb://localhost:27017/")
db = client["app_catalogojoyero"]

# Verifica que la colección existe
collection_name = "67b8c24a7fdc72dd4d8703cf"  # Asegúrate de que es la correcta
collection = db[collection_name]

# Insertar un documento de prueba
test_doc = {"test": "MongoDB funcionando"}
collection.insert_one(test_doc)

# Comprobar si el documento se insertó
print("Documento insertado:", collection.find_one({"test": "MongoDB funcionando"}))