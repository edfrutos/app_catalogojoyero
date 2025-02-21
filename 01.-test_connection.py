import certifi
from pymongo import MongoClient

# Conexión a MongoDB Atlas
MONGO_URI = "mongodb+srv://edfrutos:rYjwUC6pUNrLtbaI@cluster0.pmokh.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"

client = MongoClient(
    MONGO_URI,
    tls=True,
    tlsCAFile=certifi.where(),
    serverSelectionTimeoutMS=30000
)

# Probar la conexión
try:
    client.admin.command('ping')
    print("Conexión exitosa a MongoDB Atlas")
except Exception as e:
    print("Error conectando a MongoDB:", e)