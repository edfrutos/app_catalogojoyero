# update_catalogs.py
from pymongo import MongoClient
from dotenv import load_dotenv
import os

# Cargar variables de entorno
load_dotenv()

# Obtener la URL de conexión desde las variables de entorno
MONGO_URI = os.getenv('MONGO_URI', 'mongodb://localhost:27017/')

def update_unnamed_catalogs():
    try:
        # Conectar a MongoDB
        client = MongoClient(MONGO_URI)
        db = client.catalogojoyero  # Reemplaza con el nombre de tu base de datos
        spreadsheets_collection = db.spreadsheets  # Reemplaza con el nombre de tu colección

        # Encontrar catálogos sin nombre
        unnamed_catalogs = spreadsheets_collection.find({"name": {"$exists": False}})
        count = 0

        # Actualizar cada catálogo
        for catalog in unnamed_catalogs:
            result = spreadsheets_collection.update_one(
                {"_id": catalog["_id"]},
                {"$set": {"name": f"Catálogo {catalog['filename']}"}}
            )
            if result.modified_count > 0:
                count += 1
                print(f"Actualizado: Catálogo {catalog['filename']}")

        print(f"\nActualización completada. {count} catálogos actualizados.")

    except Exception as e:
        print(f"Error durante la actualización: {str(e)}")
    finally:
        client.close()

if __name__ == "__main__":
    print("Iniciando actualización de catálogos sin nombre...")
    update_unnamed_catalogs()
