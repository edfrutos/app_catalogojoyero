from pymongo import MongoClient

client = MongoClient("mongodb://localhost:27017/")
db = client["app_catalogojoyero"]
collection = db["spreadsheets"]

for doc in collection.find():
    print(doc)