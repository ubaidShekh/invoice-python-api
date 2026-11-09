
from fastapi import FastAPI
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from dotenv import load_dotenv
import os
import uvicorn

load_dotenv()

app = FastAPI()

# MongoDB Connection
MONGODB_URI = os.getenv("MONGODB_URI")

client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)
db = client["invoice_delivery"]
collection = db["invoices"]


@app.get("/")
def home():
    return {"message": "Python API is running"}


@app.get("/api")
def api():
    return {"message": "API working successfully"}


@app.get("/api/db")
def check_database():
    try:
        client.admin.command("ping")

        return {
            "message": "MongoDB connected successfully",
            "database": db.name,
            "status": "connected"
        }

    except PyMongoError as error:
        return {
            "message": "MongoDB connection failed",
            "error": str(error),
            "status": "disconnected"
        }


if __name__ == "__main__":
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 3000))
    )
