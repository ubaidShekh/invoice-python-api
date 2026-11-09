
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from pymongo import MongoClient
from pymongo.errors import PyMongoError

from dotenv import load_dotenv
from pydantic import BaseModel

from datetime import datetime, timezone
from io import BytesIO

import os
import uvicorn
import uuid
import fitz
import re

from openpyxl import Workbook


# =====================================================
# ENVIRONMENT
# =====================================================

load_dotenv()

app = FastAPI()


# =====================================================
# CORS
# =====================================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =====================================================
# MONGODB CONNECTION
# =====================================================

MONGODB_URI = os.getenv("MONGODB_URI")

if not MONGODB_URI:
    raise RuntimeError("MONGODB_URI is missing from .env")

client = MongoClient(
    MONGODB_URI,
    serverSelectionTimeoutMS=5000
)

db = client["invoice_delivery"]

invoices_collection = db["invoices"]
delivery_collection = db["delivery_items"]


# =====================================================
# PYDANTIC MODELS
# =====================================================

class DeliveryItemCreate(BaseModel):
    invoice: str
    model: str
    size: str
    barcode: str


class BarcodeScan(BaseModel):
    barcode: str


# =====================================================
# HELPER FUNCTIONS
# =====================================================

def generate_barcode():
    return "INV-" + uuid.uuid4().hex[:12].upper()


def get_invoice_number(text: str, filename: str):

    patterns = [
        r"Invoice\s*(?:No|Number|#)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        r"Inv\s*(?:No|Number|#)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        r"Bill\s*(?:No|Number|#)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return os.path.splitext(filename)[0]


def create_delivery_item_from_upload(
    invoice_number: str,
    filename: str
):

    new_item = {
        "invoice": invoice_number,
        "model": os.path.splitext(filename)[0],
        "size": "N/A",
        "barcode": generate_barcode(),
        "status": "Pending",
        "delivered_at": None,
        "created_at": datetime.now(timezone.utc)
    }

    result = delivery_collection.insert_one(new_item)

    return {
        "id": str(result.inserted_id),
        "invoice": new_item["invoice"],
        "model": new_item["model"],
        "size": new_item["size"],
        "barcode": new_item["barcode"],
        "status": new_item["status"]
    }


def get_report_items(status: str = "all"):

    status = status.lower().strip()

    if status not in [
        "all",
        "delivered",
        "undelivered"
    ]:
        raise HTTPException(
            status_code=400,
            detail="Invalid report status"
        )

    query = {}

    if status == "delivered":
        query["status"] = "Delivered"

    elif status == "undelivered":
        query["status"] = "Pending"

    items = list(
        delivery_collection.find(query).sort(
            "created_at",
            -1
        )
    )

    report_items = []

    for item in items:

        delivered_at = item.get("delivered_at")

        if delivered_at:
            delivered_at = delivered_at.isoformat()
        else:
            delivered_at = "-"

        report_items.append({
            "id": str(item["_id"]),
            "invoice": item.get("invoice", ""),
            "model": item.get("model", ""),
            "size": item.get("size", ""),
            "barcode": item.get("barcode", ""),
            "status": item.get("status", "Pending"),
            "deliveredAt": delivered_at
        })

    return report_items


# =====================================================
# HOME API
# =====================================================

@app.get("/")
def home():

    return {
        "message": "Python API is running"
    }


# =====================================================
# UPLOAD PDF API
# =====================================================

@app.post("/upload_pdf")
async def upload_pdf(invoice: UploadFile = File(...)):

    if not invoice.filename:
        raise HTTPException(
            status_code=400,
            detail="Filename is required"
        )

    if not invoice.filename.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Only PDF files are allowed"
        )

    pdf_data = await invoice.read()

    if not pdf_data:
        raise HTTPException(
            status_code=400,
            detail="Uploaded PDF is empty"
        )

    os.makedirs("uploads", exist_ok=True)

    invoice_id = str(uuid.uuid4())

    file_path = os.path.join(
        "uploads",
        f"{invoice_id}.pdf"
    )

    try:

        # Save original PDF
        with open(file_path, "wb") as file:
            file.write(pdf_data)

        # Extract PDF text
        pdf_document = fitz.open(
            stream=pdf_data,
            filetype="pdf"
        )

        extracted_pages = []

        for page_number, page in enumerate(
            pdf_document,
            start=1
        ):

            page_text = page.get_text("text")

            extracted_pages.append({
                "page_number": page_number,
                "text": page_text
            })

        extracted_text = "\n".join(
            page["text"]
            for page in extracted_pages
        )

        total_pages = len(extracted_pages)

        pdf_document.close()

        # Extract Invoice Number
        invoice_number = get_invoice_number(
            extracted_text,
            invoice.filename
        )

        # Prepare Invoice MongoDB Document
        invoice_document = {
            "invoice_id": invoice_id,
            "invoice_number": invoice_number,
            "filename": invoice.filename,
            "file_path": file_path,
            "content_type": invoice.content_type,
            "file_size": len(pdf_data),
            "total_pages": total_pages,
            "extracted_text": extracted_text,
            "pages": extracted_pages,
            "created_at": datetime.now(timezone.utc)
        }

        # Save Invoice
        invoice_result = invoices_collection.insert_one(
            invoice_document
        )

        # Create Delivery Item Automatically
        delivery_item = create_delivery_item_from_upload(
            invoice_number,
            invoice.filename
        )

        print("Invoice MongoDB ID:", invoice_result.inserted_id)
        print("Delivery Item:", delivery_item)

        return {
            "message": "PDF uploaded successfully",
            "invoice_id": invoice_id,
            "mongodb_id": str(invoice_result.inserted_id),
            "filename": invoice.filename,
            "invoice_number": invoice_number,
            "total_pages": total_pages,
            "delivery_item": delivery_item,
            "extracted_text": extracted_text
        }

    except HTTPException:
        raise

    except Exception as error:

        print("PDF Processing Error:", str(error))

        raise HTTPException(
            status_code=500,
            detail=f"PDF processing failed: {str(error)}"
        )


# =====================================================
# GET ALL INVOICES API
# =====================================================

@app.get("/api/invoices")
def get_invoices():

    try:

        invoices = list(
            invoices_collection.find().sort(
                "created_at",
                -1
            )
        )

        for invoice in invoices:

            invoice["_id"] = str(invoice["_id"])

            if invoice.get("created_at"):
                invoice["date"] = (
                    invoice["created_at"].isoformat()
                )
            else:
                invoice["date"] = ""

        return invoices

    except PyMongoError as error:

        print("MongoDB Fetch Error:", str(error))

        raise HTTPException(
            status_code=500,
            detail="Failed to fetch invoices"
        )


# =====================================================
# GET ALL DELIVERY ITEMS API
# =====================================================

@app.get("/api/delivery-items")
def get_delivery_items():

    try:

        items = list(
            delivery_collection.find().sort(
                "created_at",
                -1
            )
        )

        for item in items:

            item["_id"] = str(item["_id"])

            if item.get("delivered_at"):
                item["date"] = (
                    item["delivered_at"].isoformat()
                )
            else:
                item["date"] = "-"

            if item.get("created_at"):
                item["created_at"] = (
                    item["created_at"].isoformat()
                )

        return items

    except PyMongoError as error:

        print("Delivery Fetch Error:", str(error))

        raise HTTPException(
            status_code=500,
            detail="Failed to fetch delivery items"
        )


# =====================================================
# CREATE DELIVERY ITEM API
# =====================================================

@app.post("/api/delivery-items")
def create_delivery_item_api(item: DeliveryItemCreate):

    barcode = item.barcode.strip()

    if not barcode:
        raise HTTPException(
            status_code=400,
            detail="Barcode is required"
        )

    existing_item = delivery_collection.find_one({
        "barcode": barcode
    })

    if existing_item:
        raise HTTPException(
            status_code=409,
            detail="Barcode already exists"
        )

    new_item = {
        "invoice": item.invoice.strip(),
        "model": item.model.strip(),
        "size": item.size.strip(),
        "barcode": barcode,
        "status": "Pending",
        "delivered_at": None,
        "created_at": datetime.now(timezone.utc)
    }

    result = delivery_collection.insert_one(new_item)

    return {
        "message": "Delivery item created successfully",
        "id": str(result.inserted_id)
    }


# =====================================================
# BARCODE SCAN API
# =====================================================

@app.post("/api/delivery-items/scan")
def deliver_item(data: BarcodeScan):

    barcode = data.barcode.strip()

    if not barcode:
        raise HTTPException(
            status_code=400,
            detail="Barcode is required"
        )

    item = delivery_collection.find_one({
        "barcode": barcode
    })

    if not item:
        raise HTTPException(
            status_code=404,
            detail="Item not found"
        )

    if item.get("status") == "Delivered":
        raise HTTPException(
            status_code=409,
            detail="This item has already been delivered"
        )

    delivered_at = datetime.now(timezone.utc)

    delivery_collection.update_one(
        {"_id": item["_id"]},
        {
            "$set": {
                "status": "Delivered",
                "delivered_at": delivered_at
            }
        }
    )

    return {
        "message": "Item delivered successfully",
        "barcode": barcode,
        "status": "Delivered",
        "date": delivered_at.isoformat()
    }


# =====================================================
# REPORTS API
# =====================================================

@app.get("/api/reports")
def get_reports(status: str = "all"):

    try:

        items = get_report_items(status)

        return {
            "status": status,
            "total": len(items),
            "items": items
        }

    except HTTPException:
        raise

    except PyMongoError as error:

        print("Reports Fetch Error:", str(error))

        raise HTTPException(
            status_code=500,
            detail="Failed to fetch reports"
        )


# =====================================================
# EXPORT REPORTS TO EXCEL API
# =====================================================

@app.get("/api/invoices/export")
def export_reports(status: str = "all"):

    try:

        # Fetch Filtered Items
        items = get_report_items(status)

        # Create Workbook
        workbook = Workbook()

        worksheet = workbook.active

        worksheet.title = "Delivery Report"

        # Excel Header
        worksheet.append([
            "Invoice",
            "Model",
            "Size",
            "Barcode",
            "Status",
            "Delivered At"
        ])

        # Excel Data
        for item in items:

            worksheet.append([
                item["invoice"],
                item["model"],
                item["size"],
                item["barcode"],
                item["status"],
                item["deliveredAt"]
            ])

        # Column Widths
        worksheet.column_dimensions["A"].width = 20
        worksheet.column_dimensions["B"].width = 15
        worksheet.column_dimensions["C"].width = 12
        worksheet.column_dimensions["D"].width = 28
        worksheet.column_dimensions["E"].width = 18
        worksheet.column_dimensions["F"].width = 30

        # Save Excel to Memory
        excel_file = BytesIO()

        workbook.save(excel_file)

        excel_file.seek(0)

        filename = f"{status}-report.xlsx"

        return StreamingResponse(
            excel_file,
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": (
                    f'attachment; filename="{filename}"'
                )
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print("Excel Export Error:", str(error))

        raise HTTPException(
            status_code=500,
            detail="Excel export failed"
        )


# =====================================================
# SERVER
# =====================================================

if __name__ == "__main__":

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 3000))
    )
