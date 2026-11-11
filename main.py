from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from dotenv import load_dotenv
from datetime import datetime, timezone
from openpyxl import Workbook
from io import BytesIO
from typing import Optional, List
import os
import re
import uuid
import fitz
import uvicorn

# =====================================================
# CONFIGURATION
# =====================================================
load_dotenv()

app = FastAPI(title="Invoice Management API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

MONGODB_URI = os.getenv("MONGODB_URI")
if not MONGODB_URI:
    raise RuntimeError("MONGODB_URI environment variable is missing")

client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=10000, connectTimeoutMS=10000)
db = client["invoice_delivery"]
invoices_collection = db["invoices"]
delivery_collection = db["delivery_items"]

invoices_collection.create_index("invoice_id")
delivery_collection.create_index("barcode", unique=True)
delivery_collection.create_index("invoice")

MAX_FILE_SIZE = 25 * 1024 * 1024
MAX_PAGES = 150

# =====================================================
# MODELS
# =====================================================
class DeliveryItemCreate(BaseModel):
    invoice: str
    model: str
    size: str
    barcode: str

class BarcodeScan(BaseModel):
    barcode: str

class DeliveryItemUpdate(BaseModel):
    invoice: Optional[str] = None
    model: Optional[str] = None
    size: Optional[str] = None

# =====================================================
# HELPERS
# =====================================================
def now_utc():
    return datetime.now(timezone.utc)

def generate_barcode():
    return "INV-" + uuid.uuid4().hex[:12].upper()

def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()

# --- NEW EXTRACTION LOGIC ---
ORDER_NUMBER_RE = re.compile(r"#(\d{7,12})")
MODEL_RE = re.compile(r"#(\d{2,4})\b")
PRODUCT_CODE_RE = re.compile(r"\b(\d{8,10})\b")
SIZE_RE = re.compile(r"^\d{2,3}$|^(XXS|XS|S|M|L|XL|XXL|XXXL)$", re.IGNORECASE)

def extract_invoices_from_pdf(full_text, filename):
    matches = list(ORDER_NUMBER_RE.finditer(full_text))
    if not matches:
        return [{"invoice_number": os.path.splitext(filename)[0], "text": full_text}]

    sections = []
    seen_numbers = set()

    for index, match in enumerate(matches):
        invoice_number = match.group(1)
        if invoice_number in seen_numbers:
            continue
        seen_numbers.add(invoice_number)

        start = match.start()
        next_match = None
        for following in matches[index + 1:]:
            if following.group(1) != invoice_number:
                next_match = following
                break

        end = next_match.start() if next_match else len(full_text)
        section_text = full_text[start:end].strip()

        if section_text:
            sections.append({
                "invoice_number": invoice_number,
                "text": section_text
            })

    return sections

def extract_products(section_text):
    products = []
    seen = set()
    lines = section_text.split('\n')
    
    for i, line in enumerate(lines):
        line = clean_text(line)
        
        # Find Model (e.g., #230)
        model_match = MODEL_RE.search(line)
        if model_match:
            model = "#" + model_match.group(1)
            product_code = ""
            
            # Product code might be on the same line or nearby
            code_match = PRODUCT_CODE_RE.search(line)
            if code_match:
                product_code = code_match.group(1)
            else:
                for j in range(max(0, i-2), min(i+3, len(lines))):
                    code_match = PRODUCT_CODE_RE.search(clean_text(lines[j]))
                    if code_match and code_match.group(1) != model_match.group(1):
                        product_code = code_match.group(1)
                        break

            # Size extraction: search current and next 5 lines
            size = "N/A"
            for j in range(i, min(i + 6, len(lines))):
                next_line = clean_text(lines[j])
                # Check if line is just a number (like 54)
                if SIZE_RE.fullmatch(next_line) and next_line != model_match.group(1) and next_line != product_code:
                    size = next_line.upper()
                    break
                # Check if line contains "Size" or Arabic "المقاس" and a number
                size_match = re.search(r"(?:size|المقاس)\s*[:]?\s*([0-9]{2,3}|XXS|XS|S|M|L|XL|XXL|XXXL)\b", next_line, re.IGNORECASE)
                if size_match:
                    size = size_match.group(1).upper()
                    break

            key = (model, size, product_code)
            if key not in seen:
                seen.add(key)
                products.append({
                    "model": model,
                    "size": size,
                    "product_code": product_code,
                    "confidence": 0.95,
                    "review_required": False
                })

    return products

# --- END NEW EXTRACTION LOGIC ---

def serialize_invoice(document):
    return {
        "_id": str(document["_id"]),
        "invoice_id": document.get("invoice_id", ""),
        "invoice_number": document.get("invoice_number", ""),
        "filename": document.get("filename", ""),
        "file_size": document.get("file_size", 0),
        "total_pages": document.get("total_pages", 0),
        "extracted_text": document.get("extracted_text", ""),
        "products": document.get("products", []),
        "review_required": document.get("review_required", False),
        "extraction_status": document.get("extraction_status", "completed"),
        "date": document["created_at"].isoformat() if document.get("created_at") else ""
    }

def serialize_delivery_item(item):
    delivered_at = item.get("delivered_at")
    return {
        "_id": str(item["_id"]),
        "invoice": item.get("invoice", ""),
        "model": item.get("model", ""),
        "size": item.get("size", "N/A"),
        "barcode": item.get("barcode", ""),
        "status": item.get("status", "Pending"),
        "delivered_at": delivered_at.isoformat() if delivered_at else None,
        "date": delivered_at.isoformat() if delivered_at else "-",
        "created_at": item["created_at"].isoformat() if item.get("created_at") else "",
        "review_required": item.get("review_required", False),
        "confidence": item.get("confidence", 0)
    }

def get_report_items(status="all"):
    status = status.lower().strip()
    if status not in ["all", "delivered", "undelivered"]:
        raise HTTPException(status_code=400, detail="Invalid report status")

    query = {}
    if status == "delivered":
        query["status"] = "Delivered"
    elif status == "undelivered":
        query["status"] = "Pending"

    items = delivery_collection.find(query).sort("created_at", -1)
    return [
        {
            "id": str(item["_id"]),
            "invoice": item.get("invoice", ""),
            "model": item.get("model", ""),
            "size": item.get("size", "N/A"),
            "barcode": item.get("barcode", ""),
            "status": item.get("status", "Pending"),
            "deliveredAt": item["delivered_at"].isoformat() if item.get("delivered_at") else "-",
            "review_required": item.get("review_required", False)
        }
        for item in items
    ]

# =====================================================
# ROUTES
# =====================================================
@app.get("/")
def home():
    return {"message": "Invoice Management API is running", "status": "healthy"}

@app.get("/health")
def health():
    try:
        client.admin.command("ping")
        return {"api": "ok", "mongodb": "connected"}
    except Exception:
        raise HTTPException(status_code=503, detail="MongoDB connection failed")

@app.post("/upload_pdf")
async def upload_pdf(invoices: List[UploadFile] = File(...)):
    if not invoices:
        raise HTTPException(status_code=400, detail="No files uploaded")

    all_processed = []
    total_delivery_items = 0

    for invoice in invoices:
        if not invoice.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"Only PDF files allowed. {invoice.filename} is invalid.")

        pdf_data = await invoice.read(MAX_FILE_SIZE + 1)
        if not pdf_data or len(pdf_data) > MAX_FILE_SIZE:
            raise HTTPException(status_code=400, detail=f"File {invoice.filename} is empty or too large.")
        if not pdf_data.startswith(b"%PDF-"):
            raise HTTPException(status_code=400, detail=f"Invalid PDF file: {invoice.filename}")

        invoice_id = str(uuid.uuid4())
        
        try:
            pdf_document = fitz.open(stream=pdf_data, filetype="pdf")
            
            if pdf_document.needs_pass:
                pdf_document.close()
                raise HTTPException(status_code=400, detail=f"Password-protected PDF: {invoice.filename}")

            total_pages = len(pdf_document)
            if total_pages == 0 or total_pages > MAX_PAGES:
                pdf_document.close()
                raise HTTPException(status_code=400, detail=f"Invalid page count for {invoice.filename}")

            full_text = "\n".join(page.get_text("text", sort=True) for page in pdf_document)
            pdf_document.close()

            if not full_text.strip():
                raise HTTPException(status_code=422, detail=f"No text found in {invoice.filename}")

            sections = extract_invoices_from_pdf(full_text, invoice.filename)
            
            upload_time = now_utc()
            invoice_documents = []
            delivery_documents = []

            for section_index, section in enumerate(sections):
                products = extract_products(section["text"])
                invoice_review = len(products) == 0 or any(p["review_required"] for p in products)

                invoice_document = {
                    "invoice_id": invoice_id,
                    "invoice_number": section["invoice_number"],
                    "invoice_sequence": section_index + 1,
                    "filename": invoice.filename,
                    "content_type": invoice.content_type,
                    "file_size": len(pdf_data),
                    "total_pages": total_pages,
                    "extracted_text": section["text"],
                    "products": products,
                    "product_count": len(products),
                    "review_required": invoice_review,
                    "extraction_status": "manual_review" if invoice_review else "completed",
                    "created_at": upload_time
                }
                invoice_documents.append(invoice_document)

                for product in products:
                    delivery_documents.append({
                        "invoice": section["invoice_number"],
                        "model": product["model"],
                        "size": product["size"],
                        "barcode": generate_barcode(),
                        "status": "Pending",
                        "delivered_at": None,
                        "created_at": upload_time,
                        "invoice_id": invoice_id,
                        "review_required": product["review_required"],
                        "confidence": product["confidence"]
                    })

            if invoice_documents:
                invoices_collection.insert_many(invoice_documents)
            if delivery_documents:
                delivery_collection.insert_many(delivery_documents)
                total_delivery_items += len(delivery_documents)

            all_processed.append({
                "filename": invoice.filename,
                "invoices_created": len(invoice_documents),
                "products_created": len(delivery_documents)
            })

        except HTTPException:
            raise
        except Exception as e:
            print(f"Error processing {invoice.filename}:", str(e))
            raise HTTPException(status_code=500, detail=f"Failed to process {invoice.filename}")

    return {
        "message": "PDFs processed successfully",
        "total_files_processed": len(all_processed),
        "total_delivery_items_created": total_delivery_items,
        "details": all_processed
    }

@app.get("/api/invoices")
def get_invoices():
    try:
        invoices = invoices_collection.find().sort("created_at", -1)
        return [serialize_invoice(inv) for inv in invoices]
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Failed to fetch invoices")

@app.get("/api/delivery-items")
def get_delivery_items():
    try:
        items = delivery_collection.find().sort("created_at", -1)
        return [serialize_delivery_item(item) for item in items]
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Failed to fetch delivery items")

@app.post("/api/delivery-items")
def create_delivery_item_api(item: DeliveryItemCreate):
    barcode = item.barcode.strip()
    if not barcode:
        raise HTTPException(status_code=400, detail="Barcode is required")
    if delivery_collection.find_one({"barcode": barcode}):
        raise HTTPException(status_code=409, detail="Barcode already exists")

    new_item = {
        "invoice": item.invoice.strip(),
        "model": item.model.strip(),
        "size": item.size.strip(),
        "barcode": barcode,
        "status": "Pending",
        "delivered_at": None,
        "created_at": now_utc(),
        "review_required": False,
        "confidence": 1.0
    }
    result = delivery_collection.insert_one(new_item)
    return {"message": "Delivery item created", "id": str(result.inserted_id)}

@app.patch("/api/delivery-items/{item_id}")
def update_delivery_item(item_id: str, data: DeliveryItemUpdate):
    from bson import ObjectId
    from bson.errors import InvalidId
    try:
        object_id = ObjectId(item_id)
    except InvalidId:
        raise HTTPException(status_code=400, detail="Invalid ID")

    updates = {}
    for field in ["invoice", "model", "size"]:
        value = getattr(data, field)
        if value is not None:
            value = value.strip()
            if not value:
                raise HTTPException(status_code=400, detail=f"{field} cannot be empty")
            updates[field] = value

    if not updates:
        raise HTTPException(status_code=400, detail="No fields supplied")

    updates["review_required"] = False
    updates["reviewed_at"] = now_utc()

    result = delivery_collection.update_one({"_id": object_id}, {"$set": updates})
    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="Item not found")
    return {"message": "Item updated"}

@app.post("/api/delivery-items/scan")
def deliver_item(data: BarcodeScan):
    barcode = data.barcode.strip()
    if not barcode:
        raise HTTPException(status_code=400, detail="Barcode required")

    item = delivery_collection.find_one({"barcode": barcode})
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    if item.get("status") == "Delivered":
        raise HTTPException(status_code=409, detail="Already delivered")

    delivered_at = now_utc()
    delivery_collection.update_one(
        {"_id": item["_id"]},
        {"$set": {"status": "Delivered", "delivered_at": delivered_at}}
    )
    return {"message": "Item delivered successfully", "barcode": barcode, "status": "Delivered", "date": delivered_at.isoformat()}

@app.get("/api/reports")
def get_reports(status: str = "all"):
    try:
        items = get_report_items(status)
        return {"status": status, "total": len(items), "items": items}
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Failed to fetch reports")

@app.get("/api/invoices/export")
def export_reports(status: str = "all"):
    try:
        items = get_report_items(status)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Delivery Report"

        worksheet.append(["Invoice", "Model", "Size", "Barcode", "Status", "Delivered At", "Manual Review"])
        for item in items:
            worksheet.append([
                item["invoice"], item["model"], item["size"], item["barcode"],
                item["status"], item["deliveredAt"], "Yes" if item["review_required"] else "No"
            ])

        for col, width in {"A": 22, "B": 25, "C": 15, "D": 25, "E": 18, "F": 30, "G": 25}.items():
            worksheet.column_dimensions[col].width = width

        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)

        return StreamingResponse(
            excel_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{status}-report.xlsx"'}
        )
    except Exception:
        raise HTTPException(status_code=500, detail="Excel export failed")

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))