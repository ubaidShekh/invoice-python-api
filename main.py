from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from dotenv import load_dotenv
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
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

client = MongoClient(
    MONGODB_URI,
    serverSelectionTimeoutMS=10000,
    connectTimeoutMS=10000
)

db = client["invoice_delivery"]
invoices_collection = db["invoices"]
delivery_collection = db["delivery_items"]

# Create indexes
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
# REGEX PATTERNS
# =====================================================

ORDER_NUMBER_RE = re.compile(r"#(\d{7,12})")
MODEL_RE = re.compile(r"#(\d{2,4})\b")
PRODUCT_CODE_RE = re.compile(r"\b(\d{8,10})\b")

# Customer info patterns
PHONE_RE = re.compile(r"(\+966\d{9,10})")
AMOUNT_RE = re.compile(r"SAR\s*([\d,]+)")
PAYMENT_METHOD_RE = re.compile(r"(مدى|تاي|تالي|تاني|البطاقة الائتمانية|البطاقة الالتمانية|أس تي سي باي|تابي|متمم)")


# =====================================================
# HELPERS
# =====================================================

def now_utc():
    return datetime.now(timezone.utc)


def generate_barcode():
    return "INV-" + uuid.uuid4().hex[:12].upper()


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def extract_invoice_sections(full_text):
    """
    Detect separate invoice sections using order numbers (#123456789).
    If a page doesn't have a new order number, it belongs to the previous invoice.
    """
    matches = list(ORDER_NUMBER_RE.finditer(full_text))

    if not matches:
        return [{
            "invoice_number": "Unknown",
            "text": full_text
        }]

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
    
    # Saare models aur unki positions dhundo
    model_matches = list(MODEL_RE.finditer(section_text))
    
    for idx, match in enumerate(model_matches):
        model = "#" + match.group(1)
        start_pos = match.start()
        
        # Model ke aage ka text (agla model aane tak) ka chunk lo
        end_pos = model_matches[idx + 1].start() if idx + 1 < len(model_matches) else len(section_text)
        chunk = section_text[start_pos:end_pos]
        
        # Product Code dhundo (8-10 digit number, lekin Model ka number nahi)
        product_code = ""
        for code_match in PRODUCT_CODE_RE.finditer(chunk):
            if code_match.group(1) != match.group(1):
                product_code = code_match.group(1)
                break

        # SIZE EXTRACTION - Saare 2-digit numbers dhundo jo valid size (30-70) ho
        size = "N/A"
        
        # Chunk me se saare numbers nikalo
        all_numbers = re.findall(r"\b(\d{2,3})\b", chunk)
        
        for num in all_numbers:
            try:
                num_int = int(num)
                # Valid size range (30 to 70) aur Model/Product Code se alag ho
                if 30 <= num_int <= 70:
                    if num != match.group(1) and num != product_code and not num.startswith("966"):
                        size = num
                        break
            except ValueError:
                continue

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


def extract_customer_info(section_text):
    customer_name = "Unknown"
    address = "Unknown"
    phone = "Unknown"
    total_amount = "0"
    payment_method = "Unknown"

    # 1. PHONE
    phone_match = PHONE_RE.search(section_text)
    if phone_match:
        phone = phone_match.group(1)

    # 2. AMOUNT
    amount_match = AMOUNT_RE.search(section_text)
    if amount_match:
        total_amount = amount_match.group(1).replace(",", "")

    # 3. PAYMENT METHOD
    payment_match = PAYMENT_METHOD_RE.search(section_text)
    if payment_match:
        payment_method = clean_text(payment_match.group(1))

    # 4. CUSTOMER NAME (Multiple strategies)
    lines = section_text.split('\n')
    
    # Strategy A: Look for Arabic label and take the next non-empty line
    arabic_label_variants = [
        "مصدرة إلى", "مصدرة الى", "مصدرة إىل", 
        "مصدرة الي", "إلى", "الى"
    ]
    
    for i, line in enumerate(lines):
        cleaned = clean_text(line)
        for label in arabic_label_variants:
            if label in cleaned:
                # Check if name is on the same line after the label
                after_label = cleaned.split(label, 1)[-1].strip()
                if after_label and len(after_label) > 2:
                    customer_name = after_label
                    break
                # Otherwise check next few lines for a name
                for j in range(i + 1, min(i + 4, len(lines))):
                    next_clean = clean_text(lines[j])
                    if next_clean and not any(city in next_clean for city in ["السعودية", "الرياض", "جدة", "الظهران", "المجمعة"]) and len(next_clean) > 2:
                        customer_name = next_clean
                        break
                if customer_name != "Unknown":
                    break
        if customer_name != "Unknown":
            break

    # Strategy B: If name still unknown, look for English name near phone
    if customer_name == "Unknown" and phone != "Unknown":
        for i, line in enumerate(lines):
            if phone in line:
                # Check previous 4 lines for an English name
                for j in range(max(0, i - 4), i):
                    prev_clean = clean_text(lines[j])
                    if re.match(r"^[A-Z][a-zA-Z\s]{2,}$", prev_clean):
                        customer_name = prev_clean
                        break
                break

    # Strategy C: Last fallback - any English line that's not a city/keyword
    if customer_name == "Unknown":
        skip_keywords = ["PAID", "SAR", "سمسا", "بواسطة", "مدى", "تاي", "تالي", "تاني", 
                        "البطاقة", "تفاصيل", "المبلغ", "طريقة", "الدفع", "الكمية", 
                        "الرقم", "المنتج", "خيارات", "المقاس", "azuriteabaya", "الرقم الضريبي"]
        
        for line in lines:
            cleaned = clean_text(line)
            if (re.match(r"^[A-Z][a-zA-Z\s]{3,}$", cleaned) 
                and not any(kw.lower() in cleaned.lower() for kw in skip_keywords)
                and cleaned not in ["Unknown", ""]):
                customer_name = cleaned
                break

    # 5. ADDRESS
    try:
        for i, line in enumerate(lines):
            if any(city in line for city in ["السعودية", "الرياض", "جدة", "الظهران", "المجمعة"]):
                if i + 1 < len(lines):
                    address = clean_text(line + " " + lines[i + 1])
                else:
                    address = clean_text(line)
                break
    except Exception:
        pass

    return {
        "customer_name": customer_name,
        "address": address,
        "phone": phone,
        "total_amount": total_amount,
        "payment_method": payment_method,
        "status": "Paid"
    }


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
        "customer_info": document.get("customer_info", {}),
        "review_required": document.get("review_required", False),
        "date": document["created_at"].isoformat() if document.get("created_at") else ""
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
            "normal_rate": "",
            "shak": "",
            "tatriz": "",
            "amount": ""
        }
        for item in items
    ]


# =====================================================
# ROUTES
# =====================================================

@app.get("/")
def home():
    return {
        "message": "Invoice Management API is running",
        "status": "healthy"
    }


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
            raise HTTPException(
                status_code=400,
                detail=f"Only PDF files allowed. {invoice.filename} is invalid."
            )

        pdf_data = await invoice.read(MAX_FILE_SIZE + 1)

        if not pdf_data or len(pdf_data) > MAX_FILE_SIZE:
            raise HTTPException(
                status_code=400,
                detail=f"File {invoice.filename} is empty or too large."
            )

        if not pdf_data.startswith(b"%PDF-"):
            raise HTTPException(
                status_code=400,
                detail=f"Invalid PDF file: {invoice.filename}"
            )

        invoice_id = str(uuid.uuid4())

        try:
            pdf_document = fitz.open(stream=pdf_data, filetype="pdf")

            if pdf_document.needs_pass:
                pdf_document.close()
                raise HTTPException(
                    status_code=400,
                    detail=f"Password-protected PDF: {invoice.filename}"
                )

            total_pages = len(pdf_document)

            if total_pages == 0 or total_pages > MAX_PAGES:
                pdf_document.close()
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid page count for {invoice.filename}"
                )

            full_text = "\n".join(
                page.get_text("text", sort=True) for page in pdf_document
            )
            pdf_document.close()

            if not full_text.strip():
                raise HTTPException(
                    status_code=422,
                    detail=f"No text found in {invoice.filename}"
                )

            sections = extract_invoice_sections(full_text)

            upload_time = now_utc()
            invoice_documents = []
            delivery_documents = []

            for section_index, section in enumerate(sections):
                products = extract_products(section["text"])
                customer_info = extract_customer_info(section["text"])
                invoice_review = len(products) == 0

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
                    "customer_info": customer_info,
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
        except fitz.FileDataError:
            raise HTTPException(
                status_code=400,
                detail=f"The uploaded PDF is damaged or invalid: {invoice.filename}"
            )
        except PyMongoError as error:
            print("MongoDB Upload Error:", str(error))
            raise HTTPException(status_code=503, detail="Database operation failed")
        except Exception as error:
            print(f"Error processing {invoice.filename}:", repr(error))
            raise HTTPException(
                status_code=500,
                detail=f"Failed to process {invoice.filename}"
            )

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
    except PyMongoError as error:
        print("Invoice Fetch Error:", str(error))
        raise HTTPException(status_code=500, detail="Failed to fetch invoices")


@app.get("/api/delivery-items")
def get_delivery_items():
    try:
        items = delivery_collection.find().sort("created_at", -1)
        return [{
            "_id": str(item["_id"]),
            "invoice": item.get("invoice", ""),
            "model": item.get("model", ""),
            "size": item.get("size", "N/A"),
            "barcode": item.get("barcode", ""),
            "status": item.get("status", "Pending"),
            "date": item["delivered_at"].isoformat() if item.get("delivered_at") else "-",
            "review_required": item.get("review_required", False)
        } for item in items]
    except PyMongoError as error:
        print("Delivery Fetch Error:", str(error))
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
        raise HTTPException(status_code=400, detail="Invalid delivery item ID")

    updates = {}

    for field in ["invoice", "model", "size"]:
        value = getattr(data, field)

        if value is not None:
            value = value.strip()

            if not value:
                raise HTTPException(
                    status_code=400,
                    detail=f"{field} cannot be empty"
                )
            updates[field] = value

    if not updates:
        raise HTTPException(status_code=400, detail="No fields supplied for update")

    updates["review_required"] = False
    updates["reviewed_at"] = now_utc()

    result = delivery_collection.update_one(
        {"_id": object_id},
        {"$set": updates}
    )

    if result.matched_count == 0:
        raise HTTPException(status_code=404, detail="Delivery item not found")

    return {"message": "Delivery item updated successfully"}


@app.post("/api/delivery-items/scan")
def deliver_item(data: BarcodeScan):
    barcode = data.barcode.strip()

    if not barcode:
        raise HTTPException(status_code=400, detail="Barcode is required")

    item = delivery_collection.find_one({"barcode": barcode})

    if not item:
        raise HTTPException(status_code=404, detail="Item not found")

    if item.get("status") == "Delivered":
        raise HTTPException(status_code=409, detail="Item already delivered")

    delivered_at = now_utc()

    delivery_collection.update_one(
        {"_id": item["_id"]},
        {"$set": {"status": "Delivered", "delivered_at": delivered_at}}
    )

    return {
        "message": "Item delivered successfully",
        "barcode": barcode,
        "status": "Delivered",
        "date": delivered_at.isoformat()
    }


@app.get("/api/reports")
def get_reports(status: str = "all"):
    try:
        items = get_report_items(status)
        return {"status": status, "total": len(items), "items": items}
    except HTTPException:
        raise
    except PyMongoError as error:
        print("Reports Error:", str(error))
        raise HTTPException(status_code=500, detail="Failed to fetch reports")


@app.get("/api/invoices/export")
def export_reports(status: str = "all"):
    try:
        items = get_report_items(status)

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Delivery Report"

        headers = [
            "Invoice No.", "Model", "Size", "Barcode",
            "Normal Rate", "Shak", "Tatriz", "Amount", "Delivered"
        ]
        worksheet.append(headers)

        green_fill = PatternFill(start_color="28A745", end_color="28A745", fill_type="solid")
        white_font = Font(bold=True, color="FFFFFF")

        for cell in worksheet[1]:
            cell.fill = green_fill
            cell.font = white_font

        for item in items:
            worksheet.append([
                item["invoice"],
                item["model"],
                item["size"],
                item["barcode"],
                item["normal_rate"],
                item["shak"],
                item["tatriz"],
                item["amount"],
                item["status"]
            ])

        widths = {"A": 15, "B": 10, "C": 10, "D": 35, "E": 15, "F": 15, "G": 15, "H": 15, "I": 15}
        for col, width in widths.items():
            worksheet.column_dimensions[col].width = width

        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)

        return StreamingResponse(
            excel_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{status}-report.xlsx"'}
        )

    except HTTPException:
        raise
    except Exception as error:
        print("Excel Export Error:", repr(error))
        raise HTTPException(status_code=500, detail="Excel export failed")


# =====================================================
# SERVER
# =====================================================

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 3000))
    )