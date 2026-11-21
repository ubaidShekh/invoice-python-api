from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, Response
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

client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=10000, connectTimeoutMS=10000)

db = client["invoice_delivery"]
invoices_collection = db["invoices"]
delivery_collection = db["delivery_items"]
pdf_files_collection = db["pdf_files"]
pdf_viewer_delivery_collection = db["pdf_viewer_delivery_items"]
customers_collection = db["customers"]

# Indexes
invoices_collection.create_index("invoice_id")
delivery_collection.create_index("barcode", unique=True)
delivery_collection.create_index("invoice")
pdf_files_collection.create_index("filename", unique=True)
pdf_viewer_delivery_collection.create_index("barcode", unique=True)
pdf_viewer_delivery_collection.create_index([("filename", 1), ("invoice", 1)])
pdf_files_collection.create_index("customer_id")
pdf_viewer_delivery_collection.create_index("customer_id") # NEW Index

MAX_FILE_SIZE = 25 * 1024 * 1024
MONGODB_BINARY_LIMIT = 15 * 1024 * 1024
MAX_PAGES = 150

# =====================================================
# MODELS
# =====================================================
class CustomerCreate(BaseModel):
    name: str
    company: Optional[str] = ""
    phone: Optional[str] = ""
    email: Optional[str] = ""

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

class InvoiceSelection(BaseModel):
    invoice_number: str
    model: str
    size: str

class PdfFooterRequest(BaseModel):
    selections: List[InvoiceSelection]

class PdfBulkDeleteRequest(BaseModel):
    filenames: List[str]

class PdfPrintAllSelection(BaseModel):
    filename: str
    invoice_number: str
    model: str
    size: str

class PdfPrintAllRequest(BaseModel):
    selections: List[PdfPrintAllSelection]

class DeliveryReportUpdate(BaseModel):
    normal_rate: Optional[str] = None
    shak: Optional[str] = None
    tatriz: Optional[str] = None
    amount: Optional[str] = None

# =====================================================
# REGEX PATTERNS
# =====================================================
ORDER_NUMBER_RE = re.compile(r"#(\d{7,12})")
MODEL_RE = re.compile(r"#(\d{2,4})\b")
PRODUCT_CODE_RE = re.compile(r"\b(\d{8,10})\b")
PHONE_RE = re.compile(r"(\+966\d{9,10})")
AMOUNT_RE = re.compile(r"SAR\s*([\d,]+)")
PAYMENT_METHOD_RE = re.compile(r"(مدى|تاي|تالي|تاني|البطاقة الائتمانية|البطاقة الالتمانية|أس تي سي باي|تابي|متمم)")

# =====================================================
# HELPERS
# =====================================================
def now_utc(): return datetime.now(timezone.utc)
def generate_barcode(): return "PV-" + uuid.uuid4().hex[:12].upper()
def clean_text(value): return re.sub(r"\s+", " ", str(value or "")).strip()

def extract_invoice_sections(full_text):
    matches = list(ORDER_NUMBER_RE.finditer(full_text))
    if not matches:
        return [{"invoice_number": "Unknown", "text": full_text}]
    sections = []
    seen_numbers = set()
    for index, match in enumerate(matches):
        invoice_number = match.group(1)
        if invoice_number in seen_numbers: continue
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
            sections.append({"invoice_number": invoice_number, "text": section_text})
    return sections

def extract_products(section_text):
    products = []
    seen = set()
    model_matches = list(MODEL_RE.finditer(section_text))
    for idx, match in enumerate(model_matches):
        model = "#" + match.group(1)
        start_pos = match.start()
        end_pos = model_matches[idx + 1].start() if idx + 1 < len(model_matches) else len(section_text)
        chunk = section_text[start_pos:end_pos]
        product_code = ""
        for code_match in PRODUCT_CODE_RE.finditer(chunk):
            if code_match.group(1) != match.group(1):
                product_code = code_match.group(1)
                break
        size = "N/A"
        all_numbers = re.findall(r"\b(\d{2,3})\b", chunk)
        for num in all_numbers:
            try:
                num_int = int(num)
                if 30 <= num_int <= 70:
                    if num != match.group(1) and num != product_code and not num.startswith("966"):
                        size = num
                        break
            except ValueError: continue
        key = (model, size, product_code)
        if key not in seen:
            seen.add(key)
            products.append({"model": model, "size": size, "product_code": product_code,
                             "confidence": 0.95, "review_required": False})
    return products

def extract_customer_info(section_text):
    customer_name = "Unknown"
    address = "Unknown"
    phone = "Unknown"
    total_amount = "0"
    payment_method = "Unknown"
    phone_match = PHONE_RE.search(section_text)
    if phone_match: phone = phone_match.group(1)
    amount_match = AMOUNT_RE.search(section_text)
    if amount_match: total_amount = amount_match.group(1).replace(",", "")
    payment_match = PAYMENT_METHOD_RE.search(section_text)
    if payment_match: payment_method = clean_text(payment_match.group(1))
    lines = section_text.split('\n')
    arabic_label_variants = ["مصدرة إلى", "مصدرة الى", "مصدرة إىل", "مصدرة الي", "إلى", "الى"]
    for i, line in enumerate(lines):
        cleaned = clean_text(line)
        for label in arabic_label_variants:
            if label in cleaned:
                after_label = cleaned.split(label, 1)[-1].strip()
                if after_label and len(after_label) > 2:
                    customer_name = after_label
                    break
                for j in range(i + 1, min(i + 4, len(lines))):
                    next_clean = clean_text(lines[j])
                    if next_clean and not any(city in next_clean for city in ["السعودية", "الرياض", "جدة", "الظهران", "المجمعة"]) and len(next_clean) > 2:
                        customer_name = next_clean
                        break
                if customer_name != "Unknown": break
        if customer_name != "Unknown": break
    if customer_name == "Unknown" and phone != "Unknown":
        for i, line in enumerate(lines):
            if phone in line:
                for j in range(max(0, i - 4), i):
                    prev_clean = clean_text(lines[j])
                    if re.match(r"^[A-Z][a-zA-Z\s]{2,}$", prev_clean):
                        customer_name = prev_clean
                        break
                break
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
    try:
        for i, line in enumerate(lines):
            if any(city in line for city in ["السعودية", "الرياض", "جدة", "الظهران", "المجمعة"]):
                if i + 1 < len(lines):
                    address = clean_text(line + " " + lines[i + 1])
                else:
                    address = clean_text(line)
                break
    except Exception: pass
    return {"customer_name": customer_name, "address": address, "phone": phone,
            "total_amount": total_amount, "payment_method": payment_method, "status": "Paid"}

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
    if status == "delivered": query["status"] = "Delivered"
    elif status == "undelivered": query["status"] = "Pending"
    items = delivery_collection.find(query).sort("delivered_at", -1)
    return [{
        "id": str(item["_id"]), "invoice": item.get("invoice", ""),
        "model": item.get("model", ""), "size": item.get("size", "N/A"),
        "barcode": item.get("barcode", ""), "status": item.get("status", "Pending"),
        "deliveredAt": item["delivered_at"].isoformat() if item.get("delivered_at") else "-",
        "normal_rate": item.get("normal_rate", ""), "shak": item.get("shak", ""),
        "tatriz": item.get("tatriz", ""), "amount": item.get("amount", "")
    } for item in items]

def save_selected_delivery_item(filename, invoice_number, model, size, customer_id):
    existing = pdf_viewer_delivery_collection.find_one({
        "filename": filename,
        "invoice": invoice_number
    })
    if existing:
        if existing.get("model") == model and existing.get("size") == size:
            return
        pdf_viewer_delivery_collection.update_one(
            {"_id": existing["_id"]},
            {"$set": {
                "model": model,
                "size": size,
                "customer_id": customer_id, # NEW
                "updated_at": now_utc()
            }}
        )
    else:
        pdf_viewer_delivery_collection.insert_one({
            "invoice": invoice_number,
            "model": model,
            "size": size,
            "barcode": generate_barcode(),
            "status": "Pending",
            "delivered_at": None,
            "created_at": now_utc(),
            "filename": filename,
            "customer_id": customer_id, # NEW
        })

# =====================================================
# ROUTES
# =====================================================
@app.get("/")
def home(): return {"message": "Invoice Management API is running", "status": "healthy"}

@app.get("/health")
def health():
    try:
        client.admin.command("ping")
        return {"api": "ok", "mongodb": "connected"}
    except Exception:
        raise HTTPException(status_code=503, detail="MongoDB connection failed")

# ----- CUSTOMERS -----
@app.post("/api/customers")
def create_customer(customer: CustomerCreate):
    try:
        result = customers_collection.insert_one({
            "name": customer.name,
            "company": customer.company,
            "phone": customer.phone,
            "email": customer.email,
            "created_at": now_utc()
        })
        return {"message": "Customer created", "id": str(result.inserted_id)}
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Failed to create customer")

@app.get("/api/customers")
def get_customers():
    try:
        customers = list(customers_collection.find().sort("created_at", -1))
        return [{
            "_id": str(c["_id"]),
            "name": c.get("name", ""),
            "company": c.get("company", ""),
            "phone": c.get("phone", ""),
            "email": c.get("email", ""),
            "created_at": c["created_at"].isoformat() if c.get("created_at") else ""
        } for c in customers]
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Failed to fetch customers")

# ----- PDF VIEWER (CUSTOMER SPECIFIC) -----
@app.post("/upload_pdf_viewer")
async def upload_pdf_viewer(
    invoices: List[UploadFile] = File(...),
    customer_id: str = Form(...)
):
    if not invoices: raise HTTPException(status_code=400, detail="No files uploaded")
    all_processed = []
    for invoice in invoices:
        if not invoice.filename.lower().endswith(".pdf"): continue
        pdf_data = await invoice.read(MAX_FILE_SIZE + 1)
        if not pdf_data or len(pdf_data) > MAX_FILE_SIZE: continue
        if not pdf_data.startswith(b"%PDF-"): continue
        invoice_id = str(uuid.uuid4())
        try:
            pdf_document = fitz.open(stream=pdf_data, filetype="pdf")
            full_text = "\n".join(page.get_text("text", sort=True) for page in pdf_document)
            total_pages = len(pdf_document)
            pdf_document.close()
            if not full_text.strip(): continue
            sections = extract_invoice_sections(full_text)
            upload_time = now_utc()
            invoice_documents = []
            for section_index, section in enumerate(sections):
                products = extract_products(section["text"])
                invoice_document = {
                    "invoice_id": invoice_id, "invoice_number": section["invoice_number"],
                    "filename": invoice.filename, "file_size": len(pdf_data),
                    "total_pages": total_pages, "extracted_text": section["text"],
                    "products": products, "product_count": len(products),
                    "created_at": upload_time, "source": "pdf_viewer",
                    "customer_id": customer_id
                }
                invoice_documents.append(invoice_document)
            if invoice_documents: invoices_collection.insert_many(invoice_documents)
            if len(pdf_data) <= MONGODB_BINARY_LIMIT:
                pdf_files_collection.update_one(
                    {"filename": invoice.filename},
                    {"$set": {"filename": invoice.filename, "content_type": invoice.content_type,
                              "file_data": pdf_data, "file_size": len(pdf_data),
                              "uploaded_at": upload_time, "customer_id": customer_id}},
                    upsert=True
                )
            all_processed.append({"filename": invoice.filename, "invoices_created": len(invoice_documents)})
        except Exception as e:
            print(f"Error processing {invoice.filename}:", str(e))
            continue
    return {"message": "Files processed for viewer", "total_files_processed": len(all_processed),
            "details": all_processed}

@app.get("/api/pdf-files")
def get_pdf_files(customer_id: Optional[str] = None):
    try:
        query = {}
        if customer_id: query["customer_id"] = customer_id
        files = pdf_files_collection.find(query).sort("uploaded_at", -1)
        result = []
        for f in files:
            invoice_count = invoices_collection.count_documents({"filename": f["filename"], "source": "pdf_viewer"})
            delivery_count = pdf_viewer_delivery_collection.count_documents({"filename": f["filename"]})
            result.append({
                "filename": f["filename"], "file_size": f.get("file_size", 0),
                "uploaded_at": f["uploaded_at"].isoformat() if f.get("uploaded_at") else "",
                "total_invoices": invoice_count, "total_delivery_items": delivery_count
            })
        return result
    except PyMongoError: raise HTTPException(status_code=500, detail="Failed to fetch PDF files")

@app.get("/api/pdf-invoice-details/{filename}")
def get_pdf_invoice_details(filename: str):
    try:
        invoices = invoices_collection.find({"filename": filename, "source": "pdf_viewer"}).sort("created_at", 1)
        result = []
        for inv in invoices:
            products = inv.get("products", [])
            selected_product = products[0] if products else {"model": "N/A", "size": "N/A"}
            result.append({"_id": str(inv["_id"]), "invoice_number": inv.get("invoice_number", ""),
                           "products": products, "selected_product": selected_product, "name": ""})
        return {"filename": filename, "total_invoices": len(result), "invoices": result}
    except PyMongoError: raise HTTPException(status_code=500, detail="Failed to fetch PDF details")

@app.get("/api/pdf-viewer-delivery-items")
def get_pdf_viewer_delivery_items():
    try:
        items = list(pdf_viewer_delivery_collection.find().sort("created_at", -1))
        # Fetch all customers to map ID -> Name
        customers = {str(c["_id"]): c.get("name", "Unknown") for c in customers_collection.find()}
        
        result = []
        for item in items:
            cust_id = item.get("customer_id")
            cust_name = customers.get(cust_id, "Unknown Customer")
            
            result.append({
                "_id": str(item["_id"]), "invoice": item.get("invoice", ""),
                "model": item.get("model", ""), "size": item.get("size", "N/A"),
                "barcode": item.get("barcode", ""), "status": item.get("status", "Pending"),
                "filename": item.get("filename", ""),
                "customer_name": cust_name, # NEW
                "date": item["delivered_at"].isoformat() if item.get("delivered_at") else "-"
            })
        return result
    except PyMongoError: raise HTTPException(status_code=500, detail="Failed to fetch PDF viewer delivery items")

@app.post("/api/pdf-viewer-delivery-items/scan")
def deliver_pdf_viewer_item(data: BarcodeScan):
    barcode = data.barcode.strip()
    if not barcode: raise HTTPException(status_code=400, detail="Barcode is required")
    item = pdf_viewer_delivery_collection.find_one({"barcode": barcode})
    if not item: raise HTTPException(status_code=404, detail="Item not found in PDF Viewer data")
    if item.get("status") == "Delivered": raise HTTPException(status_code=409, detail="Item already delivered")
    delivered_at = now_utc()
    pdf_viewer_delivery_collection.update_one({"_id": item["_id"]}, {"$set": {"status": "Delivered", "delivered_at": delivered_at}})
    return {"message": "Item delivered successfully", "barcode": barcode, "status": "Delivered", "date": delivered_at.isoformat()}

@app.post("/api/pdf-with-footer/{filename}")
def get_pdf_with_footer(filename: str, request: PdfFooterRequest):
    try:
        file_doc = pdf_files_collection.find_one({"filename": filename})
        if not file_doc: raise HTTPException(status_code=404, detail="PDF not found")
        pdf_data = file_doc.get("file_data")
        if not pdf_data: raise HTTPException(status_code=500, detail="PDF binary empty")
        
        customer_id = file_doc.get("customer_id") # NEW
        
        invoice_map = {}
        for sel in request.selections:
            invoice_map[sel.invoice_number] = {"invoice": sel.invoice_number, "model": sel.model, "size": sel.size}
        pdf_document = fitz.open(stream=pdf_data, filetype="pdf")
        page_invoices = []
        for page_num, page in enumerate(pdf_document):
            page_text = page.get_text("text", sort=True)
            order_match = re.search(r"#(\d{7,12})", page_text)
            page_invoices.append(order_match.group(1) if order_match else None)
        invoice_last_page = {}
        for page_num, inv_num in enumerate(page_invoices):
            if inv_num: invoice_last_page[inv_num] = page_num
        for page_num, page in enumerate(pdf_document):
            current_inv = page_invoices[page_num]
            if not current_inv: continue
            if invoice_last_page.get(current_inv) != page_num: continue
            data = invoice_map.get(current_inv)
            if not data: continue
            page_rect = page.rect
            footer_height = 36
            footer_y = page_rect.height - footer_height - 5
            footer_rect = fitz.Rect(15, footer_y, page_rect.width - 15, page_rect.height - 5)
            page.draw_rect(footer_rect, color=(0.85, 0.85, 0.85), fill=(0.97, 0.97, 0.97), width=0.5)
            full_text = f"Inv: #{data['invoice']}    |    Model: {data['model']}    |    Size: {data['size']}    |    Name: ________________"
            font_size = 12
            fontname = "helv"
            max_width = page_rect.width * 0.80
            while font_size > 8:
                text_width = fitz.get_text_length(full_text, fontname=fontname, fontsize=font_size)
                if text_width <= max_width: break
                font_size -= 0.5
            text_y = footer_y + footer_height / 2 + font_size / 3
            page.insert_text((20, text_y), full_text, fontsize=font_size, color=(0.1, 0.1, 0.1), fontname=fontname)
        output = BytesIO()
        pdf_document.save(output, garbage=3, deflate=True)
        pdf_document.close()
        output.seek(0)
        try:
            for sel in request.selections:
                save_selected_delivery_item(filename, sel.invoice_number, sel.model, sel.size, customer_id) # NEW
        except PyMongoError as e:
            print("Failed to save delivery items:", str(e))
        return Response(content=output.getvalue(), media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="modified_{filename}"',
                     "Content-Length": str(len(output.getvalue())),
                     "Cache-Control": "no-cache, no-store, must-revalidate"})
    except HTTPException: raise
    except Exception as error:
        print("PDF Footer Generation Error:", repr(error))
        raise HTTPException(status_code=500, detail=f"Failed to generate PDF footer: {str(error)}")

@app.post("/api/pdf-print-all-with-selections")
def print_all_pdfs_with_selections(request: PdfPrintAllRequest):
    try:
        if not request.selections: raise HTTPException(status_code=400, detail="No selections provided")
        selections_by_file = {}
        for sel in request.selections:
            if sel.filename not in selections_by_file:
                selections_by_file[sel.filename] = {}
            selections_by_file[sel.filename][sel.invoice_number] = {"model": sel.model, "size": sel.size}
        output_doc = fitz.open()
        processed_files = 0
        for filename, invoice_map in selections_by_file.items():
            file_doc = pdf_files_collection.find_one({"filename": filename})
            if not file_doc: continue
            pdf_data = file_doc.get("file_data")
            if not pdf_data: continue
            
            customer_id = file_doc.get("customer_id") # NEW
            
            pdf_document = fitz.open(stream=pdf_data, filetype="pdf")
            page_invoices = []
            for page_num, page in enumerate(pdf_document):
                page_text = page.get_text("text", sort=True)
                order_match = re.search(r"#(\d{7,12})", page_text)
                page_invoices.append(order_match.group(1) if order_match else None)
            invoice_last_page = {}
            for page_num, inv_num in enumerate(page_invoices):
                if inv_num: invoice_last_page[inv_num] = page_num
            for page_num, page in enumerate(pdf_document):
                current_inv = page_invoices[page_num]
                if not current_inv: continue
                if invoice_last_page.get(current_inv) != page_num: continue
                data = invoice_map.get(current_inv)
                if not data: continue
                page_rect = page.rect
                footer_height = 36
                footer_y = page_rect.height - footer_height - 5
                footer_rect = fitz.Rect(15, footer_y, page_rect.width - 15, page_rect.height - 5)
                page.draw_rect(footer_rect, color=(0.85, 0.85, 0.85), fill=(0.97, 0.97, 0.97), width=0.5)
                full_text = f"Inv: #{current_inv}    |    Model: {data['model']}    |    Size: {data['size']}    |    Name: ________________"
                font_size = 12
                fontname = "helv"
                max_width = page_rect.width * 0.80
                while font_size > 8:
                    text_width = fitz.get_text_length(full_text, fontname=fontname, fontsize=font_size)
                    if text_width <= max_width: break
                    font_size -= 0.5
                text_y = footer_y + footer_height / 2 + font_size / 3
                page.insert_text((20, text_y), full_text, fontsize=font_size, color=(0.1, 0.1, 0.1), fontname=fontname)
            output_doc.insert_pdf(pdf_document)
            pdf_document.close()
            processed_files += 1
            
            # Save delivery items for this file
            try:
                for inv_num, data in invoice_map.items():
                    save_selected_delivery_item(filename, inv_num, data["model"], data["size"], customer_id) # NEW
            except PyMongoError as e:
                print("Failed to save delivery items:", str(e))
                
        if processed_files == 0: raise HTTPException(status_code=404, detail="No PDFs could be processed")
        output = BytesIO()
        output_doc.save(output, garbage=3, deflate=True)
        output_doc.close()
        output.seek(0)
        return Response(content=output.getvalue(), media_type="application/pdf",
            headers={"Content-Disposition": 'inline; filename="all_invoices_with_footer.pdf"',
                     "Content-Length": str(len(output.getvalue())),
                     "Cache-Control": "no-cache, no-store, must-revalidate"})
    except HTTPException: raise
    except Exception as error:
        print("Print All with Selections Error:", repr(error))
        raise HTTPException(status_code=500, detail=f"Failed to generate combined PDF: {str(error)}")

@app.delete("/api/pdf-files/{filename}")
def delete_pdf_file(filename: str):
    try:
        pdf_result = pdf_files_collection.delete_one({"filename": filename})
        invoice_result = invoices_collection.delete_many({"filename": filename, "source": "pdf_viewer"})
        delivery_result = pdf_viewer_delivery_collection.delete_many({"filename": filename})
        if pdf_result.deleted_count == 0 and invoice_result.deleted_count == 0:
            raise HTTPException(status_code=404, detail="PDF not found")
        return {"message": "PDF deleted successfully", "filename": filename,
                "deleted_binary": pdf_result.deleted_count, "deleted_invoices": invoice_result.deleted_count,
                "deleted_delivery_items": delivery_result.deleted_count}
    except HTTPException: raise
    except PyMongoError: raise HTTPException(status_code=500, detail="Failed to delete PDF")

@app.post("/api/pdf-files/delete-bulk")
def delete_pdf_files_bulk(request: PdfBulkDeleteRequest):
    try:
        deleted_count = 0
        deleted_invoice_count = 0
        deleted_delivery_count = 0
        not_found = []
        for filename in request.filenames:
            pdf_result = pdf_files_collection.delete_one({"filename": filename})
            invoice_result = invoices_collection.delete_many({"filename": filename, "source": "pdf_viewer"})
            delivery_result = pdf_viewer_delivery_collection.delete_many({"filename": filename})
            if pdf_result.deleted_count > 0 or invoice_result.deleted_count > 0:
                deleted_count += 1
                deleted_invoice_count += invoice_result.deleted_count
                deleted_delivery_count += delivery_result.deleted_count
            else:
                not_found.append(filename)
        return {"message": f"{deleted_count} PDF(s) deleted successfully", "deleted_count": deleted_count,
                "deleted_invoices": deleted_invoice_count, "deleted_delivery_items": deleted_delivery_count,
                "not_found": not_found}
    except PyMongoError: raise HTTPException(status_code=500, detail="Failed to delete PDFs")

@app.patch("/api/delivery-items/{item_id}/report")
def update_delivery_report(item_id: str, data: DeliveryReportUpdate):
    from bson import ObjectId
    from bson.errors import InvalidId
    try: object_id = ObjectId(item_id)
    except InvalidId: raise HTTPException(status_code=400, detail="Invalid item ID")
    updates = {}
    for field in ["normal_rate", "shak", "tatriz", "amount"]:
        value = getattr(data, field)
        if value is not None: updates[field] = value.strip()
    if not updates: raise HTTPException(status_code=400, detail="No fields to update")
    result = delivery_collection.update_one({"_id": object_id}, {"$set": updates})
    if result.matched_count == 0: raise HTTPException(status_code=404, detail="Item not found")
    return {"message": "Report updated"}

@app.get("/api/pdf-viewer-reports")
def get_pdf_viewer_reports(status: str = "all"):
    try:
        query = {}
        if status == "delivered": query["status"] = "Delivered"
        elif status == "undelivered": query["status"] = "Pending"
        items_list = list(pdf_viewer_delivery_collection.find(query))
        
        # Fetch all customers to map ID -> Name
        customers = {str(c["_id"]): c.get("name", "Unknown") for c in customers_collection.find()}
        
        def get_sort_key(item):
            dt = item.get("delivered_at")
            if dt is None: return datetime.min.replace(tzinfo=timezone.utc)
            return dt
        items_list.sort(key=get_sort_key, reverse=True)
        result_items = []
        for item in items_list:
            delivered_at = item.get("delivered_at")
            if delivered_at is None: delivered_at_str = "-"
            elif isinstance(delivered_at, str): delivered_at_str = delivered_at
            else: delivered_at_str = delivered_at.isoformat()
            
            cust_id = item.get("customer_id")
            cust_name = customers.get(cust_id, "Unknown Customer")
            
            result_items.append({
                "id": str(item["_id"]), "invoice": item.get("invoice", ""),
                "model": item.get("model", ""), "size": item.get("size", "N/A"),
                "barcode": item.get("barcode", ""), "status": item.get("status", "Pending"),
                "deliveredAt": delivered_at_str,
                "customer_name": cust_name, # NEW
                "normal_rate": item.get("normal_rate", ""), "shak": item.get("shak", ""),
                "tatriz": item.get("tatriz", ""), "amount": item.get("amount", "")
            })
        return {"status": status, "total": len(result_items), "items": result_items}
    except PyMongoError as error:
        print("PDF Viewer Reports Error:", str(error))
        raise HTTPException(status_code=500, detail="Failed to fetch PDF viewer reports")
    except Exception as error:
        print("Unexpected Error in Reports:", repr(error))
        raise HTTPException(status_code=500, detail=f"Error: {str(error)}")

@app.patch("/api/pdf-viewer-delivery-items/{item_id}/report")
def update_pdf_viewer_report(item_id: str, data: DeliveryReportUpdate):
    from bson import ObjectId
    from bson.errors import InvalidId
    try: object_id = ObjectId(item_id)
    except InvalidId: raise HTTPException(status_code=400, detail="Invalid item ID")
    updates = {}
    for field in ["normal_rate", "shak", "tatriz", "amount"]:
        value = getattr(data, field)
        if value is not None: updates[field] = value.strip()
    if not updates: raise HTTPException(status_code=400, detail="No fields to update")
    result = pdf_viewer_delivery_collection.update_one({"_id": object_id}, {"$set": updates})
    if result.matched_count == 0: raise HTTPException(status_code=404, detail="Item not found")
    return {"message": "Report updated"}

@app.post("/api/admin/migrate-clean-old-delivery-items")
def migrate_clean_old_delivery_items():
    try:
        result = pdf_viewer_delivery_collection.delete_many({})
        return {"message": "Old delivery items cleared", "deleted_count": result.deleted_count}
    except PyMongoError:
        raise HTTPException(status_code=500, detail="Migration failed")

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))