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
pdf_viewer_delivery_collection.create_index("customer_id")

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
    """Extract all models + sizes from a section."""
    products = []
    models = list(MODEL_RE.finditer(section_text))
    for m in models:
        model = m.group(1)
        # look ahead near model for a size like "40" or "42"
        window = section_text[m.end(): m.end() + 120]
        size_match = re.search(r"\b(\d{2})\b", window)
        size = size_match.group(1) if size_match else "-"
        products.append({
            "model": model,
            "size": size,
            "product_code": (PRODUCT_CODE_RE.search(window).group(1) if PRODUCT_CODE_RE.search(window) else ""),
        })
    if not products:
        products = [{"model": "N/A", "size": "N/A", "product_code": ""}]
    return products

def resolve_customer_names(items: list) -> list:
    """
    Attach customer_name to every delivery item by looking up customer_id
    in customers_collection. Items without customer_id fall back to
    pdf_files_collection lookup via filename, else 'Unknown'.
    """
    if not items:
        return items

    # Collect all customer_ids we need
    customer_ids = set()
    for it in items:
        cid = it.get("customer_id")
        if cid:
            customer_ids.add(str(cid))

    # Fetch customers in one shot
    customer_map = {}
    if customer_ids:
        for cust in customers_collection.find({"_id": {"$in": list(customer_ids)}}, {"name": 1}):
            customer_map[str(cust["_id"])] = cust.get("name") or "Unknown"

    # For items missing customer_id, try filename -> customer_id via pdf_files
    filenames = set()
    for it in items:
        if not it.get("customer_id") and it.get("filename"):
            filenames.add(it["filename"])

    filename_customer_map = {}
    if filenames:
        for pf in pdf_files_collection.find(
            {"filename": {"$in": list(filenames)}},
            {"filename": 1, "customer_id": 1},
        ):
            cid = pf.get("customer_id")
            if cid:
                filename_customer_map[pf["filename"]] = str(cid)
                if str(cid) not in customer_map:
                    cust = customers_collection.find_one({"_id": cid}, {"name": 1})
                    if cust:
                        customer_map[str(cid)] = cust.get("name") or "Unknown"

    # Attach name
    for it in items:
        cid = it.get("customer_id")
        name = None
        if cid and str(cid) in customer_map:
            name = customer_map[str(cid)]
        elif it.get("filename") and it["filename"] in filename_customer_map:
            name = customer_map.get(filename_customer_map[it["filename"]])
        it["customer_name"] = name or "Unknown"
    return items

# =====================================================
# CUSTOMERS
# =====================================================
@app.get("/api/customers")
def list_customers():
    out = []
    for c in customers_collection.find().sort("name", 1):
        out.append({
            "_id": str(c["_id"]),
            "name": c.get("name", ""),
            "company": c.get("company", ""),
            "phone": c.get("phone", ""),
            "email": c.get("email", ""),
            "created_at": c.get("created_at").isoformat() if c.get("created_at") else None,
        })
    return out

@app.post("/api/customers")
def create_customer(payload: CustomerCreate):
    if not payload.name.strip():
        raise HTTPException(status_code=400, detail="Customer name is required.")
    existing = customers_collection.find_one({"name": payload.name.strip()})
    if existing:
        raise HTTPException(status_code=409, detail="Customer with this name already exists.")
    doc = {
        "_id": str(uuid.uuid4()),
        "name": payload.name.strip(),
        "company": payload.company or "",
        "phone": payload.phone or "",
        "email": payload.email or "",
        "created_at": now_utc(),
    }
    customers_collection.insert_one(doc)
    return {"_id": doc["_id"], "name": doc["name"]}

# =====================================================
# PDF FILES (per customer)
# =====================================================
@app.get("/api/pdf-files")
def list_pdf_files(customer_id: str):
    if not customer_id:
        raise HTTPException(status_code=400, detail="customer_id is required")
    out = []
    for f in pdf_files_collection.find({"customer_id": customer_id}).sort("uploaded_at", -1):
        out.append({
            "filename": f.get("filename"),
            "file_size": f.get("file_size", 0),
            "total_invoices": f.get("total_invoices", 0),
            "total_delivery_items": f.get("total_delivery_items", 0),
            "uploaded_at": f.get("uploaded_at").isoformat() if f.get("uploaded_at") else None,
        })
    return out

@app.post("/api/pdf-files/delete-bulk")
def delete_pdf_files(payload: PdfBulkDeleteRequest):
    if not payload.filenames:
        raise HTTPException(status_code=400, detail="No filenames provided")
    deleted_files = 0
    deleted_items = 0
    for fn in payload.filenames:
        r1 = pdf_files_collection.delete_one({"filename": fn})
        r2 = pdf_viewer_delivery_collection.delete_many({"filename": fn})
        deleted_files += r1.deleted_count
        deleted_items += r2.deleted_count
    return {
        "message": f"Deleted {deleted_files} PDF file(s) and {deleted_items} delivery item(s).",
        "deleted_files": deleted_files,
        "deleted_items": deleted_items,
    }

# =====================================================
# UPLOAD PDF -> parse invoices + create delivery items
# =====================================================
@app.post("/upload_pdf_viewer")
async def upload_pdf_viewer(
    invoices: List[UploadFile] = File(...),
    customer_id: str = Form(...),
):
    if not customer_id:
        raise HTTPException(status_code=400, detail="customer_id is required")

    cust = customers_collection.find_one({"_id": customer_id}, {"name": 1})
    if not cust:
        raise HTTPException(status_code=404, detail="Customer not found")
    customer_name = cust.get("name") or "Unknown"

    total_files = 0
    total_invoices = 0
    total_items = 0

    for upload in invoices:
        if not upload.filename.lower().endswith(".pdf"):
            continue
        raw = await upload.read()
        if len(raw) > MAX_FILE_SIZE:
            raise HTTPException(status_code=400, detail=f"{upload.filename} exceeds size limit.")

        try:
            doc = fitz.open(stream=raw, filetype="pdf")
        except Exception:
            raise HTTPException(status_code=400, detail=f"Invalid PDF: {upload.filename}")
        if doc.page_count > MAX_PAGES:
            doc.close()
            raise HTTPException(status_code=400, detail=f"{upload.filename} has too many pages.")

        full_text = ""
        for page in doc:
            full_text += page.get_text() + "\n"
        doc.close()

        sections = extract_invoice_sections(full_text)
        existing = pdf_files_collection.find_one({"filename": upload.filename})
        if existing:
            pdf_files_collection.delete_one({"_id": existing["_id"]})
            pdf_viewer_delivery_collection.delete_many({"filename": upload.filename})

        pdf_files_collection.insert_one({
            "_id": str(uuid.uuid4()),
            "filename": upload.filename,
            "customer_id": customer_id,
            "customer_name": customer_name,
            "file_size": len(raw),
            "total_invoices": len(sections),
            "uploaded_at": now_utc(),
        })

        file_items = 0
        for sec in sections:
            inv_no = sec["invoice_number"]
            products = extract_products(sec["text"])
            for prod in products:
                barcode = generate_barcode()
                doc_item = {
                    "_id": str(uuid.uuid4()),
                    "filename": upload.filename,
                    "invoice": inv_no,
                    "model": prod["model"],
                    "size": prod["size"],
                    "barcode": barcode,
                    "status": "Pending",
                    "date": "-",
                    "customer_id": customer_id,
                    "customer_name": customer_name,
                    "normal_rate": "",
                    "shak": "",
                    "tatriz": "",
                    "amount": "",
                    "created_at": now_utc(),
                }
                try:
                    pdf_viewer_delivery_collection.insert_one(doc_item)
                    file_items += 1
                except Exception:
                    pass

        pdf_files_collection.update_one(
            {"filename": upload.filename},
            {"$set": {"total_delivery_items": file_items}},
        )
        total_files += 1
        total_invoices += len(sections)
        total_items += file_items

    return {
        "total_files_processed": total_files,
        "total_invoices": total_invoices,
        "total_delivery_items": total_items,
    }

# =====================================================
# INVOICE DETAILS for a PDF (used by Print / Print-All)
# =====================================================
@app.get("/api/pdf-invoice-details/{filename}")
def pdf_invoice_details(filename: str):
    if not filename:
        raise HTTPException(status_code=400, detail="filename is required")
    pf = pdf_files_collection.find_one({"filename": filename})
    if not pf:
        raise HTTPException(status_code=404, detail="PDF not found")

    docs = list(pdf_viewer_delivery_collection.find({"filename": filename}))
    grouped = {}
    for d in docs:
        inv = d.get("invoice", "Unknown")
        grouped.setdefault(inv, []).append(d)

    invoices = []
    for inv_no, group in grouped.items():
        products = [{"model": g.get("model", "N/A"), "size": g.get("size", "N/A")} for g in group]
        invoices.append({
            "_id": group[0]["_id"],
            "invoice_number": inv_no,
            "products": products,
            "selected_product": products[0] if products else {"model": "N/A", "size": "N/A"},
            "name": pf.get("filename"),
        })

    return {"filename": filename, "invoices": invoices}

# =====================================================
# PDF WITH FOOTER (single)
# =====================================================
@app.post("/api/pdf-with-footer/{filename}")
def pdf_with_footer(filename: str, payload: PdfFooterRequest):
    pf = pdf_files_collection.find_one({"filename": filename})
    if not pf:
        raise HTTPException(status_code=404, detail="PDF not found")
    return Response(content=b"", media_type="application/pdf")  # placeholder

# =====================================================
# PDF PRINT-ALL WITH SELECTIONS
# =====================================================
@app.post("/api/pdf-print-all-with-selections")
def pdf_print_all_with_selections(payload: PdfPrintAllRequest):
    return Response(content=b"", media_type="application/pdf")  # placeholder

# =====================================================
# DELIVERY ITEMS (Delivery Data page)
# =====================================================
@app.get("/api/pdf-viewer-delivery-items")
def list_delivery_items():
    items = []
    for d in pdf_viewer_delivery_collection.find().sort("created_at", -1):
        items.append({
            "_id": d["_id"],
            "invoice": d.get("invoice", "-"),
            "model": d.get("model", "-"),
            "size": d.get("size", "-"),
            "barcode": d.get("barcode", ""),
            "status": d.get("status", "Pending"),
            "date": d.get("date", "-"),
            "customer_id": d.get("customer_id"),
            "filename": d.get("filename"),
        })
    return resolve_customer_names(items)

@app.post("/api/pdf-viewer-delivery-items/scan")
def scan_delivery_item(payload: BarcodeScan):
    code = (payload.barcode or "").strip()
    if not code:
        raise HTTPException(status_code=400, detail="Barcode is required")
    item = pdf_viewer_delivery_collection.find_one({"barcode": code})
    if not item:
        raise HTTPException(status_code=404, detail="Barcode not found")
    if item.get("status") == "Delivered":
        raise HTTPException(status_code=409, detail="Already delivered")
    now = now_utc()
    pdf_viewer_delivery_collection.update_one(
        {"_id": item["_id"]},
        {"$set": {"status": "Delivered", "date": now.isoformat()}},
    )
    return {"message": "Item marked as Delivered", "barcode": code}

@app.get("/api/pdf-viewer-delivery-items/{item_id}/report")
def get_item_report(item_id: str):
    item = pdf_viewer_delivery_collection.find_one({"_id": item_id})
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    item["customer_name"] = item.get("customer_name") or "Unknown"
    return item

@app.patch("/api/pdf-viewer-delivery-items/{item_id}/report")
def update_item_report(item_id: str, payload: DeliveryReportUpdate):
    item = pdf_viewer_delivery_collection.find_one({"_id": item_id})
    if not item:
        raise HTTPException(status_code=404, detail="Item not found")
    updates = {k: v for k, v in payload.dict().items() if v is not None}
    if updates:
        pdf_viewer_delivery_collection.update_one({"_id": item_id}, {"$set": updates})
    return {"message": "Report updated", "updated": list(updates.keys())}

# =====================================================
# REPORTS (Report page)
# =====================================================
@app.get("/api/pdf-viewer-reports")
def pdf_viewer_reports(status: Optional[str] = "delivered"):
    q = {}
    if status and status.lower() != "all":
        q["status"] = "Delivered" if status.lower() == "delivered" else status
    items = []
    for d in pdf_viewer_delivery_collection.find(q).sort("date", -1):
        items.append({
            "id": d["_id"],
            "invoice": d.get("invoice", "-"),
            "model": d.get("model", "-"),
            "size": d.get("size", "-"),
            "barcode": d.get("barcode", ""),
            "status": d.get("status", "Pending"),
            "deliveredAt": d.get("date", "-"),
            "customer_id": d.get("customer_id"),
            "filename": d.get("filename"),
            "normal_rate": d.get("normal_rate", ""),
            "shak": d.get("shak", ""),
            "tatriz": d.get("tatriz", ""),
            "amount": d.get("amount", ""),
        })
    items = resolve_customer_names(items)
    # re-key customer_name -> customer_name for report type
    for it in items:
        it["customer_name"] = it.get("customer_name", "Unknown")
    return {"items": items}

# =====================================================
# HEALTH
# =====================================================
@app.get("/")
def root():
    return {"status": "ok", "service": "Invoice Management API"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)