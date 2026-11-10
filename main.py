
import os
import re
import uuid
import logging
from pathlib import Path
from datetime import datetime, timezone
from io import BytesIO
from typing import Optional

import pymupdf
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from bson import ObjectId
from bson.errors import InvalidId
from openpyxl import Workbook

# =========================================================
# CONFIGURATION
# =========================================================

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("invoice-api")

MONGODB_URI = os.getenv("MONGODB_URI") or os.getenv("MONGO_URI")
DB_NAME = os.getenv("DB_NAME", "invoice_delivery")
UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", "uploads"))
MAX_PDF_SIZE_MB = int(os.getenv("MAX_PDF_SIZE_MB", "25"))

if not MONGODB_URI:
    raise RuntimeError("MONGODB_URI is missing.")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

client = MongoClient(
    MONGODB_URI,
    serverSelectionTimeoutMS=10000
)

db = client[DB_NAME]

invoices_collection = db["invoices"]
delivery_collection = db["delivery_items"]

app = FastAPI(
    title="Invoice & Delivery Management API",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================================================
# MODELS
# =========================================================

class DeliveryItemCreate(BaseModel):
    invoice: str
    model: Optional[str] = None
    product_number: Optional[str] = None
    size: Optional[str] = None
    quantity: int = Field(default=1, ge=1)
    barcode: Optional[str] = None


class BarcodeScan(BaseModel):
    barcode: str


class ReviewUpdate(BaseModel):
    invoice: Optional[str] = None
    model: Optional[str] = None
    product_number: Optional[str] = None
    size: Optional[str] = None
    quantity: Optional[int] = Field(default=None, ge=1)


# =========================================================
# COMMON HELPERS
# =========================================================

def now():
    return datetime.now(timezone.utc)


def clean(value):
    if value is None:
        return None

    value = re.sub(r"\s+", " ", str(value)).strip(" \t\r\n:;|")

    if value.lower() in ("", "-", "--", "n/a", "na", "none", "null"):
        return None

    return value


def generate_barcode():
    return f"INV-{uuid.uuid4().hex[:12].upper()}"


def mongo_safe(document):
    if not document:
        return None

    document = dict(document)

    if "_id" in document:
        document["_id"] = str(document["_id"])

    return document


def normalize_header(value):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        str(value or "").lower()
    ).strip()


# =========================================================
# INVOICE NUMBER EXTRACTION
# =========================================================

INVOICE_PATTERNS = [
    r"\b(?:invoice|order)\s*(?:number|no\.?|num|#|id)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9/_\-]{2,})",
    r"\b(?:invoice|order)\s*[:#]\s*([A-Z0-9][A-Z0-9/_\-]{2,})",
    r"\b(?:bill)\s*(?:number|no\.?)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9/_\-]{2,})",
]


def extract_invoice_number(text, filename, page_number):

    for pattern in INVOICE_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            return clean(match.group(1)), "high"

    # Common order format: #294763905
    match = re.search(r"#\s*(\d{6,})\b", text)

    if match:
        return match.group(1), "medium"

    # Do not treat filename as a real invoice number.
    filename = Path(filename).stem

    return (
        f"UNKNOWN-{filename}-P{page_number}",
        "low"
    )


# =========================================================
# PRODUCT TABLE EXTRACTION
# =========================================================

HEADER_FIELDS = {
    "model": {
        "model",
        "model no",
        "model number",
        "style",
        "style no",
        "article",
        "article no",
        "item model"
    },

    "product_number": {
        "product",
        "product no",
        "product number",
        "sku",
        "sku no",
        "item code",
        "product code",
        "article code"
    },

    "size": {
        "size",
        "sz",
        "garment size"
    },

    "quantity": {
        "qty",
        "quantity",
        "pcs",
        "units"
    }
}


def classify_header(value):

    value = normalize_header(value)

    for field, names in HEADER_FIELDS.items():

        if value in names:
            return field

    if value.startswith("model"):
        return "model"

    if value.startswith("size") or value == "sz":
        return "size"

    if value.startswith("qty") or value.startswith("quantity"):
        return "quantity"

    if value.startswith("sku") or value.startswith("product code"):
        return "product_number"

    return None


def extract_table_products(page):

    products = []

    try:
        finder = page.find_tables()
        tables = finder.tables if finder else []
    except Exception:
        return products

    for table in tables:

        try:
            rows = table.extract()
        except Exception:
            continue

        if not rows:
            continue

        header_index = None
        column_map = {}

        for index, row in enumerate(rows[:5]):

            mapping = {}

            for col_index, cell in enumerate(row or []):

                field = classify_header(cell)

                if field and field not in mapping:
                    mapping[field] = col_index

            if (
                "size" in mapping
                and (
                    "model" in mapping
                    or "product_number" in mapping
                )
            ):
                header_index = index
                column_map = mapping
                break

        if header_index is None:
            continue

        for row in rows[header_index + 1:]:

            if not row:
                continue

            joined = " ".join(str(x or "") for x in row).lower()

            if any(
                word in joined
                for word in [
                    "subtotal",
                    "grand total",
                    "total amount",
                    "shipping",
                    "tax"
                ]
            ):
                continue

            item = {}

            for field, col_index in column_map.items():

                if col_index < len(row):

                    value = clean(row[col_index])

                    if value:
                        item[field] = value

            if item.get("model") or item.get("product_number"):
                item.setdefault("quantity", "1")
                products.append(item)

    return products


# =========================================================
# POSITION-BASED EXTRACTION
# =========================================================

def extract_position_products(page):

    try:
        words = page.get_text("words")
    except Exception:
        return []

    if not words:
        return []

    lines = {}

    for word in words:

        x0, y0, x1, y1, token = word[:5]

        key = round(float(y0), 1)

        lines.setdefault(key, []).append(
            (float(x0), float(x1), str(token))
        )

    ordered_lines = []

    for y in sorted(lines):

        row_words = sorted(
            lines[y],
            key=lambda item: item[0]
        )

        ordered_lines.append(
            (y, row_words, " ".join(w[2] for w in row_words))
        )

    header_index = None
    columns = {}

    for index, (_, row_words, _) in enumerate(ordered_lines):

        mapping = {}

        for x0, x1, token in row_words:

            field = classify_header(token)

            if field and field not in mapping:
                mapping[field] = (x0 + x1) / 2

        if (
            "size" in mapping
            and (
                "model" in mapping
                or "product_number" in mapping
            )
        ):
            header_index = index
            columns = mapping
            break

    if header_index is None:
        return []

    sorted_columns = sorted(
        (x, field)
        for field, x in columns.items()
    )

    boundaries = []

    for index, (x, field) in enumerate(sorted_columns):

        left = (
            (sorted_columns[index - 1][0] + x) / 2
            if index else float("-inf")
        )

        right = (
            (x + sorted_columns[index + 1][0]) / 2
            if index + 1 < len(sorted_columns)
            else float("inf")
        )

        boundaries.append((left, right, field))

    products = []

    for _, row_words, line_text in ordered_lines[header_index + 1:]:

        lower = line_text.lower()

        if any(
            word in lower
            for word in [
                "subtotal",
                "grand total",
                "total amount",
                "tax amount"
            ]
        ):
            continue

        cells = {field: [] for field in columns}

        for x0, x1, token in row_words:

            center = (x0 + x1) / 2

            for left, right, field in boundaries:

                if left <= center < right:
                    cells[field].append(token)
                    break

        item = {
            field: clean(" ".join(tokens))
            for field, tokens in cells.items()
        }

        item = {
            key: value
            for key, value in item.items()
            if value
        }

        if item.get("model") or item.get("product_number"):

            item.setdefault("quantity", "1")
            products.append(item)

    return products


# =========================================================
# LABELED TEXT EXTRACTION
# =========================================================

def extract_labeled_products(text):

    model_pattern = (
        r"\b(?:model|style|article)\s*"
        r"(?:no\.?|number|#|code)?\s*[:#\-]?\s*"
        r"([A-Z0-9][A-Z0-9./_\-]{0,20})"
    )

    model_matches = list(
        re.finditer(model_pattern, text, re.IGNORECASE)
    )

    product_pattern = (
        r"\b(?:sku|product\s*(?:no\.?|number|code)|item\s*code)"
        r"\s*[:#\-]?\s*([A-Z0-9][A-Z0-9./_\-]{2,})"
    )

    product_matches = list(
        re.finditer(product_pattern, text, re.IGNORECASE)
    )

    products = []

    for index, match in enumerate(model_matches):

        start = match.start()

        end = (
            model_matches[index + 1].start()
            if index + 1 < len(model_matches)
            else len(text)
        )

        segment = text[start:end]

        item = {
            "model": clean(match.group(1)),
            "quantity": "1"
        }

        size_match = re.search(
            r"\b(?:size|sz)\s*[:#\-]?\s*"
            r"([A-Z0-9][A-Z0-9./_\-]{0,10})",
            segment,
            re.IGNORECASE
        )

        product_match = re.search(
            product_pattern,
            segment,
            re.IGNORECASE
        )

        qty_match = re.search(
            r"\b(?:qty|quantity)\s*[:#\-]?\s*(\d+)\b",
            segment,
            re.IGNORECASE
        )

        if size_match:
            item["size"] = clean(size_match.group(1))

        if product_match:
            item["product_number"] = clean(product_match.group(1))

        if qty_match:
            item["quantity"] = qty_match.group(1)

        products.append(item)

    if not products and product_matches:

        for match in product_matches:

            products.append({
                "model": None,
                "product_number": clean(match.group(1)),
                "size": None,
                "quantity": "1"
            })

    return products


# =========================================================
# EXTRACTION PIPELINE
# =========================================================

def deduplicate_products(products):

    result = []
    seen = set()

    for item in products:

        normalized = {
            "model": clean(item.get("model")),
            "product_number": clean(item.get("product_number")),
            "size": clean(item.get("size")),
            "quantity": clean(item.get("quantity")) or "1"
        }

        key = tuple(
            (normalized.get(field) or "").lower()
            for field in [
                "model",
                "product_number",
                "size",
                "quantity"
            ]
        )

        if (
            key not in seen
            and any(
                normalized.get(field)
                for field in ["model", "product_number", "size"]
            )
        ):
            seen.add(key)
            result.append(normalized)

    return result


def extract_products(page, text):

    products = extract_table_products(page)
    method = "table"

    if not products:
        products = extract_position_products(page)
        method = "word_positions"

    if not products:
        products = extract_labeled_products(text)
        method = "labeled_text"

    return deduplicate_products(products), method


# =========================================================
# VALIDATION
# =========================================================

def validate_product(product, invoice_number, invoice_confidence):

    missing = []

    if (
        not invoice_number
        or invoice_confidence == "low"
        or invoice_number.startswith("UNKNOWN-")
    ):
        missing.append("invoice_number")

    if not product.get("model"):
        missing.append("model")

    if not product.get("size"):
        missing.append("size")

    if not product.get("product_number"):
        missing.append("product_number")

    try:
        quantity = int(product.get("quantity") or 1)

        if quantity < 1:
            raise ValueError

    except (TypeError, ValueError):
        missing.append("quantity")

    return missing, ("high" if not missing else "low")


# =========================================================
# PDF PROCESSING
# =========================================================

def process_pdf(pdf_path, original_filename):

    created_invoice_ids = []
    created_item_ids = []

    try:
        document = pymupdf.open(pdf_path)
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid or unreadable PDF: {exc}"
        )

    try:

        total_pages = len(document)

        if total_pages == 0:
            raise HTTPException(
                status_code=400,
                detail="PDF contains no pages."
            )

        for page_index, page in enumerate(document):

            page_number = page_index + 1

            text = page.get_text("text") or ""

            invoice_number, invoice_confidence = extract_invoice_number(
                text,
                original_filename,
                page_number
            )

            products, extraction_method = extract_products(page, text)

            invoice_document = {
                "invoice_number": invoice_number,
                "original_filename": original_filename,
                "stored_file": str(pdf_path),
                "page_number": page_number,
                "total_pages": total_pages,
                "extracted_text": text,
                "extraction_method": extraction_method,
                "invoice_confidence": invoice_confidence,
                "review_required": invoice_confidence == "low",
                "review_status": (
                    "Pending"
                    if invoice_confidence == "low"
                    else "Not Required"
                ),
                "created_at": now()
            }

            invoice_result = invoices_collection.insert_one(
                invoice_document
            )

            invoice_id = str(invoice_result.inserted_id)

            created_invoice_ids.append(invoice_id)

            if not products:

                products = [{
                    "model": None,
                    "product_number": None,
                    "size": None,
                    "quantity": "1"
                }]

                extraction_method = "no_product_rows_found"

            for product_index, product in enumerate(products, start=1):

                missing_fields, confidence = validate_product(
                    product,
                    invoice_number,
                    invoice_confidence
                )

                try:
                    quantity = max(
                        1,
                        int(product.get("quantity") or 1)
                    )
                except (TypeError, ValueError):
                    quantity = 1

                delivery_document = {
                    "invoice_id": invoice_id,

                    # Existing frontend field preserved.
                    "invoice": invoice_number,
                    "invoice_number": invoice_number,

                    "source_filename": original_filename,
                    "source_page": page_number,
                    "product_index": product_index,

                    "model": product.get("model"),
                    "product_number": product.get("product_number"),
                    "size": product.get("size"),
                    "quantity": quantity,

                    "barcode": generate_barcode(),

                    "status": "Pending",
                    "delivered_at": None,

                    "extraction_method": extraction_method,
                    "confidence": confidence,
                    "missing_fields": missing_fields,

                    "review_required": bool(missing_fields),
                    "review_status": (
                        "Pending"
                        if missing_fields
                        else "Not Required"
                    ),

                    "created_at": now(),
                    "updated_at": now()
                }

                result = delivery_collection.insert_one(
                    delivery_document
                )

                created_item_ids.append(str(result.inserted_id))

            has_review_items = delivery_collection.find_one({
                "invoice_id": invoice_id,
                "review_required": True
            })

            if has_review_items:

                invoices_collection.update_one(
                    {"_id": invoice_result.inserted_id},
                    {
                        "$set": {
                            "review_required": True,
                            "review_status": "Pending"
                        }
                    }
                )

    finally:
        document.close()

    review_count = delivery_collection.count_documents({
        "_id": {
            "$in": [
                ObjectId(item_id)
                for item_id in created_item_ids
            ]
        },
        "review_required": True
    })

    return {
        "message": "PDF processed successfully.",
        "filename": original_filename,
        "total_pages": total_pages,
        "invoices_created": len(created_invoice_ids),
        "delivery_items_created": len(created_item_ids),
        "review_required_count": review_count,
        "invoice_ids": created_invoice_ids,
        "delivery_item_ids": created_item_ids
    }


# =========================================================
# REPORT HELPER
# =========================================================

def get_report_items(status="all"):

    query = {}

    if status == "delivered":
        query["status"] = "Delivered"

    elif status == "undelivered":
        query["status"] = {"$ne": "Delivered"}

    elif status != "all":
        raise HTTPException(
            status_code=400,
            detail="status must be all, delivered or undelivered"
        )

    return [
        mongo_safe(item)
        for item in delivery_collection.find(query).sort(
            "created_at", -1
        )
    ]


# =========================================================
# BASIC ROUTES
# =========================================================

@app.get("/")
def root():

    return {
        "message": "Invoice & Delivery Management API is running."
    }


@app.get("/health")
def health():

    try:
        client.admin.command("ping")

        return {
            "status": "ok",
            "database": "connected"
        }

    except Exception as exc:

        raise HTTPException(
            status_code=503,
            detail=f"Database unavailable: {exc}"
        )


# =========================================================
# PDF UPLOAD
# =========================================================

@app.post("/upload_pdf")
async def upload_pdf(invoice: UploadFile = File(...)):

    filename = Path(
        invoice.filename or "invoice.pdf"
    ).name

    if not filename.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Only PDF files are allowed."
        )

    content = await invoice.read()

    if not content:
        raise HTTPException(
            status_code=400,
            detail="Uploaded PDF is empty."
        )

    if len(content) > MAX_PDF_SIZE_MB * 1024 * 1024:
        raise HTTPException(
            status_code=413,
            detail=f"PDF exceeds {MAX_PDF_SIZE_MB} MB."
        )

    pdf_path = UPLOAD_DIR / f"{uuid.uuid4().hex}.pdf"

    pdf_path.write_bytes(content)

    try:

        return process_pdf(pdf_path, filename)

    except HTTPException:
        pdf_path.unlink(missing_ok=True)
        raise

    except PyMongoError as exc:

        logger.exception("MongoDB processing error")

        raise HTTPException(
            status_code=503,
            detail=f"Database operation failed: {exc}"
        )

    except Exception as exc:

        logger.exception("PDF processing error")

        raise HTTPException(
            status_code=500,
            detail=f"PDF processing failed: {exc}"
        )


# =========================================================
# INVOICE APIs
# =========================================================

@app.get("/api/invoices")
def get_invoices():

    return [
        mongo_safe(item)
        for item in invoices_collection.find().sort(
            "created_at", -1
        )
    ]


@app.get("/api/invoices/{invoice_id}")
def get_invoice_detail(invoice_id: str):

    try:
        object_id = ObjectId(invoice_id)
    except InvalidId:
        raise HTTPException(
            status_code=400,
            detail="Invalid invoice ID."
        )

    invoice_document = invoices_collection.find_one({
        "_id": object_id
    })

    if not invoice_document:
        raise HTTPException(
            status_code=404,
            detail="Invoice not found."
        )

    items = [
        mongo_safe(item)
        for item in delivery_collection.find({
            "invoice_id": invoice_id
        }).sort("product_index", 1)
    ]

    return {
        "invoice": mongo_safe(invoice_document),
        "delivery_items": items
    }


# =========================================================
# DELIVERY APIs
# =========================================================

@app.get("/api/delivery-items")
def get_delivery_items():

    return [
        mongo_safe(item)
        for item in delivery_collection.find().sort(
            "created_at", -1
        )
    ]


@app.post("/api/delivery-items")
def create_delivery_item(item: DeliveryItemCreate):

    barcode = item.barcode or generate_barcode()

    if delivery_collection.find_one({"barcode": barcode}):
        raise HTTPException(
            status_code=409,
            detail="Barcode already exists."
        )

    document = {
        "invoice": item.invoice,
        "invoice_number": item.invoice,
        "model": item.model,
        "product_number": item.product_number,
        "size": item.size,
        "quantity": item.quantity,
        "barcode": barcode,
        "status": "Pending",
        "delivered_at": None,
        "review_required": False,
        "review_status": "Not Required",
        "created_at": now(),
        "updated_at": now()
    }

    result = delivery_collection.insert_one(document)

    document["_id"] = str(result.inserted_id)

    return {
        "message": "Delivery item created.",
        "item": document
    }


@app.post("/api/delivery-items/scan")
def scan_barcode(data: BarcodeScan):

    barcode = data.barcode.strip()

    if not barcode:
        raise HTTPException(
            status_code=400,
            detail="Barcode is required."
        )

    item = delivery_collection.find_one({
        "barcode": barcode
    })

    if not item:
        raise HTTPException(
            status_code=404,
            detail="Barcode not found."
        )

    if item.get("status") == "Delivered":

        return {
            "message": "Item already delivered.",
            "item": mongo_safe(item)
        }

    delivery_collection.update_one(
        {"_id": item["_id"]},
        {
            "$set": {
                "status": "Delivered",
                "delivered_at": now(),
                "updated_at": now()
            }
        }
    )

    updated = delivery_collection.find_one({
        "_id": item["_id"]
    })

    return {
        "message": "Item marked as delivered.",
        "item": mongo_safe(updated)
    }


# =========================================================
# REPORT APIs
# =========================================================

@app.get("/api/reports")
def get_reports(status: str = "all"):

    return get_report_items(status)


@app.get("/api/invoices/export")
def export_invoices(status: str = "all"):

    items = get_report_items(status)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Delivery Report"

    sheet.append([
        "Invoice Number",
        "Model",
        "Product Number",
        "Size",
        "Quantity",
        "Barcode",
        "Status",
        "Delivered At",
        "Review Required",
        "Missing Fields",
        "Source Page"
    ])

    for item in items:

        delivered_at = item.get("delivered_at")

        if isinstance(delivered_at, datetime):
            delivered_at = delivered_at.isoformat()

        sheet.append([
            item.get("invoice_number") or item.get("invoice"),
            item.get("model"),
            item.get("product_number"),
            item.get("size"),
            item.get("quantity", 1),
            item.get("barcode"),
            item.get("status"),
            delivered_at,
            item.get("review_required", False),
            ", ".join(item.get("missing_fields", [])),
            item.get("source_page")
        ])

    for column in sheet.columns:

        max_length = max(
            len(str(cell.value or ""))
            for cell in column
        )

        sheet.column_dimensions[
            column[0].column_letter
        ].width = min(max_length + 2, 35)

    output = BytesIO()

    workbook.save(output)
    output.seek(0)

    filename = f"delivery_report_{status}.xlsx"

    return StreamingResponse(
        output,
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


# =========================================================
# MANUAL REVIEW APIs
# =========================================================

@app.get("/api/delivery-items/review")
def get_items_for_review():

    query = {
        "review_required": True,
        "review_status": "Pending"
    }

    return [
        mongo_safe(item)
        for item in delivery_collection.find(query).sort(
            "created_at", -1
        )
    ]


@app.patch("/api/delivery-items/{item_id}/review")
def review_delivery_item(
    item_id: str,
    update: ReviewUpdate
):

    try:
        object_id = ObjectId(item_id)
    except InvalidId:
        raise HTTPException(
            status_code=400,
            detail="Invalid delivery item ID."
        )

    item = delivery_collection.find_one({
        "_id": object_id
    })

    if not item:
        raise HTTPException(
            status_code=404,
            detail="Delivery item not found."
        )

    changes = update.model_dump(exclude_unset=True)

    merged = {
        **item,
        **changes
    }

    missing = []

    if not merged.get("invoice_number") and not merged.get("invoice"):
        missing.append("invoice_number")

    if not merged.get("model"):
        missing.append("model")

    if not merged.get("size"):
        missing.append("size")

    if not merged.get("product_number"):
        missing.append("product_number")

    changes["missing_fields"] = missing
    changes["review_required"] = bool(missing)

    changes["review_status"] = (
        "Pending" if missing else "Reviewed"
    )

    changes["updated_at"] = now()

    delivery_collection.update_one(
        {"_id": object_id},
        {"$set": changes}
    )

    updated = delivery_collection.find_one({
        "_id": object_id
    })

    return {
        "message": "Review saved.",
        "item": mongo_safe(updated)
    }


# =========================================================
# START SERVER
# =========================================================

if __name__ == "__main__":

    import uvicorn

    port = int(os.getenv("PORT", "8000"))

    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port
    )
