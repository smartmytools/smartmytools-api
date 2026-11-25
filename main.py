from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from typing import List
import io

from excel_to_pdf import convert_excel_to_pdf
from pdf_to_excel import convert_pdf_to_excel
from pdf_to_powerpoint import convert_pdf_to_powerpoint
from page_number import add_page_numbers
from pdf_to_word import convert_pdf_to_word
from word_to_pdf import convert_word_to_pdf


app = FastAPI(
    title="SmartMyTools API",
    version="1.0.0"
)


# =========================================
# CORS
# =========================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)


# =========================================
# HOME
# =========================================

@app.get("/")
def root():

    return {
        "success": True,
        "message": "SmartMyTools backend is working!"
    }


# =========================================
# HEALTH CHECK
# =========================================

@app.get("/api/health")
def health():

    return {
        "success": True,
        "status": "healthy"
    }


# =========================================
# PDF → WORD
# =========================================

@app.post("/api/pdf-to-word")
async def pdf_to_word(
    file: UploadFile = File(...)
):

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was selected."
        )

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PDF file."
        )

    try:

        pdf_bytes = await file.read()

        if not pdf_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )

        docx_bytes = convert_pdf_to_word(
            pdf_bytes
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + ".docx"
        )

        return StreamingResponse(
            io.BytesIO(docx_bytes),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "PDF → Word error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to convert this PDF. "
                "Please try another PDF."
            )
        )


# =========================================
# WORD → PDF
# =========================================

@app.post("/api/word-to-pdf")
async def word_to_pdf(
    file: UploadFile = File(...)
):

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was selected."
        )

    if not (
        file.filename.lower().endswith(".doc")
        or
        file.filename.lower().endswith(".docx")
    ):
        raise HTTPException(
            status_code=400,
            detail="Please upload a DOC or DOCX file."
        )

    try:

        word_bytes = await file.read()

        if not word_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded Word document is empty."
            )

        pdf_bytes = convert_word_to_pdf(
            word_bytes,
            file.filename
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + ".pdf"
        )

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "Word → PDF error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to convert this Word document. "
                "Please try another document."
            )
        )


# =========================================
# EXCEL → PDF
# =========================================

@app.post("/api/excel-to-pdf")
async def excel_to_pdf(
    file: UploadFile = File(...)
):

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was selected."
        )

    if not file.filename.lower().endswith(
        (".xlsx", ".xlsm")
    ):
        raise HTTPException(
            status_code=400,
            detail="Please upload an Excel file."
        )

    try:

        excel_bytes = await file.read()

        if not excel_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded Excel file is empty."
            )

        pdf_bytes = convert_excel_to_pdf(
            excel_bytes
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + ".pdf"
        )

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "Excel → PDF error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to convert this Excel file. "
                "Please try another file."
            )
        )


# =========================================
# PDF → EXCEL
# =========================================

@app.post("/api/pdf-to-excel")
async def pdf_to_excel(
    file: UploadFile = File(...)
):

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was selected."
        )

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PDF file."
        )

    try:

        pdf_bytes = await file.read()

        if not pdf_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )

        excel_bytes = convert_pdf_to_excel(
            pdf_bytes
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + ".xlsx"
        )

        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "PDF → Excel error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to convert this PDF. "
                "Please try another PDF."
            )
        )


# =========================================
# PDF → POWERPOINT
# =========================================

@app.post("/api/pdf-to-powerpoint")
async def pdf_to_powerpoint(
    file: UploadFile = File(...)
):

    try:

        if not file.filename:
            raise HTTPException(
                status_code=400,
                detail="No PDF file was provided."
            )

        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(
                status_code=400,
                detail="Only PDF files are allowed."
            )

        pdf_bytes = await file.read()

        if not pdf_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )

        powerpoint_bytes = convert_pdf_to_powerpoint(
            pdf_bytes
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + ".pptx"
        )

        return StreamingResponse(
            io.BytesIO(powerpoint_bytes),
            media_type=(
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            ),
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "PDF → PowerPoint error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to convert this PDF to PowerPoint. "
                "Please try another PDF."
            )
        )


# =========================================
# ADD PAGE NUMBERS TO PDF
# =========================================

@app.post("/api/page-number")
async def page_number_endpoint(
    file: UploadFile = File(...),
    position: str = Form("bottom"),
    alignment: str = Form("center"),
    starting_number: int = Form(1)
):

    try:

        if not file.filename:
            raise HTTPException(
                status_code=400,
                detail="No PDF file was provided."
            )

        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(
                status_code=400,
                detail="Only PDF files are allowed."
            )

        pdf_bytes = await file.read()

        if not pdf_bytes:
            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )

        numbered_pdf_bytes = add_page_numbers(
            pdf_bytes,
            position,
            alignment,
            starting_number
        )

        output_filename = (
            file.filename.rsplit(".", 1)[0]
            + "-numbered.pdf"
        )

        return StreamingResponse(
            io.BytesIO(numbered_pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition":
                    f'attachment; filename="{output_filename}"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "Page Number PDF error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to add page numbers to this PDF. "
                "Please try another PDF."
            )
        )


