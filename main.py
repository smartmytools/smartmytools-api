from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from typing import List

from excel_to_pdf import convert_excel_to_pdf
from pdf_to_word import convert_pdf_to_word
from word_to_pdf import convert_word_to_pdf
from pdf_to_jpg import convert_pdf_to_jpg
from merge_pdf import merge_pdfs
from compress_pdf import compress_pdf
from split_pdf import (
    split_every_page,
    extract_selected_pages
)

from pdf_to_excel import convert_pdf_to_excel

import io


app = FastAPI(
    title="SmartMyTools API",
    version="1.0.0"
)

@app.post("/api/excel-to-pdf")
async def excel_to_pdf(file: UploadFile = File(...)):
    try:
        excel_bytes = await file.read()

        if not file.filename.lower().endswith(
            (".xlsx", ".xlsm")
        ):
            raise HTTPException(
                status_code=400,
                detail="Please upload an Excel file (.xlsx or .xlsm)."
            )

        pdf_bytes = convert_excel_to_pdf(excel_bytes)

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": (
                    'attachment; filename="converted.pdf"'
                )
            }
        )

    except HTTPException:
        raise

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"Excel to PDF conversion failed: {error}"
        )

@app.post("/api/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()

        excel_bytes = convert_pdf_to_excel(pdf_bytes)

        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": (
                    'attachment; filename="converted.xlsx"'
                )
            }
        )

    except ValueError as error:
        raise HTTPException(
            status_code=400,
            detail=str(error)
        )

    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"PDF to Excel conversion failed: {error}"
        )
# =========================================
# CORS
# =========================================

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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

    # -----------------------------------------
    # Check file name
    # -----------------------------------------

    if not file.filename:
        raise HTTPException(
            status_code=400,
            detail="No file was selected."
        )


    # -----------------------------------------
    # Check PDF extension
    # -----------------------------------------

    if not file.filename.lower().endswith(".pdf"):

        raise HTTPException(
            status_code=400,
            detail="Please upload a PDF file."
        )


    try:

        # -----------------------------------------
        # Read uploaded PDF
        # -----------------------------------------

        pdf_bytes = await file.read()


        if not pdf_bytes:

            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )


        # -----------------------------------------
        # Convert PDF → DOCX
        # -----------------------------------------

        docx_bytes = convert_pdf_to_word(
            pdf_bytes
        )


        # -----------------------------------------
        # Return Word document
        # -----------------------------------------

        output_filename = (
            file.filename.rsplit(
                ".",
                1
            )[0]
            + ".docx"
        )


        return StreamingResponse(

            io.BytesIO(
                docx_bytes
            ),

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
                "Please try another file."
            )
        )

# =========================================
# MERGE PDF
# =========================================

@app.post("/api/merge-pdf")
async def merge_pdf(
    files: List[UploadFile] = File(...)
):

    if len(files) < 2:

        raise HTTPException(
            status_code=400,
            detail="Please upload at least two PDF files."
        )


    try:

        pdf_files = []


        for file in files:

            if not file.filename:

                raise HTTPException(
                    status_code=400,
                    detail="One of the files has no filename."
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
                    detail=(
                        f"The file '{file.filename}' is empty."
                    )
                )


            pdf_files.append(
                pdf_bytes
            )


        merged_pdf_bytes = merge_pdfs(
            pdf_files
        )


        return StreamingResponse(

            io.BytesIO(
                merged_pdf_bytes
            ),

            media_type="application/pdf",

            headers={
                "Content-Disposition":
                    'attachment; filename="smartmytools-merged.pdf"'
            }
        )


    except HTTPException:
        raise


    except Exception as error:

        print(
            "Merge PDF error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail=(
                "Unable to merge these PDF files. "
                "Please make sure they are valid PDFs."
            )
        )

        # =========================================
# SPLIT EVERY PAGE
# =========================================

@app.post("/api/split-every-page")
async def split_every_page_api(
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

        zip_bytes = split_every_page(
            pdf_bytes,
            file.filename
        )

        return StreamingResponse(
            io.BytesIO(zip_bytes),
            media_type="application/zip",
            headers={
                "Content-Disposition":
                    'attachment; filename="split-pdf-pages.zip"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

        print(
            "Split every page error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail="Unable to split this PDF."
        )


# =========================================
# EXTRACT SELECTED PAGES
# =========================================

@app.post("/api/extract-selected-pages")
async def extract_selected_pages_api(
    file: UploadFile = File(...),
    pages: str = Form(...)
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

        selected_pages = [
            int(page.strip())
            for page in pages.split(",")
            if page.strip()
        ]

        if not selected_pages:

            raise HTTPException(
                status_code=400,
                detail="Please select at least one page."
            )

        pdf_bytes = await file.read()

        if not pdf_bytes:

            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )

        extracted_pdf_bytes = extract_selected_pages(
            pdf_bytes,
            selected_pages
        )

        return StreamingResponse(
            io.BytesIO(extracted_pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition":
                    'attachment; filename="extracted-pages.pdf"'
            }
        )

    except HTTPException:
        raise

    except ValueError as error:

        raise HTTPException(
            status_code=400,
            detail=str(error)
        )

    except Exception as error:

        print(
            "Extract selected pages error:",
            repr(error)
        )

        raise HTTPException(
            status_code=500,
            detail="Unable to extract the selected pages."
        )

        # =========================================
# COMPRESS PDF
# =========================================

@app.post("/api/compress-pdf")
async def compress_pdf_api(
    file: UploadFile = File(...),
    compression_level: str = Form(...)
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


        allowed_levels = [
            "recommended",
            "maximum",
            "quality"
        ]


        if compression_level not in allowed_levels:

            raise HTTPException(
                status_code=400,
                detail="Invalid compression level."
            )


        pdf_bytes = await file.read()


        if not pdf_bytes:

            raise HTTPException(
                status_code=400,
                detail="The uploaded PDF is empty."
            )


        compressed_pdf_bytes = compress_pdf(
            pdf_bytes,
            compression_level
        )


        return StreamingResponse(

            io.BytesIO(
                compressed_pdf_bytes
            ),

            media_type="application/pdf",

            headers={
                "Content-Disposition":
                    'attachment; filename="compressed.pdf"'
            }

        )


    except HTTPException:
        raise


    except ValueError as error:

        raise HTTPException(
            status_code=400,
            detail=str(error)
        )


    except Exception as error:

        print(
            "Compress PDF error:",
            repr(error)
        )


        raise HTTPException(
            status_code=500,
            detail="Unable to compress this PDF."
        )

        # =========================================
# PDF TO JPG
# =========================================

@app.post("/api/pdf-to-jpg")
async def pdf_to_jpg_api(
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


        jpg_zip_bytes = convert_pdf_to_jpg(
            pdf_bytes
        )


        return StreamingResponse(

            io.BytesIO(
                jpg_zip_bytes
            ),

            media_type="application/zip",

            headers={
                "Content-Disposition":
                    'attachment; filename="pdf-pages-jpg.zip"'
            }

        )


    except HTTPException:
        raise


    except Exception as error:

        print(
            "PDF to JPG error:",
            repr(error)
        )


        raise HTTPException(
            status_code=500,
            detail="Unable to convert this PDF to JPG."
        )