import io

import fitz


def merge_pdfs(pdf_files):

    if not pdf_files:
        raise ValueError("No PDF files were provided.")

    merged_document = fitz.open()

    try:

        for pdf_bytes in pdf_files:

            if not pdf_bytes:
                raise ValueError(
                    "One of the uploaded PDF files is empty."
                )

            source_document = fitz.open(
                stream=pdf_bytes,
                filetype="pdf"
            )

            try:

                merged_document.insert_pdf(
                    source_document
                )

            finally:

                source_document.close()


        output_stream = io.BytesIO()

        merged_document.save(
            output_stream
        )

        output_stream.seek(0)

        return output_stream.getvalue()


    finally:

        merged_document.close()