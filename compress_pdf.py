import io

import fitz


def compress_pdf(
    pdf_bytes,
    compression_level
):

    source_document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:

        output_buffer = io.BytesIO()


        # =====================================
        # RECOMMENDED
        # =====================================

        if compression_level == "recommended":

            source_document.save(
                output_buffer,
                garbage=4,
                deflate=True,
                clean=True
            )


        # =====================================
        # MAXIMUM COMPRESSION
        # =====================================

        elif compression_level == "maximum":

            source_document.save(
                output_buffer,
                garbage=4,
                deflate=True,
                clean=True,
                deflate_images=True,
                deflate_fonts=True
            )


        # =====================================
        # HIGH QUALITY
        # =====================================

        elif compression_level == "quality":

            source_document.save(
                output_buffer,
                garbage=3,
                deflate=True,
                clean=True,
                deflate_images=False
            )


        else:

            raise ValueError(
                "Invalid compression level."
            )


        output_buffer.seek(0)

        return output_buffer.getvalue()


    finally:

        source_document.close()
        