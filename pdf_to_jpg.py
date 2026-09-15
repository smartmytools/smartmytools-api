import io
import zipfile

import fitz


def convert_pdf_to_jpg(
    pdf_bytes,
    quality=90
):

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:

        output_buffer = io.BytesIO()


        with zipfile.ZipFile(
            output_buffer,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            for page_number in range(
                len(document)
            ):

                page = document.load_page(
                    page_number
                )


                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(
                        2,
                        2
                    ),
                    alpha=False
                )


                jpg_bytes = pixmap.tobytes(
                    "jpeg",
                    jpg_quality=quality
                )


                zip_file.writestr(
                    f"page-{page_number + 1}.jpg",
                    jpg_bytes
                )


        output_buffer.seek(0)

        return output_buffer.getvalue()


    finally:

        document.close()