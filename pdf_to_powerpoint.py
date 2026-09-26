import io

import fitz
from pptx import Presentation


def convert_pdf_to_powerpoint(pdf_bytes):

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:

        if len(document) == 0:
            raise ValueError(
                "The uploaded PDF contains no pages."
            )

        presentation = Presentation()

        # Remove the default blank slide.
        if len(presentation.slides) > 0:

            slide_id = presentation.slides._sldIdLst[0]

            presentation.part.drop_rel(
                slide_id.rId
            )

            del presentation.slides._sldIdLst[0]

        # -----------------------------------------
        # Set PowerPoint slide size from first
        # PDF page aspect ratio
        # -----------------------------------------

        first_page = document.load_page(0)

        pdf_width = first_page.rect.width
        pdf_height = first_page.rect.height

        max_width = 13.333
        max_height = 7.5

        pdf_ratio = (
            pdf_width / pdf_height
        )

        if pdf_ratio >= (
            max_width / max_height
        ):

            slide_width_inches = max_width

            slide_height_inches = (
                max_width / pdf_ratio
            )

        else:

            slide_height_inches = max_height

            slide_width_inches = (
                max_height * pdf_ratio
            )

        presentation.slide_width = int(
            slide_width_inches * 914400
        )

        presentation.slide_height = int(
            slide_height_inches * 914400
        )

        # -----------------------------------------
        # Convert every PDF page into one
        # PowerPoint slide
        # -----------------------------------------

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

            image_bytes = pixmap.tobytes(
                "jpeg",
                jpg_quality=90
            )

            image_stream = io.BytesIO(
                image_bytes
            )

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[6]
            )

            slide.shapes.add_picture(
                image_stream,
                0,
                0,
                width=presentation.slide_width,
                height=presentation.slide_height
            )

        # -----------------------------------------
        # Save PowerPoint to memory
        # -----------------------------------------

        output_buffer = io.BytesIO()

        presentation.save(
            output_buffer
        )

        output_buffer.seek(0)

        return output_buffer.getvalue()

    finally:

        document.close()