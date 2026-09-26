import io
import fitz


def rotate_pdf(pdf_bytes, rotation_angle, rotation_scope="all"):
    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        if len(document) == 0:
            raise ValueError("The uploaded PDF contains no pages.")

        rotation_angle = int(rotation_angle)

        if rotation_angle not in (90, 180, 270):
            raise ValueError("Invalid rotation angle.")

        for page_number in range(len(document)):

            page = document.load_page(page_number)

            current_rotation = page.rotation

            new_rotation = (
                current_rotation + rotation_angle
            ) % 360

            page.set_rotation(new_rotation)

        output_buffer = io.BytesIO()

        document.save(
            output_buffer,
            garbage=4,
            deflate=True
        )

        output_buffer.seek(0)

        return output_buffer.getvalue()

    finally:
        document.close()