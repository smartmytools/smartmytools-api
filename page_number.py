import io

import fitz


def add_page_numbers(
    pdf_bytes,
    position="bottom",
    alignment="center",
    starting_number=1
):
    input_pdf = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        page_number = int(starting_number)

        if page_number < 1:
            raise ValueError(
                "Starting number must be at least 1."
            )

        for page in input_pdf:
            page_rect = page.rect

            margin_x = 36
            margin_y = 30

            font_size = 10

            text = str(page_number)

            text_width = fitz.get_text_length(
                text,
                fontname="helv",
                fontsize=font_size
            )

            if alignment == "left":
                x = page_rect.x0 + margin_x

            elif alignment == "right":
                x = (
                    page_rect.x1
                    - margin_x
                    - text_width
                )

            else:
                x = (
                    page_rect.x0
                    + (
                        page_rect.width
                        - text_width
                    ) / 2
                )

            if position == "top":
                y = page_rect.y0 + margin_y

            else:
                y = page_rect.y1 - margin_y

            page.insert_text(
                fitz.Point(x, y),
                text,
                fontsize=font_size,
                fontname="helv",
                color=(0, 0, 0)
            )

            page_number += 1

        output = io.BytesIO()

        input_pdf.save(
            output,
            garbage=4,
            deflate=True
        )

        return output.getvalue()

    finally:
        input_pdf.close()