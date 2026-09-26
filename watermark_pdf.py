import io
import fitz


def add_watermark(
    pdf_bytes,
    watermark_text,
    position="center"
):
    if not watermark_text or not watermark_text.strip():
        raise ValueError("Watermark text cannot be empty.")

    if position not in ("top", "center", "bottom"):
        raise ValueError("Invalid watermark position.")

    input_pdf = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        for page in input_pdf:
            page_rect = page.rect

            font_size = 36
            margin = 36

            text_width = fitz.get_text_length(
                watermark_text,
                fontname="helv",
                fontsize=font_size
            )

            max_width = page_rect.width - (margin * 2)

            if text_width > max_width:
                font_size = max(
                    8,
                    font_size * max_width / text_width
                )

                text_width = fitz.get_text_length(
                    watermark_text,
                    fontname="helv",
                    fontsize=font_size
                )

            x = page_rect.x0 + (
                page_rect.width - text_width
            ) / 2

            if position == "top":
                y = page_rect.y0 + 70

            elif position == "bottom":
                y = page_rect.y1 - 50

            else:
                y = page_rect.y0 + (
                    page_rect.height / 2
                )

            page.insert_text(
                fitz.Point(x, y),
                watermark_text,
                fontsize=font_size,
                fontname="helv",
                color=(0.65, 0.65, 0.65),
                overlay=True
            )

        output = io.BytesIO()

        input_pdf.save(
            output,
            garbage=4,
            deflate=True
        )

        return output.getvalue()

    finally:
        input_pdf.close()