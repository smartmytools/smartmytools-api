import io
import zipfile

import fitz


def split_every_page(pdf_bytes, original_filename):

    source_document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:

        zip_buffer = io.BytesIO()


        with zipfile.ZipFile(
            zip_buffer,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zip_file:

            base_name = original_filename.rsplit(
                ".",
                1
            )[0]


            for page_index in range(
                len(source_document)
            ):

                new_document = fitz.open()

                try:

                    new_document.insert_pdf(
                        source_document,
                        from_page=page_index,
                        to_page=page_index
                    )


                    page_buffer = io.BytesIO()


                    new_document.save(
                        page_buffer
                    )


                    page_file_name = (
                        f"{base_name} - "
                        f"Page {page_index + 1}.pdf"
                    )


                    zip_file.writestr(
                        page_file_name,
                        page_buffer.getvalue()
                    )

                finally:

                    new_document.close()


        zip_buffer.seek(0)

        return zip_buffer.getvalue()


    finally:

        source_document.close()


def extract_selected_pages(
    pdf_bytes,
    selected_pages
):

    source_document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    extracted_document = fitz.open()

    try:

        total_pages = len(
            source_document
        )


        for page_number in selected_pages:

            page_index = page_number - 1


            if (
                page_index < 0
                or page_index >= total_pages
            ):

                raise ValueError(
                    f"Page {page_number} "
                    "does not exist."
                )


            extracted_document.insert_pdf(
                source_document,
                from_page=page_index,
                to_page=page_index
            )


        output_buffer = io.BytesIO()


        extracted_document.save(
            output_buffer
        )


        output_buffer.seek(0)

        return output_buffer.getvalue()


    finally:

        source_document.close()

        extracted_document.close()