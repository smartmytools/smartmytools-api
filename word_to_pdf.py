import os
import io
import tempfile
import subprocess


def convert_word_to_pdf(word_bytes, original_filename):

    if not word_bytes:
        raise ValueError("The Word document is empty.")

    # Get file extension
    extension = os.path.splitext(
        original_filename
    )[1].lower()

    if extension not in [".doc", ".docx"]:
        raise ValueError(
            "Please upload a DOC or DOCX file."
        )

    # Create temporary working folder
    with tempfile.TemporaryDirectory() as temp_dir:

        # Input file path
        input_path = os.path.join(
            temp_dir,
            "input" + extension
        )

        # Output folder
        output_dir = os.path.join(
            temp_dir,
            "output"
        )

        os.makedirs(
            output_dir,
            exist_ok=True
        )

        # Save uploaded Word file
        with open(input_path, "wb") as file:
            file.write(word_bytes)

        # LibreOffice path
        libreoffice_path = (
            r"C:\Program Files\LibreOffice\program\soffice.exe"
        )

        # Run LibreOffice conversion
        result = subprocess.run(
            [
                libreoffice_path,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                output_dir,
                input_path
            ],
            capture_output=True,
            text=True,
            timeout=120
        )

        # Expected PDF path
        pdf_path = os.path.join(
            output_dir,
            "input.pdf"
        )

        if not os.path.exists(pdf_path):

            raise RuntimeError(
                "Unable to convert Word document to PDF. "
                + result.stderr
            )

        # Read converted PDF
        with open(pdf_path, "rb") as pdf_file:
            pdf_bytes = pdf_file.read()

        if not pdf_bytes:
            raise RuntimeError(
                "The converted PDF is empty."
            )

        return pdf_bytes