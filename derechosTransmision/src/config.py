from pathlib import Path

SCOPES = ["https://www.googleapis.com/auth/drive"]

BASE_DIR = Path(__file__).resolve().parent
TOKEN_PATH = BASE_DIR / "token.json"
CREDS_PATH = BASE_DIR / "credentials.json"

REQUIRED_FILES = {
    "ANEXOA.pdf": "anexoA",
    "ANEXOB.pdf": "anexoB",
    "ANEXOC.pdf": "anexoC",
    "acta_data.xlsx": "acta_data",
}

PDFS_TO_VALIDATE_AS_PUBLIC = {"ANEXOA.pdf", "ANEXOB.pdf", "ANEXOC.pdf"}

RUNS_PATH = BASE_DIR.parent / "runs"

TEMPLATES_PATH = BASE_DIR.parent / "templates"
RELACION_ANEXOS_TEMPLATE_PATH = TEMPLATES_PATH / "relacion_anexos_qr_template.docx"
ACTA_TEMPLATE_PATH = TEMPLATES_PATH / "plantilla_acta_1t.docx"

ANEXO_QR_FILENAMES = {
    "anexoA": "ANEXOA_QR.png",
    "anexoB": "ANEXOB_QR.png",
    "anexoC": "ANEXOC_QR.png",
}

RELACION_ANEXOS_OUTPUT_FILENAME = "RelacionAnexosQR.pdf"
RELACION_ANEXOS_OUTPUT_DOCX_FILENAME = "RelacionAnexosQR.docx"
RELACION_ANEXOS_QR_IMAGE_WIDTH_INCHES = 1.75
RELACION_ANEXOS_QR_OUTPUT_FILENAME = "RelacionAnexosQR_QR.png"

ACTA_DATA_LOCAL_FILENAME = "acta_data.xlsx"

ACTA_1T_TEMPLATE_PATH = TEMPLATES_PATH / "plantilla_acta_1t.docx"
ACTA_A_TEMPLATE_PATH = TEMPLATES_PATH / "plantilla_acta_a.docx"

ACTA_1T_OUTPUT_DOCX_FILENAME = "ACTA_1T.docx"
ACTA_1T_OUTPUT_PDF_FILENAME = "ACTA_1T.pdf"

ACTA_A_OUTPUT_DOCX_FILENAME = "ACTA_A.docx"
ACTA_A_OUTPUT_PDF_FILENAME = "ACTA_A.pdf"

ACTA_REQUIRED_FIELDS = [
    "numero de acta",
    "fecha escrita",
    "hora inicio diligencia",
    "dirección y referencias",
    "nombre del restaurante",
    "numero de mesas",
    "numero de pantallas",
    "partido",
    "hora fin diligencia",
]

ACTA_PLACEHOLDER_MAP = {
    "{{ numero de acta }}": "numero de acta",
    "{{ fecha escrita }}": "fecha escrita",
    "{{ hora inicio diligencia }}": "hora inicio diligencia",
    "{{ dirección y referencias }}": "dirección y referencias",
    "{{ nombre del restaurante }}": "nombre del restaurante",
    "{{ numero de mesas }}": "numero de mesas",
    "{{ numero de pantallas }}": "numero de pantallas",
    "{{ partido }}": "partido",
    "{{ hora fin diligencia }}": "hora fin diligencia",
}

ACTA_QR_MARKER_TEXT = "Aquí se insertará automáticamente el QR de RelacionAnexosQR_QR.png"
ACTA_QR_IMAGE_WIDTH_INCHES = 2.0