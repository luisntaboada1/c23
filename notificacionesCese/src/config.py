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

ACTA_INPUT_FIELDS = [
    "numero de acta",
    "fecha de acta",
    "fecha de diligencia",
    "hora inicio diligencia",
    "dirección del restaurante",
    "nombre / denominación del restaurante",
    "fecha de entrega",
    "sexo de quien recibió la carta",
    "nombre de quien recibió la carta",
    "clausula_firma",
    "hora fin diligencia",
    "numero de instrumento",
    "fecha del instrumento",
]

ACTA_TEMPLATE_FIELDS = [
    "acta_numero",
    "acta_numero_letra",
    "fecha_acta_num",
    "fecha_acta_letra",
    "fecha_diligencia_letra",
    "hora_inicio_letra",
    "direccion_restaurante",
    "nombre_restaurante",
    "fecha_entrega_letra",
    "sexo_recibe",
    "nombre_recibe",
    "clausula_firma",
    "hora_fin_letra",
    "instrumento_numero_letra",
    "instrumento_fecha_letra",
]

ACTA_PLACEHOLDER_MAP = {
    "{{ acta_numero }}": "acta_numero",
    "{{ acta_numero_letra }}": "acta_numero_letra",
    "{{ fecha_acta_num }}": "fecha_acta_num",
    "{{ fecha_acta_letra }}": "fecha_acta_letra",
    "{{ fecha_diligencia_letra }}": "fecha_diligencia_letra",
    "{{ hora_inicio_letra }}": "hora_inicio_letra",
    "{{ direccion_restaurante }}": "direccion_restaurante",
    "{{ nombre_restaurante }}": "nombre_restaurante",
    "{{ fecha_entrega_letra }}": "fecha_entrega_letra",
    "{{ sexo_recibe }}": "sexo_recibe",
    "{{ nombre_recibe }}": "nombre_recibe",
    "{{ clausula_firma }}": "clausula_firma",
    "{{ hora_fin_letra }}": "hora_fin_letra",
    "{{ instrumento_numero_letra }}": "instrumento_numero_letra",
    "{{ instrumento_fecha_letra }}": "instrumento_fecha_letra",
}

ACTA_QR_MARKER_TEXT = "Aquí se insertará automáticamente el QR de RelacionAnexosQR_QR.png"
ACTA_QR_IMAGE_WIDTH_INCHES = 2.0
