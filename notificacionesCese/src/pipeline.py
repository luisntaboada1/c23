import os
import re
import shutil
import tempfile
import time
from pathlib import Path

from config import (
    ACTA_1T_OUTPUT_DOCX_FILENAME,
    ACTA_1T_OUTPUT_PDF_FILENAME,
    ACTA_1T_TEMPLATE_PATH,
    ACTA_A_OUTPUT_DOCX_FILENAME,
    ACTA_A_OUTPUT_PDF_FILENAME,
    ACTA_A_TEMPLATE_PATH,
    ACTA_DATA_LOCAL_FILENAME,
    RELACION_ANEXOS_OUTPUT_DOCX_FILENAME,
    RELACION_ANEXOS_OUTPUT_FILENAME,
    RELACION_ANEXOS_TEMPLATE_PATH,
    RUNS_PATH,
)
from drive_service import (
    download_file_from_drive,
    get_drive_service,
    upload_file_to_folder,
    validate_required_folder_files,
)
from pdf_service import (
    convert_docx_to_pdf,
    generate_acta_docx,
    generate_anexo_qrs,
    generate_relacion_anexos_docx,
    generate_relacion_anexos_qr,
    read_acta_data_xlsx,
)


def _sanitize_folder_name(folder_name: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*]', "_", folder_name).strip()
    sanitized = sanitized.rstrip(". ")
    return sanitized or "unnamed_folder"

def _create_run_folder(folder_name: str) -> Path:
    RUNS_PATH.mkdir(parents=True, exist_ok=True)
    prefix = f"{_sanitize_folder_name(folder_name)}-"
    return Path(tempfile.mkdtemp(prefix=prefix, dir=RUNS_PATH))


def _handle_remove_readonly(func, path, exc_info):
    error = exc_info[1]
    if isinstance(error, PermissionError):
        os.chmod(path, 0o700)
        func(path)
        return
    raise error


def _delete_run_folder(run_folder: Path, retries: int = 5, delay_seconds: float = 1.0) -> dict:
    cleanup_result = {
        "attempted": True,
        "deleted": False,
        "run_folder": str(run_folder),
        "error": None,
    }

    if not run_folder.exists():
        cleanup_result["deleted"] = True
        return cleanup_result

    last_error = None
    for attempt in range(1, retries + 1):
        try:
            shutil.rmtree(run_folder, onerror=_handle_remove_readonly)
            cleanup_result["deleted"] = not run_folder.exists()
            if cleanup_result["deleted"]:
                return cleanup_result
        except Exception as error:
            last_error = error

        if attempt < retries:
            time.sleep(delay_seconds)

    cleanup_result["error"] = (
        f"{type(last_error).__name__}: {last_error}" if last_error else "Unknown cleanup failure."
    )
    return cleanup_result


def _generate_optional_acta(
    *,
    template_path: Path,
    template_label: str,
    docx_filename: str,
    pdf_filename: str,
    result_key_prefix: str,
    result: dict,
    acta_data: dict,
    relacion_pdf_qr: str,
    folder_id: str,
    run_folder: Path,
    service,
):
    if not template_path.exists():
        result.setdefault("skipped_files", []).append(
            {
                "label": template_label,
                "reason": f"Template not found: {template_path}",
            }
        )
        return

    acta_docx = generate_acta_docx(
        acta_data=acta_data,
        relacion_anexos_qr_path=relacion_pdf_qr,
        output_folder=run_folder,
        template_docx_path=template_path,
        output_filename=docx_filename,
    )
    result[f"{result_key_prefix}_docx"] = acta_docx
    result["generated_files"].append({"label": f"{result_key_prefix}_docx", "path": acta_docx})

    uploaded_acta_docx = upload_file_to_folder(
        service=service,
        folder_id=folder_id,
        local_file_path=acta_docx,
    )
    result[f"uploaded_{result_key_prefix}_docx"] = uploaded_acta_docx
    result["uploaded_files"].append(
        {"label": f"{result_key_prefix}_docx", "link": uploaded_acta_docx["link"]}
    )

    acta_pdf = convert_docx_to_pdf(
        docx_path=acta_docx,
        output_folder=run_folder,
        output_filename=pdf_filename,
    )
    result[f"{result_key_prefix}_pdf"] = acta_pdf
    result["generated_files"].append({"label": f"{result_key_prefix}_pdf", "path": acta_pdf})

    uploaded_acta_pdf = upload_file_to_folder(
        service=service,
        folder_id=folder_id,
        local_file_path=acta_pdf,
    )
    result[f"uploaded_{result_key_prefix}_pdf"] = uploaded_acta_pdf
    result["uploaded_files"].append(
        {"label": f"{result_key_prefix}_pdf", "link": uploaded_acta_pdf["link"]}
    )


def run_pipeline(folder_link: str, service=None):
    service = service or get_drive_service()

    validation_result = validate_required_folder_files(service, folder_link)
    if validation_result is None:
        return None

    folder_id = validation_result["folder_id"]
    run_folder = _create_run_folder(validation_result["folder_name"])
    result = {
        "folder_id": validation_result["folder_id"],
        "folder_name": validation_result["folder_name"],
        "run_folder": str(run_folder),
        "links": validation_result["links"],
        "ids": validation_result["ids"],
        "generated_files": [],
        "downloaded_files": [],
        "uploaded_files": [],
        "skipped_files": [],
        "cleanup": {
            "attempted": False,
            "deleted": False,
            "run_folder": str(run_folder),
            "error": None,
        },
    }

    try:
        local_acta_data_path = download_file_from_drive(
            service=service,
            file_id=validation_result["ids"]["acta_data"],
            destination_path=run_folder / ACTA_DATA_LOCAL_FILENAME,
        )
        result["downloaded_files"].append(
            {
                "label": "acta_data.xlsx",
                "path": local_acta_data_path,
            }
        )

        acta_data = read_acta_data_xlsx(local_acta_data_path)
        result["acta_data"] = acta_data

        qr_files = generate_anexo_qrs(validation_result, run_folder)
        result["qr_files"] = qr_files
        for alias, path in qr_files.items():
            result["generated_files"].append({"label": alias, "path": path})

        relacion_docx = generate_relacion_anexos_docx(
            qr_files=qr_files,
            output_folder=run_folder,
            template_docx_path=RELACION_ANEXOS_TEMPLATE_PATH,
            output_filename=RELACION_ANEXOS_OUTPUT_DOCX_FILENAME,
        )
        result["relacion_docx"] = relacion_docx
        result["generated_files"].append({"label": "relacion_docx", "path": relacion_docx})

        relacion_pdf = convert_docx_to_pdf(
            docx_path=relacion_docx,
            output_folder=run_folder,
            output_filename=RELACION_ANEXOS_OUTPUT_FILENAME,
        )
        result["relacion_pdf"] = relacion_pdf
        result["generated_files"].append({"label": "relacion_pdf", "path": relacion_pdf})

        uploaded_qr_files = {}
        for alias, local_path in qr_files.items():
            uploaded_qr_files[alias] = upload_file_to_folder(
                service=service,
                folder_id=folder_id,
                local_file_path=local_path,
            )
            result["uploaded_files"].append(
                {"label": alias, "link": uploaded_qr_files[alias]["link"]}
            )
        result["uploaded_qr_files"] = uploaded_qr_files

        uploaded_relacion_pdf = upload_file_to_folder(
            service=service,
            folder_id=folder_id,
            local_file_path=relacion_pdf,
            make_public=True,
        )
        result["uploaded_relacion_pdf"] = uploaded_relacion_pdf
        result["uploaded_files"].append(
            {"label": "relacion_pdf", "link": uploaded_relacion_pdf["link"]}
        )

        relacion_pdf_qr = generate_relacion_anexos_qr(
            relacion_anexos_link=uploaded_relacion_pdf["link"],
            output_folder=run_folder,
        )
        result["relacion_pdf_qr"] = relacion_pdf_qr
        result["generated_files"].append({"label": "relacion_pdf_qr", "path": relacion_pdf_qr})

        uploaded_relacion_pdf_qr = upload_file_to_folder(
            service=service,
            folder_id=folder_id,
            local_file_path=relacion_pdf_qr,
        )
        result["uploaded_relacion_pdf_qr"] = uploaded_relacion_pdf_qr
        result["uploaded_files"].append(
            {"label": "relacion_pdf_qr", "link": uploaded_relacion_pdf_qr["link"]}
        )

        _generate_optional_acta(
            template_path=ACTA_1T_TEMPLATE_PATH,
            template_label="acta_1t_template",
            docx_filename=ACTA_1T_OUTPUT_DOCX_FILENAME,
            pdf_filename=ACTA_1T_OUTPUT_PDF_FILENAME,
            result_key_prefix="acta_1t",
            result=result,
            acta_data=acta_data,
            relacion_pdf_qr=relacion_pdf_qr,
            folder_id=folder_id,
            run_folder=run_folder,
            service=service,
        )

        _generate_optional_acta(
            template_path=ACTA_A_TEMPLATE_PATH,
            template_label="acta_a_template",
            docx_filename=ACTA_A_OUTPUT_DOCX_FILENAME,
            pdf_filename=ACTA_A_OUTPUT_PDF_FILENAME,
            result_key_prefix="acta_a",
            result=result,
            acta_data=acta_data,
            relacion_pdf_qr=relacion_pdf_qr,
            folder_id=folder_id,
            run_folder=run_folder,
            service=service,
        )

        return result
    except Exception as error:
        error.run_summary = result
        raise
    finally:
        result["cleanup"] = _delete_run_folder(run_folder)
