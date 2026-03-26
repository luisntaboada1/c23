import traceback

from googleapiclient.errors import HttpError

from drive_service import get_drive_service, list_child_folders
from pipeline import run_pipeline


def _print_run_summary(result: dict):
    print(f"Restaurant folder: {result['folder_name']}")

    if result.get("downloaded_files"):
        print("\nDownloaded input files:")
        for item in result["downloaded_files"]:
            print(f" - {item['label']}")

    if result.get("generated_files"):
        print("\nGenerated temporary files:")
        for item in result["generated_files"]:
            print(f" - {item['label']}")

    if result.get("uploaded_files"):
        print("\nUploaded files:")
        for item in result["uploaded_files"]:
            print(f" - {item['label']}: {item['link']}")

    cleanup = result.get("cleanup", {})
    if cleanup.get("attempted"):
        status = "deleted" if cleanup.get("deleted") else "not deleted"
        print(f"\nTemporary run folder {status}: {cleanup.get('run_folder')}")
        if cleanup.get("error"):
            print(f"Cleanup warning: {cleanup['error']}")


def _normalize_mode(mode: str) -> str | None:
    normalized = mode.strip().lower()
    if normalized in {"1", "s", "single"}:
        return "single"
    if normalized in {"2", "m", "multi", "multiple", "batch"}:
        return "batch"
    return None


def _run_single_folder(folder_link: str, service=None) -> bool:
    result = run_pipeline(folder_link, service=service)

    if result is None:
        print("Validation failed.")
        return False

    print("Process completed successfully.")
    _print_run_summary(result)
    return True


def _print_batch_summary(successes: list[dict], failures: list[dict]):
    print("\nBatch run completed.")
    print(f"Successful folders: {len(successes)}")
    print(f"Failed folders: {len(failures)}")

    if successes:
        print("\nSuccessful runs:")
        for item in successes:
            print(f" - {item['name']}")

    if failures:
        print("\nFailed runs:")
        for item in failures:
            reason = item.get("reason")
            if reason:
                print(f" - {item['name']}: {reason}")
            else:
                print(f" - {item['name']}")


def _run_multiple_folders(parent_folder_link: str) -> None:
    service = get_drive_service()
    child_folders = list_child_folders(service, parent_folder_link)

    if not child_folders:
        print("No child folders were found inside the provided Google Drive folder.")
        return

    print(f"Found {len(child_folders)} child folder(s) to process.")

    successes = []
    failures = []

    for index, child_folder in enumerate(child_folders, start=1):
        print(f"\n[{index}/{len(child_folders)}] Processing: {child_folder['name']}")

        try:
            completed = _run_single_folder(child_folder["link"], service=service)
            if completed:
                successes.append({"name": child_folder["name"], "id": child_folder["id"]})
            else:
                failures.append(
                    {
                        "name": child_folder["name"],
                        "id": child_folder["id"],
                        "reason": "Validation failed",
                    }
                )
        except HttpError as error:
            print(f"An error occurred: {error}")
            run_summary = getattr(error, "run_summary", None)
            if run_summary:
                _print_run_summary(run_summary)
            traceback.print_exc()
            failures.append(
                {
                    "name": child_folder["name"],
                    "id": child_folder["id"],
                    "reason": f"{type(error).__name__}: {error}",
                }
            )
        except Exception as error:
            print(f"Unexpected error: {error}")
            run_summary = getattr(error, "run_summary", None)
            if run_summary:
                _print_run_summary(run_summary)
            traceback.print_exc()
            failures.append(
                {
                    "name": child_folder["name"],
                    "id": child_folder["id"],
                    "reason": f"{type(error).__name__}: {error}",
                }
            )

    _print_batch_summary(successes, failures)


def main():
    mode = _normalize_mode(
        input(
            "Choose mode: [1] single folder, [2] parent folder with many child folders: "
        )
    )

    if mode is None:
        print("Invalid mode. Use 1 for a single folder or 2 for a parent folder.")
        return

    try:
        if mode == "single":
            folder_link = input("Paste Google Drive folder link: ").strip()
            _run_single_folder(folder_link)
            return

        parent_folder_link = input(
            "Paste Google Drive parent folder link containing the folders to process: "
        ).strip()
        _run_multiple_folders(parent_folder_link)

    except HttpError as error:
        print(f"An error occurred: {error}")
        run_summary = getattr(error, "run_summary", None)
        if run_summary:
            _print_run_summary(run_summary)
        traceback.print_exc()
    except Exception as error:
        print(f"Unexpected error: {error}")
        run_summary = getattr(error, "run_summary", None)
        if run_summary:
            _print_run_summary(run_summary)
        traceback.print_exc()


if __name__ == "__main__":
    main()
