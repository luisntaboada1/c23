from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess


LAUNCHER_FILENAME = "run_derechos_transmision.cmd"
COMMAND_FILENAME = "RUN_DERECHOS_TRANSMISION_COMMAND.txt"


def _resolve_project_dir(raw_path: str | None) -> Path:
    if raw_path:
        return Path(raw_path).expanduser().resolve()

    entered = input(
        "Paste the full path to the derechosTransmision folder (or to src\\main.py): "
    ).strip()
    if not entered:
        raise ValueError("No path was provided.")

    candidate = Path(entered).expanduser().resolve()
    if candidate.is_file() and candidate.name.lower() == "main.py":
        return candidate.parent.parent
    return candidate


def _validate_project_dir(project_dir: Path) -> None:
    main_path = project_dir / "src" / "main.py"
    if not main_path.exists():
        raise FileNotFoundError(
            f"Could not find src\\main.py inside: {project_dir}"
        )


def _detect_python_launcher(project_dir: Path) -> tuple[str, str]:
    candidates = [
        project_dir / ".venv" / "Scripts" / "python.exe",
        project_dir / "venv" / "Scripts" / "python.exe",
        project_dir / "env" / "Scripts" / "python.exe",
    ]

    for candidate in candidates:
        if candidate.exists():
            command = f'"{candidate}" src\\main.py'
            return command, f"Virtual environment interpreter detected at {candidate}"

    for launcher_name, note in [
        ("python", "No local virtual environment was detected; using the system python command"),
        ("py", "No local virtual environment was detected; using the Windows py launcher"),
    ]:
        if _command_is_available(launcher_name):
            return f"{launcher_name} src\\main.py", note

    raise FileNotFoundError(
        "No usable Python launcher was found. Install Python or create a local virtual environment first."
    )


def _command_is_available(command_name: str) -> bool:
    command_path = shutil.which(command_name)
    if not command_path:
        return False

    try:
        completed = subprocess.run(
            [command_name, "--version"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return False

    return completed.returncode == 0


def _build_launcher_contents(project_dir: Path, python_command: str) -> str:
    return "\n".join(
        [
            "@echo off",
            f'set "PROJECT_DIR={project_dir}"',
            "pushd \"%PROJECT_DIR%\"",
            python_command,
            "set \"EXIT_CODE=%ERRORLEVEL%\"",
            "popd",
            "exit /b %EXIT_CODE%",
            "",
        ]
    )


def _build_command_text(launcher_path: Path, project_dir: Path, detection_note: str) -> str:
    powershell_command = f'& "{launcher_path}"'
    cmd_command = f'"{launcher_path}"'

    return "\n".join(
        [
            "derechosTransmision launcher command",
            "",
            f"Project folder: {project_dir}",
            f"Launcher file: {launcher_path}",
            detection_note,
            "",
            "Paste this into PowerShell:",
            powershell_command,
            "",
            "Paste this into Command Prompt:",
            cmd_command,
            "",
            "Tip:",
            "You can also double-click the .cmd file to launch the process.",
            "",
        ]
    )


def generate_launcher(project_dir: Path, output_dir: Path | None = None) -> tuple[Path, Path]:
    _validate_project_dir(project_dir)

    output_dir = (output_dir or project_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    python_command, detection_note = _detect_python_launcher(project_dir)
    launcher_path = output_dir / LAUNCHER_FILENAME
    command_path = output_dir / COMMAND_FILENAME

    launcher_path.write_text(
        _build_launcher_contents(project_dir, python_command),
        encoding="utf-8",
        newline="\r\n",
    )
    command_path.write_text(
        _build_command_text(launcher_path, project_dir, detection_note),
        encoding="utf-8",
        newline="\r\n",
    )

    return launcher_path, command_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Generate a Windows launcher file and a paste-ready command text "
            "for derechosTransmision."
        )
    )
    parser.add_argument(
        "project_dir",
        nargs="?",
        help="Path to the derechosTransmision folder or directly to src\\main.py",
    )
    parser.add_argument(
        "--output-dir",
        help="Directory where the launcher and command text should be created",
    )
    args = parser.parse_args()

    try:
        project_dir = _resolve_project_dir(args.project_dir)
        output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else None
        launcher_path, command_path = generate_launcher(project_dir, output_dir)
    except Exception as error:
        print(f"Setup failed: {error}")
        return 1

    print("Launcher created successfully.")
    print(f"Launcher file: {launcher_path}")
    print(f"Command text: {command_path}")
    print("Open the text file and copy the PowerShell command for the client to paste.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
