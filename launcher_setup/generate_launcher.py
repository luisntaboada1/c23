from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess


LAUNCHER_FILENAME = "run_derechos_transmision.cmd"
POWERSHELL_LAUNCHER_FILENAME = "run_derechos_transmision.ps1"
SHORTCUT_FILENAME = "DerechosTransmision.lnk"
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


def _detect_python_launcher(project_dir: Path) -> tuple[Path, str]:
    candidates = [
        project_dir / ".venv" / "Scripts" / "python.exe",
        project_dir / "venv" / "Scripts" / "python.exe",
        project_dir / "env" / "Scripts" / "python.exe",
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate, f"Virtual environment interpreter detected at {candidate}"

    for launcher_name, note in [
        ("python", "No local virtual environment was detected; using the system python executable"),
        ("py", "No local virtual environment was detected; using the Windows py launcher executable"),
    ]:
        command_path = _resolve_command_path(launcher_name)
        if command_path:
            return command_path, note

    raise FileNotFoundError(
        "No usable Python launcher was found. Install Python or create a local virtual environment first."
    )


def _resolve_command_path(command_name: str) -> Path | None:
    command_path = shutil.which(command_name)
    if not command_path:
        return None

    try:
        completed = subprocess.run(
            [command_path, "--version"],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return None

    if completed.returncode != 0:
        return None

    return Path(command_path).resolve()


def _command_is_available(command_name: str) -> bool:
    return _resolve_command_path(command_name) is not None


def _build_launcher_contents(project_dir: Path, python_executable: Path) -> str:
    return "\n".join(
        [
            "@echo off",
            f'set "PROJECT_DIR={project_dir}"',
            "pushd \"%PROJECT_DIR%\"",
            f'"{python_executable}" src\\main.py',
            "set \"EXIT_CODE=%ERRORLEVEL%\"",
            "popd",
            "exit /b %EXIT_CODE%",
            "",
        ]
    )


def _build_powershell_launcher_contents(project_dir: Path, python_executable: Path) -> str:
    escaped_project_dir = str(project_dir).replace("'", "''")
    escaped_python_executable = str(python_executable).replace("'", "''")
    return "\n".join(
        [
            "$ErrorActionPreference = 'Stop'",
            f"$projectDir = '{escaped_project_dir}'",
            f"$pythonExe = '{escaped_python_executable}'",
            "Push-Location $projectDir",
            "try {",
            "    & $pythonExe 'src\\main.py'",
            "    exit $LASTEXITCODE",
            "}",
            "finally {",
            "    Pop-Location",
            "}",
            "",
        ]
    )


def _create_shortcut(
    shortcut_path: Path, project_dir: Path, python_executable: Path
) -> tuple[bool, str]:
    escaped_shortcut_path = str(shortcut_path).replace("'", "''")
    escaped_project_dir = str(project_dir).replace("'", "''")
    escaped_python_executable = str(python_executable).replace("'", "''")
    powershell_script = "\n".join(
        [
            "$WshShell = New-Object -ComObject WScript.Shell",
            f"$Shortcut = $WshShell.CreateShortcut('{escaped_shortcut_path}')",
            f"$Shortcut.TargetPath = '{escaped_python_executable}'",
            "$Shortcut.Arguments = 'src\\main.py'",
            f"$Shortcut.WorkingDirectory = '{escaped_project_dir}'",
            f"$Shortcut.IconLocation = '{escaped_python_executable},0'",
            "$Shortcut.Save()",
            "",
        ]
    )

    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            powershell_script,
        ],
        capture_output=True,
        text=True,
        check=False,
    )

    if completed.returncode == 0 and shortcut_path.exists():
        return True, f"Shortcut created at {shortcut_path}"

    details = completed.stderr.strip() or completed.stdout.strip() or "unknown error"
    return False, f"Shortcut creation failed: {details}"


def _build_command_text(
    launcher_path: Path,
    powershell_launcher_path: Path,
    shortcut_path: Path,
    project_dir: Path,
    python_executable: Path,
    detection_note: str,
    shortcut_note: str,
) -> str:
    powershell_command = f'& "{launcher_path}"'
    cmd_command = f'"{launcher_path}"'
    python_command = f'& "{python_executable}" "{project_dir / "src" / "main.py"}"'
    ps1_command = f'powershell -ExecutionPolicy Bypass -File "{powershell_launcher_path}"'

    return "\n".join(
        [
            "derechosTransmision launcher command",
            "",
            f"Project folder: {project_dir}",
            f"Python executable: {python_executable}",
            f"Launcher file: {launcher_path}",
            f"PowerShell launcher file: {powershell_launcher_path}",
            f"Shortcut file: {shortcut_path}",
            detection_note,
            shortcut_note,
            "",
            "Recommended for the client:",
            "Use the .lnk shortcut instead of a shortcut to the .cmd file.",
            "",
            "Paste this into PowerShell to run Python directly:",
            python_command,
            "",
            "Optional PowerShell launcher:",
            ps1_command,
            "",
            "Legacy CMD launcher:",
            powershell_command,
            cmd_command,
            "",
            "Tip:",
            "If Norton flags the .cmd file, do not use a .cmd shortcut. Use the .lnk shortcut or run python.exe directly.",
            "",
        ]
    )


def generate_launcher(project_dir: Path, output_dir: Path | None = None) -> tuple[Path, Path, Path, Path]:
    _validate_project_dir(project_dir)

    output_dir = (output_dir or project_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    python_executable, detection_note = _detect_python_launcher(project_dir)
    launcher_path = output_dir / LAUNCHER_FILENAME
    powershell_launcher_path = output_dir / POWERSHELL_LAUNCHER_FILENAME
    shortcut_path = output_dir / SHORTCUT_FILENAME
    command_path = output_dir / COMMAND_FILENAME

    launcher_path.write_text(
        _build_launcher_contents(project_dir, python_executable),
        encoding="utf-8",
        newline="\r\n",
    )
    powershell_launcher_path.write_text(
        _build_powershell_launcher_contents(project_dir, python_executable),
        encoding="utf-8",
        newline="\r\n",
    )
    shortcut_created, shortcut_note = _create_shortcut(
        shortcut_path, project_dir, python_executable
    )
    if not shortcut_created and shortcut_path.exists():
        shortcut_path.unlink()
    command_path.write_text(
        _build_command_text(
            launcher_path,
            powershell_launcher_path,
            shortcut_path,
            project_dir,
            python_executable,
            detection_note,
            shortcut_note,
        ),
        encoding="utf-8",
        newline="\r\n",
    )

    return launcher_path, powershell_launcher_path, shortcut_path, command_path


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
        launcher_path, powershell_launcher_path, shortcut_path, command_path = generate_launcher(
            project_dir, output_dir
        )
    except Exception as error:
        print(f"Setup failed: {error}")
        return 1

    print("Launcher created successfully.")
    print(f"Launcher file: {launcher_path}")
    print(f"PowerShell launcher file: {powershell_launcher_path}")
    print(f"Shortcut file: {shortcut_path}")
    print(f"Command text: {command_path}")
    print("Open the text file and use the .lnk shortcut or the direct python command for the client.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
