"""Helper for orchestrating Microsoft Access bulk CSV imports."""
from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Iterable, Optional


def _as_path(value: Optional[str]) -> Optional[Path]:
    if value is None:
        return None
    return Path(value)


def import_csv_to_access(
    *,
    msaccess_path: str,
    database_path: str,
    csv_path: Path,
    macro: Optional[str] = None,
    timeout: int = 600,
    use_cmd_argument: bool = True,
    extra_args: Optional[Iterable[str]] = None,
    workdir: Optional[str] = None,
    log_path: Optional[Path] = None,
) -> subprocess.CompletedProcess:
    """Invoke msaccess.exe to run a macro that performs a TransferText import."""
    msaccess = _as_path(msaccess_path)
    database = _as_path(database_path)
    csv_path = Path(csv_path)

    if msaccess is None or not msaccess.exists():
        raise FileNotFoundError(f"msaccess.exe not found: {msaccess_path}")
    if database is None or not database.exists():
        raise FileNotFoundError(f"Access database not found: {database_path}")
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    cmd = [str(msaccess), str(database)]
    if macro:
        cmd.extend(["/x", macro])
    if use_cmd_argument:
        cmd.extend(["/cmd", str(csv_path)])
    if extra_args:
        cmd.extend(list(extra_args))

    completed = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
        cwd=workdir,
        check=False,
    )

    if log_path:
        Path(log_path).parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as log:
            log.write("Command: " + " ".join(cmd) + "\n")
            log.write(f"Return code: {completed.returncode}\n")
            if completed.stdout:
                log.write("--- stdout ---\n")
                log.write(completed.stdout + "\n")
            if completed.stderr:
                log.write("--- stderr ---\n")
                log.write(completed.stderr + "\n")

    if completed.returncode != 0:
        raise RuntimeError(
            "Access bulk import failed with exit code "
            f"{completed.returncode}: {completed.stderr.strip()}"
        )

    return completed
