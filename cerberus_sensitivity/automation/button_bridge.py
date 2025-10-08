from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Iterable, Sequence


def _default_project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _default_exe_path() -> Path:
    return _default_project_root() / "bin" / "Debug" / "net9.0-windows" / "Test_C.exe"


def _prepare_dump_path(dump_path: str | Path | None) -> Path | None:
    if dump_path is None:
        return None
    path_obj = Path(dump_path)
    return path_obj if path_obj.is_absolute() else (Path.cwd() / path_obj).resolve()


def invoke_button(
    button_key: str,
    *,
    dump_path: str | Path | None = None,
    window_regex: str | None = None,
    exe_path: str | Path | None = None,
    timeout: float = 180.0,
    capture_output: bool = True,
) -> subprocess.CompletedProcess[str]:
    """Invoke a repository button using the compiled C# helper."""
    exe = Path(exe_path) if exe_path is not None else _default_exe_path()
    if not exe.exists():
        raise FileNotFoundError(f"Button automation helper not found at: {exe}")

    resolved_dump = _prepare_dump_path(dump_path)

    args: list[str] = [str(exe)]
    if resolved_dump is not None:
        args.extend(["--dump", str(resolved_dump)])
    if window_regex:
        args.extend(["--window-regex", window_regex])
    args.extend(["invoke", button_key])

    return subprocess.run(
        args,
        check=True,
        timeout=timeout,
        text=True,
        capture_output=capture_output,
    )


def list_buttons(
    *,
    dump_path: str | Path | None = None,
    exe_path: str | Path | None = None,
    window_regex: str | None = None,
    timeout: float = 30.0,
) -> Sequence[str]:
    """Return the button keys that are currently defined in the inspect dump."""
    exe = Path(exe_path) if exe_path is not None else _default_exe_path()
    if not exe.exists():
        raise FileNotFoundError(f"Button automation helper not found at: {exe}")

    resolved_dump = _prepare_dump_path(dump_path)
    args: list[str] = [str(exe)]
    if resolved_dump is not None:
        args.extend(["--dump", str(resolved_dump)])
    if window_regex:
        args.extend(["--window-regex", window_regex])
    args.append("list")

    completed = subprocess.run(
        args,
        check=True,
        timeout=timeout,
        text=True,
        capture_output=True,
    )

    keys: list[str] = []
    for line in completed.stdout.splitlines():
        if line.strip().startswith("-"):
            parts = line.split(":", 1)
            if parts:
                key_part = parts[0]
                key = key_part.replace("-", "").strip()
                if key:
                    keys.append(key)
    return keys
