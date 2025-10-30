from __future__ import annotations

import json
import subprocess
from pathlib import Path
from subprocess import CalledProcessError, TimeoutExpired
import time
from typing import Mapping, Sequence, NoReturn


def _default_project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _default_exe_path() -> Path:
    return _default_project_root() / "bin" / "Debug" / "net10.0-windows" / "Drill_Down_With_C.exe"


def _prepare_dump_path(dump_path: str | Path | None) -> Path | None:
    if dump_path is None:
        return None
    path_obj = Path(dump_path)
    return path_obj if path_obj.is_absolute() else (Path.cwd() / path_obj).resolve()


def _run_automation_command(
    args: list[str],
    *,
    action: str,
    button_key: str,
    timeout: float,
    capture_output: bool,
    max_attempts: int = 2,
) -> subprocess.CompletedProcess[str]:
    attempt = 0
    while True:
        try:
            return subprocess.run(
                args,
                check=True,
                timeout=timeout,
                text=True,
                capture_output=capture_output,
            )
        except (CalledProcessError, TimeoutExpired) as exc:
            attempt += 1
            if attempt >= max_attempts:
                _raise_subprocess_error(action, button_key, exc)
            else:
                wait = min(2.0, 0.5 * attempt)
                print(f"[automation] {action} '{button_key}' failed (attempt {attempt}); retrying in {wait}s...")
                time.sleep(wait)


def _raise_subprocess_error(action: str, button_key: str, exc: Exception) -> NoReturn:
    if isinstance(exc, TimeoutExpired):
        timeout = exc.timeout if exc.timeout is not None else "unknown"
        stdout = getattr(exc, "output", "") or getattr(exc, "stdout", "") or ""
        stderr = getattr(exc, "stderr", "") or ""
        base_message = f"Timed out trying to {action} '{button_key}' after {timeout} seconds."
    elif isinstance(exc, CalledProcessError):
        stdout = exc.stdout or ""
        stderr = exc.stderr or ""
        base_message = (
            f"Failed to {action} '{button_key}' (exit code {exc.returncode})."
        )
    else:
        stdout = stderr = ""
        base_message = f"Error while attempting to {action} '{button_key}': {exc}"

    def _format(value: str) -> str:
        return value.strip() if isinstance(value, str) else ""

    stdout_text = _format(stdout)
    stderr_text = _format(stderr)

    details_parts: list[str] = []
    if stdout_text:
        details_parts.append(f"stdout:\n{stdout_text}")
    if stderr_text:
        details_parts.append(f"stderr:\n{stderr_text}")

    cmd = getattr(exc, "cmd", None)
    if cmd:
        if isinstance(cmd, (list, tuple)):
            cmd_text = subprocess.list2cmdline([str(part) for part in cmd])
        else:
            cmd_text = str(cmd)
        details_parts.insert(0, f"Command: {cmd_text}")

    details = "\n\n".join(details_parts) if details_parts else "No additional output captured."
    raise RuntimeError(f"{base_message}\n{details}") from exc


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

    return _run_automation_command(
        args,
        action="invoke",
        button_key=button_key,
        timeout=timeout,
        capture_output=capture_output,
    )


def set_button_value(
    button_key: str,
    value: str,
    *,
    dump_path: str | Path | None = None,
    window_regex: str | None = None,
    exe_path: str | Path | None = None,
    timeout: float = 180.0,
    capture_output: bool = True,
) -> subprocess.CompletedProcess[str]:
    """Set a ValuePattern-aware control to the specified value using the C# helper."""
    exe = Path(exe_path) if exe_path is not None else _default_exe_path()
    if not exe.exists():
        raise FileNotFoundError(f"Button automation helper not found at: {exe}")

    resolved_dump = _prepare_dump_path(dump_path)

    args: list[str] = [str(exe)]
    if resolved_dump is not None:
        args.extend(["--dump", str(resolved_dump)])
    if window_regex:
        args.extend(["--window-regex", window_regex])
    args.extend(["set", button_key, str(value)])

    return _run_automation_command(
        args,
        action="set value on",
        button_key=button_key,
        timeout=timeout,
        capture_output=capture_output,
    )


def collect_table(
    button_key: str,
    *,
    dump_path: str | Path | None = None,
    window_regex: str | None = None,
    exe_path: str | Path | None = None,
    timeout: float = 180.0,
) -> Mapping[str, list[list[str]]]:
    """Collect a grid/table control as structured data using the C# helper."""
    exe = Path(exe_path) if exe_path is not None else _default_exe_path()
    if not exe.exists():
        raise FileNotFoundError(f"Button automation helper not found at: {exe}")

    resolved_dump = _prepare_dump_path(dump_path)

    args: list[str] = [str(exe)]
    if resolved_dump is not None:
        args.extend(["--dump", str(resolved_dump)])
    if window_regex:
        args.extend(["--window-regex", window_regex])
    args.extend(["collect", button_key])

    completed = _run_automation_command(
        args,
        action="collect table for",
        button_key=button_key,
        timeout=timeout,
        capture_output=True,
    )

    payload = completed.stdout.strip().splitlines()[-1] if completed.stdout.strip() else ""
    if not payload:
        return {"Headers": [], "Rows": []}  # type: ignore[return-value]

    try:
        parsed = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Failed to parse table payload for '{button_key}': {payload}") from exc

    headers = parsed.get("Headers", [])
    rows = parsed.get("Rows", [])
    return {"Headers": headers, "Rows": rows}  # type: ignore[return-value]


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
