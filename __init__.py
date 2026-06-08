"""Altomizer standalone ALT-text management package."""

from __future__ import annotations

import os
from pathlib import Path


def _strip_env_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def _load_local_dotenv() -> None:
    package_dir = Path(__file__).resolve().parent
    candidates = (
        package_dir / ".env",
        package_dir.parent / ".env",
    )
    for env_path in candidates:
        if not env_path.exists() or not env_path.is_file():
            continue
        try:
            for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[7:].strip()
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if not key or key in os.environ:
                    continue
                os.environ[key] = _strip_env_quotes(value)
            break
        except OSError:
            continue


_load_local_dotenv()
