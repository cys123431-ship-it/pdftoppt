import json
import os
from pathlib import Path
from typing import Any

DEFAULT_SETTINGS: dict[str, Any] = {
    "language": "ko",
    "conflict_policy": "auto_rename",
    "render_dpi": 144,
    "jpg_quality": 90,
    "ocr_language": "kor+eng",
    "ocr_mode": "auto",
    "ocr_dpi": 300,
    "open_output_folder": True,
    "batch_target": "PPTX",
    "write_failure_log": True,
    "last_output_dir": "",
}


def get_settings_path() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        root = Path(base)
    else:
        root = Path.home() / ".config"
    return root / "PDFConverter" / "settings.json"


def load_settings(path: str | os.PathLike[str] | None = None) -> dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    target = Path(path) if path is not None else get_settings_path()
    try:
        with target.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
        if isinstance(raw, dict):
            for key in DEFAULT_SETTINGS:
                if key in raw:
                    settings[key] = raw[key]
    except (OSError, ValueError, TypeError):
        pass
    return settings


def save_settings(
    settings: dict[str, Any],
    path: str | os.PathLike[str] | None = None,
) -> None:
    target = Path(path) if path is not None else get_settings_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    safe_settings = {
        key: settings.get(key, default)
        for key, default in DEFAULT_SETTINGS.items()
    }
    temp = target.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as handle:
        json.dump(safe_settings, handle, ensure_ascii=False, indent=2)
    os.replace(temp, target)
