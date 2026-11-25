from pathlib import Path
from typing import Any

import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = PROJECT_ROOT / "config.yaml"


def load_project_config() -> dict[str, Any]:
    with CONFIG_FILE.open(encoding="utf-8") as config_file:
        config = yaml.safe_load(config_file)

    if not isinstance(config, dict):
        raise ValueError(f"{CONFIG_FILE} must contain a YAML mapping")

    return config


def get_configured_path(name: str) -> Path:
    config = load_project_config()
    files = config.get("files")
    if not isinstance(files, dict):
        raise ValueError(f"{CONFIG_FILE} must contain a top-level 'files' mapping")

    value = files.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(
            f"{CONFIG_FILE} must define a non-empty path for files.{name}"
        )

    path = Path(value.strip()).expanduser()
    return path if path.is_absolute() else PROJECT_ROOT / path
