import logging
import sys
from datetime import datetime
from pathlib import Path
from types import TracebackType

from project_config import get_configured_path


_active_logger: logging.Logger | None = None
_original_excepthook = sys.excepthook


def _log_uncaught_exception(
    exc_type: type[BaseException],
    exc_value: BaseException,
    exc_traceback: TracebackType | None,
) -> None:
    if issubclass(exc_type, KeyboardInterrupt):
        _original_excepthook(exc_type, exc_value, exc_traceback)
        return

    if _active_logger is not None:
        _active_logger.critical(
            "Unhandled exception",
            exc_info=(exc_type, exc_value, exc_traceback),
        )
    _original_excepthook(exc_type, exc_value, exc_traceback)


def setup_logger(script_file: str) -> tuple[logging.Logger, Path]:
    global _active_logger

    script_path = Path(script_file).resolve()
    logs_dir = get_configured_path("logs_folder")
    logs_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    log_file = logs_dir / f"{script_path.stem}_{timestamp}.log"

    logger = logging.getLogger(log_file.stem)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = logging.Formatter(
        "%(asctime)s.%(msecs)03d | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    _active_logger = logger
    sys.excepthook = _log_uncaught_exception

    logger.info("Run started: %s", script_path.name)
    logger.info("Log file: %s", log_file)
    return logger, log_file
