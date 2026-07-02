"""Load and validate config.toml.
Settings are read with the standard-library tomllib (no extra dependency). There are
no secrets in the config: the Outlook email backend uses your signed-in desktop
Outlook, so nothing here is sensitive.
"""

from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path


class ConfigError(Exception):
    """Raised for a missing or invalid setting."""


@dataclass(frozen=True)
class EmailConfig:
    enabled: bool
    recipients: tuple[str, ...]
    send_on: str            # "always" | "failure"


@dataclass(frozen=True)
class Config:
    reports_dir: Path
    log_dir: Path
    include: tuple[str, ...]
    exclude: tuple[str, ...]
    timeout_seconds: int
    email: EmailConfig


def load_config(path: Path) -> Config:
    if not path.is_file():
        raise ConfigError(f"Config file not found: {path}")
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{path} is not valid TOML: {exc}") from exc

    paths = data.get("paths", {})
    if "reports_dir" not in paths:
        raise ConfigError("Missing 'reports_dir' under [paths] in config.toml.")
    reports_dir = Path(paths["reports_dir"])
    log_dir = Path(paths.get("log_dir", reports_dir.parent / "excel-refresh-logs"))

    discovery = data.get("discovery", {})
    include = tuple(discovery.get("include", ["*.xlsx", "*.xlsm"]))
    exclude = tuple(discovery.get("exclude", []))

    refresh = data.get("refresh", {})
    timeout = int(refresh.get("timeout_seconds", 1800))
    if timeout <= 0:
        raise ConfigError("[refresh].timeout_seconds must be a positive number of seconds.")

    email_table = data.get("email", {})
    email = EmailConfig(
        enabled=bool(email_table.get("enabled", False)),
        recipients=tuple(email_table.get("recipients", [])),
        send_on=email_table.get("send_on", "always"),
    )
    if email.enabled and not email.recipients:
        raise ConfigError("[email].recipients must be non-empty when email is enabled.")

    return Config(
        reports_dir=reports_dir,
        log_dir=log_dir,
        include=include,
        exclude=exclude,
        timeout_seconds=timeout,
        email=email,
    )
