"""Tiny reader for non-secret config in .env (e.g. DOWNPAYMENT, UP_DATE). .env is
gitignored, so personal values stay out of the committed code."""
from datetime import date
from pathlib import Path

_ENV = Path(__file__).resolve().parent / ".env"


def load_env(path=_ENV):
    env = {}
    p = Path(path)
    if p.exists():
        for line in p.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def get_float(key, default):
    raw = load_env().get(key, "")
    try:
        return float(raw) if raw else default
    except ValueError:
        return default


def get_date(key, default):
    raw = load_env().get(key, "")
    try:
        return date.fromisoformat(raw) if raw else default
    except ValueError:
        return default
