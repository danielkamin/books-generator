from pathlib import Path
import sys


def exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def external_data_dir(cli_dir: str | None = None) -> Path:
    return Path(cli_dir) if cli_dir else exe_dir() / "data"


def find_data_file(name: str, data_dir: str | None = None) -> Path:
    p = external_data_dir(data_dir) / name
    if p.exists():
        return p
    raise FileNotFoundError(f"Data file not found: {p}")


def resolve_dir(arg_value: str | None, default_name: str) -> Path:
    base = exe_dir()
    if arg_value:
        p = Path(arg_value)
        return p if p.is_absolute() else (base / p)
    return base / default_name
