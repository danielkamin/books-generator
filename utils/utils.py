from pathlib import Path


def get_unique_file_path(base_path: Path) -> Path:
    """
    Given a base path, return a unique file path by adding a numeric suffix if needed.
    """
    counter = 1
    unique_path = base_path

    while unique_path.exists():
        unique_path = base_path.with_name(
            f"{base_path.stem}_{counter}{base_path.suffix}"
        )
        counter += 1

    return unique_path
