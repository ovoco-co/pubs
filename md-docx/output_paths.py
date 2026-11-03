"""Shared handling of the required output argument for the md-docx scripts.

Every script takes the output as its second argument, and it is always required.
Scripts that write a folder of results take a folder. Scripts that write one
combined file take either a file path or a folder; given a folder, the file is
named after the input.
"""
from pathlib import Path


def resolve_output_folder(output_arg: str) -> Path:
    """Return the output folder, creating it if needed."""
    folder = Path(output_arg).resolve()
    if folder.exists() and not folder.is_dir():
        raise SystemExit(f"Output must be a folder, not a file: {folder}")
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def resolve_output_file(output_arg: str, default_name: str, suffixes: tuple[str, ...]) -> Path:
    """Return the output file path for a file-or-folder argument, creating its folder if needed.

    A path that is an existing folder, or has no file extension, is treated as a folder and
    gets `default_name`. Otherwise the extension must be one of `suffixes`.
    """
    path = Path(output_arg).resolve()
    if path.is_dir() or not path.suffix:
        path = path / default_name
    elif path.suffix.lower() not in suffixes:
        raise SystemExit(f"Output file must end in {' or '.join(suffixes)}, or be a folder: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path
