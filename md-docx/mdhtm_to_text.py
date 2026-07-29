import sys
import subprocess
import tempfile
from pathlib import Path

SUPPORTED_INPUT_EXTENSIONS = {".md"}


def read_input_file(file_path: Path) -> str:
    encodings_to_try = ["utf-16", "utf-8-sig", "utf-8", "cp1252"]

    for enc in encodings_to_try:
        try:
            return file_path.read_text(encoding=enc)
        except UnicodeError:
            continue

    raise UnicodeError(f"Could not decode file: {file_path}")


def convert_htmlish_md_to_clean_md(input_file: Path, output_file: Path):
    content = read_input_file(input_file)

    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".html",
        delete=False,
        encoding="utf-8"
    ) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)

    try:
        cmd = [
            "pandoc",
            str(tmp_path),
            "-f", "html",
            "-t", "markdown",
            "-o", str(output_file),
            "--wrap=none"
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip())

    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def main():
    if len(sys.argv) != 3:
        print("Usage: python clean_md_html.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()
    output_folder.mkdir(parents=True, exist_ok=True)

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    md_files = [
        p for p in input_folder.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_INPUT_EXTENSIONS
    ]

    if not md_files:
        print(f"No .md files found in: {input_folder}")
        sys.exit(0)

    for file_path in md_files:
        output_file = output_folder / f"{file_path.stem}.md"
        try:
            print(f"Converting: {file_path.name} -> {output_file.name}")
            convert_htmlish_md_to_clean_md(file_path, output_file)
        except Exception as e:
            print(f"Error converting {file_path.name}: {e}")

    print("Done.")


if __name__ == "__main__":
    main()