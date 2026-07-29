import sys
import subprocess
from pathlib import Path

SUPPORTED_HTML_EXTENSIONS = {".htm", ".html"}


def convert_with_pandoc(input_file: Path, output_file: Path):
    cmd = [
        "pandoc",
        str(input_file),
        "-o",
        str(output_file),
        "--wrap=none"
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"Pandoc failed for {input_file.name}:\n{result.stderr.strip()}"
        )


def main():
    if len(sys.argv) != 3:
        print("Usage: python convert_html_to_md.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    html_files = [
        p for p in input_folder.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_HTML_EXTENSIONS
    ]

    if not html_files:
        print(f"No HTML files found in: {input_folder}")
        sys.exit(0)

    for file_path in html_files:
        output_file = output_folder / f"{file_path.stem}.md"

        try:
            print(f"Converting: {file_path} -> {output_file}")
            convert_with_pandoc(file_path, output_file)
        except Exception as e:
            print(f"Error converting {file_path}: {e}")

    print("Done.")


if __name__ == "__main__":
    main()