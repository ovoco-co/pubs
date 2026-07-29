import os
import sys
import subprocess
from pathlib import Path

def convert_doc_to_txt(doc_path: Path, output_dir: Path) -> Path:
    """
    Convert a .doc file to .txt using LibreOffice headless conversion.
    Returns the path to the generated .txt file.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice will create a text file with the same basename
    cmd = [
        "soffice",
        "--headless",
        "--convert-to",
        "txt:Text",
        "--outdir",
        str(output_dir),
        str(doc_path),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"Failed to convert {doc_path.name}:\n{result.stderr.strip()}"
        )

    txt_path = output_dir / f"{doc_path.stem}.txt"
    if not txt_path.exists():
        raise FileNotFoundError(f"Expected output not found: {txt_path}")

    return txt_path


def combine_text_files(text_files, combined_output: Path, as_markdown: bool = True):
    """
    Combine multiple text files into one output file.
    If as_markdown is True, adds headings and code fences for readability.
    """
    with combined_output.open("w", encoding="utf-8") as out:
        for txt_file in text_files:
            content = txt_file.read_text(encoding="utf-8", errors="ignore")

            if as_markdown:
                out.write(f"# {txt_file.stem}\n\n")
                out.write("```text\n")
                out.write(content.rstrip() + "\n")
                out.write("```\n\n")
            else:
                out.write(f"===== {txt_file.stem} =====\n\n")
                out.write(content.rstrip() + "\n\n")


def main():
    if len(sys.argv) < 3:
        print(
            "Usage: python convert_docs.py <input_folder> <output_file> [--txt]\n"
            "Example: python convert_docs.py ./docs combined.md"
        )
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_file = Path(sys.argv[2]).resolve()
    as_markdown = "--txt" not in sys.argv

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    work_dir = input_folder / "_converted_txt"

    doc_files = sorted(input_folder.glob("*.doc"))
    if not doc_files:
        print(f"No .doc files found in: {input_folder}")
        sys.exit(0)

    converted_files = []
    for doc_file in doc_files:
        print(f"Converting: {doc_file.name}")
        converted_files.append(convert_doc_to_txt(doc_file, work_dir))

    print(f"Combining into: {output_file.name}")
    combine_text_files(converted_files, output_file, as_markdown=as_markdown)

    print("Done.")


if __name__ == "__main__":
    main()