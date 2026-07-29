import sys
import subprocess
from pathlib import Path
from pypdf import PdfReader

SUPPORTED_PANDOC_EXTENSIONS = {".docx", ".pptx", ".xlsx"}
SUPPORTED_PDF_EXTENSION = ".pdf"

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

def pdf_to_markdown(pdf_path: Path, output_path: Path):
    reader = PdfReader(str(pdf_path))

    with output_path.open("w", encoding="utf-8") as out:
        out.write(f"# {pdf_path.stem}\n\n")
        out.write(f"*Source file: {pdf_path.name}*\n\n")

        for i, page in enumerate(reader.pages, start=1):
            text = page.extract_text()

            out.write(f"## Page {i}\n\n")

            if text and text.strip():
                out.write(text.strip() + "\n\n")
            else:
                out.write("_No text extracted from this page._\n\n")

def main():
    if len(sys.argv) != 3:
        print("Usage: python convert_folder_to_md.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    files = sorted([p for p in input_folder.iterdir() if p.is_file()])

    if not files:
        print(f"No files found in: {input_folder}")
        sys.exit(0)

    for file_path in files:
        ext = file_path.suffix.lower()
        output_file = output_folder / f"{file_path.stem}.md"

        try:
            if ext in SUPPORTED_PANDOC_EXTENSIONS:
                print(f"Converting with Pandoc: {file_path.name} -> {output_file.name}")
                convert_with_pandoc(file_path, output_file)

            elif ext == SUPPORTED_PDF_EXTENSION:
                print(f"Converting PDF: {file_path.name} -> {output_file.name}")
                pdf_to_markdown(file_path, output_file)

            else:
                print(f"Skipping unsupported file: {file_path.name}")

        except Exception as e:
            print(f"Error converting {file_path.name}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()