import sys
from pathlib import Path
from pypdf import PdfReader

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
        print("Usage: python pdf_to_md.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(
        [p for p in input_folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    )

    if not pdf_files:
        print(f"No PDF files found in: {input_folder}")
        sys.exit(0)

    for pdf_file in pdf_files:
        output_file = output_folder / f"{pdf_file.stem}.md"
        try:
            print(f"Converting: {pdf_file.name} -> {output_file.name}")
            pdf_to_markdown(pdf_file, output_file)
        except Exception as e:
            print(f"Error converting {pdf_file.name}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()