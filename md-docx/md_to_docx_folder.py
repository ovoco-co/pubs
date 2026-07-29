import sys
import pypandoc
from pathlib import Path


def convert_folder(folder: str):
    folder_path = Path(folder)

    if not folder_path.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")

    for md_file in folder_path.glob("*.md"):
        output_docx = md_file.with_suffix(".docx")
        pypandoc.convert_file(
            str(md_file),
            to="docx",
            outputfile=str(output_docx)
        )
        print(f"Converted {md_file.name} -> {output_docx.name}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_md_to_docx.py <folder_to_convert>")
        sys.exit(1)

    convert_folder(sys.argv[1])