import sys
import pypandoc
from pathlib import Path

from output_paths import resolve_output_folder


def convert_folder(folder: str, output_folder: Path):
    folder_path = Path(folder)

    if not folder_path.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")

    for md_file in folder_path.glob("*.md"):
        output_docx = output_folder / f"{md_file.stem}.docx"
        pypandoc.convert_file(
            str(md_file),
            to="docx",
            outputfile=str(output_docx)
        )
        print(f"Converted {md_file.name} -> {output_docx.name}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python md_to_docx_folder.py <input_folder> <output_folder>")
        sys.exit(1)

    convert_folder(sys.argv[1], resolve_output_folder(sys.argv[2]))
