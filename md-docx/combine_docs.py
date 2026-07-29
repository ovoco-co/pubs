from pathlib import Path
import subprocess
import sys

def convert_docx_to_md(docx_path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    md_path = output_dir / f"{docx_path.stem}.md"

    result = subprocess.run(
        ["pandoc", str(docx_path), "-o", str(md_path), "--wrap=none"],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(f"Failed to convert {docx_path.name}:\n{result.stderr}")

    return md_path

def combine_md_files(md_files, combined_path: Path):
    with combined_path.open("w", encoding="utf-8") as out:
        for md_file in md_files:
            out.write(f"# {md_file.stem}\n\n")
            out.write(md_file.read_text(encoding="utf-8", errors="ignore").strip())
            out.write("\n\n---\n\n")

def main():
    if len(sys.argv) != 3:
        print("Usage: python combine_docs.py <input_folder> <output_file>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_file = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Not a folder: {input_folder}")
        sys.exit(1)

    md_work_dir = input_folder / "_md_temp"
    docx_files = sorted(input_folder.glob("*.docx"))

    if not docx_files:
        print(f"No .docx files found in {input_folder}")
        sys.exit(0)

    converted = []
    for f in docx_files:
        print(f"Converting {f.name}...")
        converted.append(convert_docx_to_md(f, md_work_dir))

    print(f"Combining into {output_file}...")
    combine_md_files(converted, output_file)
    print("Done.")

if __name__ == "__main__":
    main()