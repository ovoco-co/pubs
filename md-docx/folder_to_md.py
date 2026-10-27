import sys
import subprocess
from pathlib import Path

SUPPORTED_PANDOC_EXTENSIONS = {".docx", ".pptx", ".xlsx", ".xlsm", ".csv"}
SUPPORTED_PDF_EXTENSION = ".pdf"

def convert_with_pandoc(
    input_file: Path,
    output_file: Path,
    *,
    extract_media_rel: str | None = None,
    from_format: str | None = None,
    workdir: Path | None = None,
) -> None:
    cmd = [
        "pandoc",
        str(input_file),
        "-o",
        str(output_file),
        "--wrap=none"
    ]

    if from_format:
        cmd.append(f"--from={from_format}")

    if extract_media_rel:
        # Pass a *relative* path so markdown image links also stay relative.
        cmd.append(f"--extract-media={extract_media_rel}")

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(workdir) if workdir else None,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"Pandoc failed for {input_file.name}:\n{result.stderr.strip()}"
        )

def pdf_to_markdown(pdf_path: Path, output_path: Path, images_root: Path) -> None:
    # PyMuPDF provides either:
    # - `import pymupdf` (preferred in newer versions)
    # - `import fitz` (older versions; still works)
    try:
        import pymupdf  # type: ignore
        doc = pymupdf.open(str(pdf_path))
    except ModuleNotFoundError:
        try:
            import fitz  # type: ignore
            doc = fitz.open(str(pdf_path))
        except ModuleNotFoundError as e:
            raise ModuleNotFoundError(
                "PyMuPDF is required to extract and save PDF images. "
                "Install with: python -m pip install pymupdf"
            ) from e

    file_images_dir = images_root / pdf_path.stem
    file_images_dir.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as out:
        out.write(f"# {pdf_path.stem}\n\n")
        out.write(f"*Source file: {pdf_path.name}*\n\n")

        total_images = 0
        for i in range(len(doc)):
            page_num = i + 1
            page = doc[i]

            out.write(f"## Page {page_num}\n\n")

            text = (page.get_text("text") or "").strip()
            if text:
                out.write(text + "\n\n")
            else:
                out.write("_No text extracted from this page._\n\n")

            # Always render a page screenshot so image output isn't empty even when
            # the PDF contains no extractable embedded bitmap images.
            try:
                page_png_name = f"page_{page_num}.png"
                page_png_path = file_images_dir / page_png_name
                pix = page.get_pixmap(dpi=200)
                pix.save(str(page_png_path))
                rel_path = Path("images") / pdf_path.stem / page_png_name
                out.write(f"![Page {page_num}]({rel_path.as_posix()})\n\n")
            except Exception as e:
                print(f"[warn] Failed to render page image {pdf_path.name} page {page_num}: {e}")

            image_list = page.get_images(full=True)
            if not image_list:
                continue

            total_images += len(image_list)

            out.write("_Extracted images:_\n\n")

            for img_index, img in enumerate(image_list, start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image.get("ext") or "png"

                image_name = f"page_{page_num}_img_{img_index}.{image_ext}"
                image_path = file_images_dir / image_name

                with image_path.open("wb") as f:
                    f.write(image_bytes)

                rel_path = Path("images") / pdf_path.stem / image_name
                out.write(f"![Image]({rel_path.as_posix()})\n\n")

        if total_images == 0:
            print(f"[warn] No embedded images found in: {pdf_path}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python folder_to_md.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"
    images_root.mkdir(parents=True, exist_ok=True)

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
                extract_media_rel = None
                from_format: str | None = None

                if ext == ".xlsm":
                    from_format = "xlsx"
                elif ext == ".csv":
                    from_format = "csv"

                # For PPTX/Docx, pandoc can extract embedded images into --extract-media.
                if ext in {".pptx", ".docx"}:
                    extract_media_rel = f"images/{file_path.stem}"
                    (images_root / file_path.stem).mkdir(parents=True, exist_ok=True)
                convert_with_pandoc(
                    file_path,
                    output_file,
                    extract_media_rel=extract_media_rel,
                    from_format=from_format,
                    workdir=output_folder,
                )
                if extract_media_rel:
                    media_dir = images_root / file_path.stem
                    has_any_file = any(
                        p.is_file() for p in media_dir.rglob("*")
                    )
                    if not has_any_file:
                        print(f"[warn] No extracted images found for: {file_path.name}")

            elif ext == SUPPORTED_PDF_EXTENSION:
                print(f"Converting PDF: {file_path.name} -> {output_file.name}")
                pdf_to_markdown(file_path, output_file, images_root)

            else:
                print(f"Skipping unsupported file: {file_path.name}")

        except Exception as e:
            print(f"Error converting {file_path.name}: {e}")

    print("Done.")

if __name__ == "__main__":
    main()
