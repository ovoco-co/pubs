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
        "--wrap=none",
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
                print(
                    f"[warn] Failed to render page image {pdf_path.name} page {page_num}: {e}"
                )

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


def main() -> None:
    if len(sys.argv) != 3:
        print("Usage: python file_to_md_with_images.py <input_file> <output_folder>")
        sys.exit(1)

    input_file = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_file.is_file():
        print(f"Input file does not exist: {input_file}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)
    images_root = output_folder / "images"
    images_root.mkdir(parents=True, exist_ok=True)

    ext = input_file.suffix.lower()
    output_md = output_folder / f"{input_file.stem}.md"

    try:
        if ext in SUPPORTED_PANDOC_EXTENSIONS:
            print(f"Converting with Pandoc: {input_file.name} -> {output_md.name}")
            from_format: str | None = None
            if ext == ".xlsm":
                from_format = "xlsx"
            elif ext == ".csv":
                from_format = "csv"

            extract_media_rel = None
            if ext != ".csv":
                extract_media_rel = f"images/{input_file.stem}"
                (images_root / input_file.stem).mkdir(parents=True, exist_ok=True)

            # Try with --extract-media (works for docx/pptx; may or may not for xlsx).
            # If it fails, retry without it rather than failing the whole conversion.
            try:
                convert_with_pandoc(
                    input_file,
                    output_md,
                    extract_media_rel=extract_media_rel,
                    from_format=from_format,
                    workdir=output_folder,
                )
            except Exception as e:
                print(
                    f"[warn] Pandoc with --extract-media failed for {input_file.name}: {e}"
                )
                print(f"[warn] Retrying without extracting media for: {input_file.name}")
                extract_media_rel = None
                convert_with_pandoc(
                    input_file,
                    output_md,
                    extract_media_rel=extract_media_rel,
                    from_format=from_format,
                    workdir=output_folder,
                )

            if extract_media_rel:
                media_dir = images_root / input_file.stem
                has_any_file = any(p.is_file() for p in media_dir.rglob("*"))
                if not has_any_file:
                    print(f"[warn] No extracted images found for: {input_file.name}")

        elif ext == SUPPORTED_PDF_EXTENSION:
            print(f"Converting PDF: {input_file.name} -> {output_md.name}")
            pdf_to_markdown(input_file, output_md, images_root)

        else:
            print(f"Skipping unsupported file: {input_file.name}")
            sys.exit(0)

    except Exception as e:
        print(f"Error converting {input_file.name}: {e}")
        sys.exit(1)

    print(f"Done. Wrote: {output_md}")


if __name__ == "__main__":
    main()
