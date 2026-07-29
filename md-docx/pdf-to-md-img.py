import os
import sys
import fitz  # PyMuPDF


def pdf_to_markdown(pdf_path: str, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    md_lines = []

    for page_num in range(len(doc)):
        page = doc[page_num]

        md_lines.append(f"\n\n# Page {page_num + 1}\n")

        # Extract text
        text = page.get_text("text").strip()
        if text:
            md_lines.append(text + "\n")

        # Extract images
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list, start=1):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]

            image_name = f"page_{page_num + 1}_img_{img_index}.{image_ext}"
            image_path = os.path.join(images_dir, image_name)

            with open(image_path, "wb") as f:
                f.write(image_bytes)

            md_lines.append(f"![Image](images/{image_name})\n")

    md_path = os.path.join(output_dir, "document.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    return md_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_md.py <pdf_path> [output_dir]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"

    if not os.path.exists(pdf_path):
        print(f"Error: PDF not found: {pdf_path}")
        sys.exit(1)

    result = pdf_to_markdown(pdf_path, output_dir)
    print(f"Saved markdown to: {result}")


if __name__ == "__main__":
    main()