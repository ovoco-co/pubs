import argparse
import re
import tempfile
import zipfile
import pypandoc
from pathlib import Path

from output_paths import resolve_output_file


# Pandoc Lua filter: drop images; a figure's caption stays as text.
NO_IMAGES_FILTER = r"""
function Figure(el)
  local caption = pandoc.utils.blocks_to_inlines(el.caption.long)
  if #caption == 0 then return {} end
  return pandoc.Para(caption)
end
function Image(el) return {} end
"""


def styles_only_reference(reference_doc: Path, folder: Path) -> Path:
    """Copy a reference .docx without the main document's images.

    Pandoc copies the reference document's images into every output, even though the
    output never uses them. Images used by headers and footers are kept.
    """
    out_path = folder / "reference.docx"
    with zipfile.ZipFile(reference_doc) as src:
        names = src.namelist()
        other_rels = "".join(
            src.read(n).decode("utf-8") for n in names
            if n.endswith(".rels") and n != "word/_rels/document.xml.rels"
        )
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as dst:
            for name in names:
                data = src.read(name)
                if name.startswith("word/media/") and name.split("/")[-1] not in other_rels:
                    continue
                if name == "word/_rels/document.xml.rels":
                    data = re.sub(rb'<Relationship [^>]*Type="[^"]*/image"[^>]*/>', b"", data)
                dst.writestr(name, data)
    return out_path


def md_to_docx(input_md: str, output: str, reference_doc: str = None, no_images: bool = False):
    input_path = Path(input_md)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_md}")

    output_docx = str(resolve_output_file(output, f"{input_path.stem}.docx", (".docx",)))

    # Find images relative to the .md file, not the folder the script is run from.
    extra_args = [f"--resource-path={input_path.resolve().parent}"]
    if reference_doc:
        ref_path = Path(reference_doc)
        if not ref_path.exists():
            raise FileNotFoundError(f"Reference doc not found: {reference_doc}")

    with tempfile.TemporaryDirectory() as tmp:
        if reference_doc:
            extra_args.append(f"--reference-doc={styles_only_reference(ref_path, Path(tmp))}")
        if no_images:
            filter_path = Path(tmp) / "no_images.lua"
            filter_path.write_text(NO_IMAGES_FILTER, encoding="utf-8")
            extra_args.append(f"--lua-filter={filter_path}")

        pypandoc.convert_file(
            str(input_path),
            to="docx",
            outputfile=output_docx,
            extra_args=extra_args
        )

    print(f"Converted {input_md} -> {output_docx}")


def main():
    parser = argparse.ArgumentParser(description="Convert Markdown to DOCX")
    parser.add_argument("input_md", help="Path to the input .md file")
    parser.add_argument("output", help="Output .docx file, or a folder (the file is named after the input)")
    parser.add_argument("-r", "--reference-doc", help="Path to a reference DOCX file")
    parser.add_argument("--no-images", action="store_true", help="Leave images out of the Word file")

    args = parser.parse_args()
    md_to_docx(args.input_md, args.output, args.reference_doc, args.no_images)


if __name__ == "__main__":
    main()