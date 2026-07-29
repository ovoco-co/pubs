import argparse
import pypandoc
from pathlib import Path


def md_to_docx(input_md: str, output_docx: str = None, reference_doc: str = None):
    input_path = Path(input_md)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_md}")

    if output_docx is None:
        output_docx = str(input_path.with_suffix(".docx"))

    extra_args = []
    if reference_doc:
        ref_path = Path(reference_doc)
        if not ref_path.exists():
            raise FileNotFoundError(f"Reference doc not found: {reference_doc}")
        extra_args.append(f"--reference-doc={ref_path}")

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
    parser.add_argument("-o", "--output", help="Path to the output .docx file")
    parser.add_argument("-r", "--reference-doc", help="Path to a reference DOCX file")

    args = parser.parse_args()
    md_to_docx(args.input_md, args.output, args.reference_doc)


if __name__ == "__main__":
    main()