"""Convert Word (.docx) files to Markdown that keeps Word's real punctuation, extracting images.

Usage:
    python docx_to_md.py <input.docx | input_folder> <output_folder> [--format FORMAT]

For each document, writes <output_folder>/<name>.md and its images to
<output_folder>/images/<name>/, with relative image links in the Markdown.
A folder input converts every .docx in it (not recursive); Word's "~$" lock
files are skipped.

--format is a pandoc output format (default: markdown-smart). Plain "markdown"
turns curly quotes into straight ones, en and em dashes into -- and ---, and
ellipses into ..., so style checkers such as Vale see pandoc's punctuation
instead of the document's. "-smart" writes the characters as they are in Word.
"""
import argparse
import subprocess
import sys
from pathlib import Path

from epub_to_md import link_safe
from output_paths import resolve_output_folder


def convert_docx(docx: Path, output_folder: Path, to_format: str) -> Path:
    output_md = output_folder / f"{docx.stem}.md"
    # Image links can't contain spaces, so the media folder gets a link-safe name.
    media_rel = "images/" + link_safe(docx.stem)

    cmd = [
        "pandoc",
        str(docx),
        "--from=docx",
        f"--to={to_format}",
        "-o",
        str(output_md),
        "--wrap=none",
        # Pass a *relative* path so markdown image links also stay relative.
        f"--extract-media={media_rel}",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(output_folder))
    if result.returncode != 0:
        raise RuntimeError(f"Pandoc failed for {docx.name}:\n{result.stderr.strip()}")
    if result.stderr.strip():
        print(f"[warn] {docx.name}: {result.stderr.strip()}")

    media_dir = output_folder / media_rel
    images = [p for p in media_dir.rglob("*") if p.is_file()] if media_dir.is_dir() else []
    print(f"Converted: {docx.name} -> {output_md.name} ({len(images)} images)")
    return output_md


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert Word files to Markdown with extracted images.")
    parser.add_argument("input", type=Path, help="A .docx file or a folder of .docx files")
    parser.add_argument("output_folder", help="Folder for the .md files and images/")
    parser.add_argument("--format", default="markdown-smart", help="Pandoc output format (default: markdown-smart)")
    args = parser.parse_args()

    source = args.input.resolve()

    if source.is_dir():
        docs = sorted(
            p for p in source.iterdir()
            if p.suffix.lower() == ".docx" and not p.name.startswith("~$")
        )
        if not docs:
            print(f"No .docx files in: {source}")
            sys.exit(1)
    elif source.is_file() and source.suffix.lower() == ".docx":
        docs = [source]
    else:
        print(f"Input must be a .docx file or a folder of them: {source}")
        sys.exit(1)

    output_folder = resolve_output_folder(args.output_folder)

    failed = 0
    for docx in docs:
        try:
            convert_docx(docx, output_folder, args.format)
        except Exception as e:
            print(f"Error converting {docx.name}: {e}")
            failed += 1

    print(f"Done. {len(docs) - failed} of {len(docs)} converted into: {output_folder}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
