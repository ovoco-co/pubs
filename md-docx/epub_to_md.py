"""Convert EPUB files to Markdown, extracting images into a folder.

Usage:
    python epub_to_md.py <input.epub | input_folder> <output_folder> [--format FORMAT]

For each EPUB, writes <output_folder>/<name>.md and its images to
<output_folder>/images/<name>/, with relative image links in the Markdown.
A folder input converts every .epub in it (not recursive).

--format is a pandoc output format (default: gfm). Use "markdown" for pandoc
Markdown. Either way, a Lua filter removes the EPUB's empty <span>/<div>
wrappers and turns <figure> blocks into plain Markdown images with the
caption below. Tables pandoc can't express in Markdown (merged cells) stay as
HTML rather than being dropped.
"""
import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path

# Pandoc Lua filter: unwrap EPUB section divs and anchor spans, flatten figures.
CLEANUP_FILTER = r"""
function Div(el) return el.content end
function Span(el) return el.content end
function Figure(el)
  local blocks = {}
  local alt = ""
  for _, block in ipairs(el.content) do
    if block.t == "Plain" or block.t == "Para" then
      for _, inline in ipairs(block.content) do
        if inline.t == "Image" then alt = pandoc.utils.stringify(inline.caption) end
      end
      table.insert(blocks, pandoc.Para(block.content))
    else
      table.insert(blocks, block)
    end
  end
  local caption = pandoc.utils.stringify(el.caption.long)
  if caption ~= "" and caption ~= alt then
    table.insert(blocks, pandoc.Para({pandoc.Emph(pandoc.utils.blocks_to_inlines(el.caption.long))}))
  end
  return blocks
end
"""


def link_safe(name: str) -> str:
    """Make a name safe for a Markdown image link: spaces and other characters become hyphens."""
    return re.sub(r"[^\w.-]+", "-", name).strip("-") or "file"


def write_cleanup_filter(folder: Path) -> Path:
    """Write the EPUB cleanup filter into `folder` and return its path (used by the folder converters too)."""
    filter_path = folder / "epub_cleanup.lua"
    filter_path.write_text(CLEANUP_FILTER, encoding="utf-8")
    return filter_path


def convert_epub(epub: Path, output_folder: Path, to_format: str, filter_path: Path) -> Path:
    output_md = output_folder / f"{epub.stem}.md"
    # Image links can't contain spaces, so the media folder gets a link-safe name.
    media_rel = "images/" + link_safe(epub.stem)
    (output_folder / media_rel).mkdir(parents=True, exist_ok=True)

    cmd = [
        "pandoc",
        str(epub),
        "--from=epub",
        f"--to={to_format}",
        "-o",
        str(output_md),
        "--wrap=none",
        f"--lua-filter={filter_path}",
        # Pass a *relative* path so markdown image links also stay relative.
        f"--extract-media={media_rel}",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(output_folder))
    if result.returncode != 0:
        raise RuntimeError(f"Pandoc failed for {epub.name}:\n{result.stderr.strip()}")
    if result.stderr.strip():
        print(f"[warn] {epub.name}: {result.stderr.strip()}")

    images = [p for p in (output_folder / media_rel).rglob("*") if p.is_file()]
    if not images:
        print(f"[warn] No images found in: {epub.name}")
    print(f"Converted: {epub.name} -> {output_md.name} ({len(images)} images)")
    return output_md


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert EPUB files to Markdown with extracted images.")
    parser.add_argument("input", type=Path, help="An .epub file or a folder of .epub files")
    parser.add_argument("output_folder", type=Path, help="Folder for the .md files and images/")
    parser.add_argument("--format", default="gfm", help="Pandoc output format (default: gfm)")
    args = parser.parse_args()

    source = args.input.resolve()
    output_folder = args.output_folder.resolve()

    if source.is_dir():
        epubs = sorted(p for p in source.iterdir() if p.suffix.lower() == ".epub")
        if not epubs:
            print(f"No .epub files in: {source}")
            sys.exit(1)
    elif source.is_file() and source.suffix.lower() == ".epub":
        epubs = [source]
    else:
        print(f"Input must be an .epub file or a folder of them: {source}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    failed = 0
    with tempfile.TemporaryDirectory() as tmp:
        filter_path = write_cleanup_filter(Path(tmp))
        for epub in epubs:
            try:
                convert_epub(epub, output_folder, args.format, filter_path)
            except Exception as e:
                print(f"Error converting {epub.name}: {e}")
                failed += 1

    print(f"Done. {len(epubs) - failed} of {len(epubs)} converted into: {output_folder}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
