import sys
import re
from pathlib import Path
from bs4 import BeautifulSoup, NavigableString, Tag

SUPPORTED_HTML_EXTENSIONS = {".htm", ".html"}


def read_html_file(file_path: Path) -> str:
    encodings_to_try = ["utf-16", "utf-8-sig", "utf-8", "cp1252"]

    for enc in encodings_to_try:
        try:
            return file_path.read_text(encoding=enc)
        except UnicodeError:
            continue

    raise UnicodeError(f"Could not decode file: {file_path}")


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def cell_text(cell: Tag) -> str:
    return clean_text(cell.get_text(" ", strip=True))


def html_table_to_markdown(table: Tag) -> str:
    rows = []

    for tr in table.find_all("tr", recursive=False):
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue
        rows.append([cell_text(c) for c in cells])

    if not rows:
        return ""

    # Normalize row lengths
    max_cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < max_cols:
            r.append("")

    # If first row looks like headers, use it
    header = rows[0]
    data_rows = rows[1:] if len(rows) > 1 else []

    md = []
    md.append("| " + " | ".join(header) + " |")
    md.append("| " + " | ".join(["---"] * len(header)) + " |")

    for row in data_rows:
        md.append("| " + " | ".join(row) + " |")

    return "\n".join(md)


def extract_visible_text_and_tables(soup: BeautifulSoup) -> str:
    body = soup.body or soup

    # Remove junk that should not appear in output
    for tag in body.find_all(["script", "style", "meta", "link", "noscript"]):
        tag.decompose()

    output = []

    # Title
    title_tag = soup.find("title")
    if title_tag:
        title = clean_text(title_tag.get_text())
        if title:
            output.append(f"# {title}")
            output.append("")

    # Walk through the body and keep text/tables in order
    for element in body.descendants:
        if isinstance(element, Tag):
            # Convert top-level-ish tables only once
            if element.name == "table":
                # Avoid nested duplicate conversion
                parent_table = element.find_parent("table")
                if parent_table is not None and parent_table != element:
                    continue

                md_table = html_table_to_markdown(element)
                if md_table:
                    output.append(md_table)
                    output.append("")

            # Headings
            elif element.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                txt = clean_text(element.get_text(" ", strip=True))
                if txt:
                    level = int(element.name[1])
                    output.append(f"{'#' * level} {txt}")
                    output.append("")

            # Paragraphs
            elif element.name == "p":
                txt = clean_text(element.get_text(" ", strip=True))
                if txt:
                    output.append(txt)
                    output.append("")

    # Clean up repeated blanks
    cleaned = []
    prev_blank = False
    for line in output:
        blank = (line.strip() == "")
        if blank and prev_blank:
            continue
        cleaned.append(line)
        prev_blank = blank

    return "\n".join(cleaned).strip() + "\n"


def convert_file(input_file: Path, output_file: Path):
    raw_html = read_html_file(input_file)
    soup = BeautifulSoup(raw_html, "html.parser")
    md = extract_visible_text_and_tables(soup)
    output_file.write_text(md, encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        print("Usage: python html_to_md_clean.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_folder = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Input folder does not exist: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    html_files = [
        p for p in input_folder.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_HTML_EXTENSIONS
    ]

    if not html_files:
        print(f"No HTML files found in: {input_folder}")
        sys.exit(0)

    for file_path in html_files:
        output_file = output_folder / f"{file_path.stem}.md"
        try:
            print(f"Converting: {file_path.name} -> {output_file.name}")
            convert_file(file_path, output_file)
        except Exception as e:
            print(f"Error converting {file_path.name}: {e}")

    print("Done.")


if __name__ == "__main__":
    main()