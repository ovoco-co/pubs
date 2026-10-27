from __future__ import annotations

from pathlib import Path
import sys
import csv
import re


def extract_first_md_path(text: str) -> str | None:
    s = text.strip()
    if not s or ".md" not in s.lower():
        return None

    # Markdown link syntax: [...](path/to/file.md)
    if "(" in s and ")" in s:
        start = s.rfind("(")
        end = s.rfind(")")
        if start != -1 and end != -1 and end > start:
            cand = s[start + 1 : end].strip().strip("<>")
            if cand.lower().endswith(".md"):
                return cand

    # Plain path, possibly embedded in other text.
    m = re.search(r"([^\s\)]+\.md)\b", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def read_md_files_from_csv(csv_path: Path) -> list[Path]:
    base_dir = csv_path.parent
    md_files: list[Path] = []
    seen: set[Path] = set()

    with csv_path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            for cell in row:
                cand = extract_first_md_path(cell)
                if not cand:
                    continue

                p = Path(cand)
                if not p.is_absolute():
                    p = (base_dir / p).resolve()
                else:
                    p = p.resolve()

                if p.suffix.lower() != ".md":
                    continue
                if not p.exists():
                    raise FileNotFoundError(f"[csv] Missing file: {p}")
                if p in seen:
                    continue
                seen.add(p)
                md_files.append(p)

    return md_files


def combine_md_files(md_files: list[Path], combined_path: Path) -> None:
    combined_path.parent.mkdir(parents=True, exist_ok=True)
    with combined_path.open("w", encoding="utf-8") as out:
        for i, md_file in enumerate(md_files):
            content = md_file.read_text(encoding="utf-8", errors="ignore").strip()
            out.write(content)
            if i != len(md_files) - 1:
                out.write("\n\n")


def main() -> None:
    if len(sys.argv) != 3:
        print("Usage: python combine-recursive.py <input_folder|input.csv> <output_file.md>")
        sys.exit(1)

    input_path = Path(sys.argv[1]).resolve()
    output_file = Path(sys.argv[2]).resolve()

    if output_file.suffix.lower() != ".md":
        print("Output must be a .md file path.")
        sys.exit(1)

    md_files: list[Path] = []
    if input_path.is_dir():
        md_files = sorted(
            (
                p
                for p in input_path.rglob("*.md")
                if p.resolve() != output_file
            ),
            key=lambda p: p.relative_to(input_path).as_posix(),
        )
        if not md_files:
            print(f"No *.md files found in {input_path}")
            sys.exit(0)
    elif input_path.is_file() and input_path.suffix.lower() == ".csv":
        md_files = read_md_files_from_csv(input_path)
        # Also skip if user accidentally listed the output itself.
        md_files = [p for p in md_files if p.resolve() != output_file]
        if not md_files:
            print(f"No *.md entries found in {input_path}")
            sys.exit(0)
    else:
        print(f"Not a folder or .csv file: {input_path}")
        sys.exit(1)

    combine_md_files(md_files, output_file)

    # Write a simple manifest CSV (best-effort).
    manifest_csv = output_file.with_suffix(".csv")
    try:
        with manifest_csv.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["index", "source_md_path"])
            for i, p in enumerate(md_files, start=1):
                writer.writerow([i, str(p)])
    except Exception as e:
        print(f"[warn] Could not write manifest CSV: {e}")

    print(f"Done. Wrote: {output_file}")


if __name__ == "__main__":
    main()
