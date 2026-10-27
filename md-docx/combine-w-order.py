from __future__ import annotations

from pathlib import Path
import sys
import csv


def read_order_file(order_path: Path) -> list[str]:
    tokens: list[str] = []
    for raw in order_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        tokens.append(line)
    return tokens


def read_order_csv(order_csv_path: Path) -> list[str]:
    tokens: list[str] = []
    with order_csv_path.open("r", encoding="utf-8", errors="ignore", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            # Take only the first non-empty cell per row.
            for cell in row:
                tok = cell.strip()
                if not tok:
                    continue
                if tok.startswith("#"):
                    break
                tokens.append(tok)
                break
    return tokens


def token_to_stem(tok: str) -> str:
    t = tok.strip()
    if t.lower().endswith(".md"):
        return Path(t).stem
    return Path(t).name


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
        print("Usage: python combine-w-order.py <input_folder> <output_file.md>")
        sys.exit(1)

    input_folder = Path(sys.argv[1]).resolve()
    output_file = Path(sys.argv[2]).resolve()

    if not input_folder.is_dir():
        print(f"Not a folder: {input_folder}")
        sys.exit(1)

    if output_file.suffix.lower() != ".md":
        print(
            "Output must be a .md file path (second argument), e.g. "
            r"\"C:\\Users\\...\\formark\\Tasks-combined.md\""
        )
        sys.exit(1)

    md_files = sorted([p for p in input_folder.glob("*.md") if p.name != ".order"])
    if not md_files:
        print(f"No .md files found in {input_folder}")
        sys.exit(0)

    order_path = input_folder / ".order"
    order_csv_path = input_folder / ".order.csv"

    tokens: list[str] | None = None
    if order_path.exists():
        tokens = read_order_file(order_path)
    elif order_csv_path.exists():
        tokens = read_order_csv(order_csv_path)
        # If CSV has a header row, drop leading tokens until the first token that
        # matches any *.md stem.
        stems_lower = {p.stem.lower() for p in md_files}
        first_match_idx = None
        for i, tok in enumerate(tokens):
            if token_to_stem(tok).lower() in stems_lower:
                first_match_idx = i
                break
        if first_match_idx is not None and first_match_idx > 0:
            tokens = tokens[first_match_idx:]

    if tokens is not None:
        stem_to_path = {p.stem: p for p in md_files}
        stem_to_path_lower = {p.stem.lower(): p for p in md_files}
        included_paths: set[Path] = set()
        ordered: list[Path] = []

        for tok in tokens:
            tok_stem = token_to_stem(tok)
            chosen: Path | None = None
            if tok_stem in stem_to_path:
                chosen = stem_to_path[tok_stem]
            elif tok_stem.lower() in stem_to_path_lower:
                chosen = stem_to_path_lower[tok_stem.lower()]
            if chosen is not None:
                if chosen not in included_paths:
                    ordered.append(chosen)
                    included_paths.add(chosen)
                continue

            # Fallback to filename matches.
            candidate = (
                input_folder / tok
                if tok.lower().endswith(".md")
                else input_folder / f"{tok}.md"
            )
            if candidate.exists():
                if candidate not in included_paths:
                    ordered.append(candidate)
                    included_paths.add(candidate)
                continue

            raise FileNotFoundError(
                f"[.order] entry '{tok}' does not match any *.md file in: {input_folder}"
            )

        remaining = [p for p in md_files if p not in included_paths]
        if remaining:
            print(
                f"[warn] {len(remaining)} *.md file(s) not listed in .order; appending them alphabetically."
            )
            ordered.extend(sorted(remaining, key=lambda p: p.name))

        md_files = ordered
    else:
        print(f"[warn] No .order found in {input_folder}; using alphabetical md order.")

    combine_md_files(md_files, output_file)

    # Best-effort manifest CSV.
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
