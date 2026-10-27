from __future__ import annotations

from pathlib import Path
import csv
import re
import sys
import shutil
from typing import Match


def extract_first_md_path(line: str) -> str | None:
    s = line.strip()
    if not s or s.startswith("#"):
        return None
    if ".md" not in s.lower():
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


def _posix_path(s: str) -> str:
    return s.replace("\\", "/")


def _strip_md_image_path(token: str) -> str:
    t = token.strip()
    if t.startswith("<") and t.endswith(">"):
        t = t[1:-1].strip()
    if (t.startswith('"') and t.endswith('"')) or (t.startswith("'") and t.endswith("'")):
        t = t[1:-1].strip()
    return t


def _first_token_with_rest(inner: str) -> tuple[str, str] | None:
    # Grab the first token (path) and keep the remaining text (likely an optional title).
    # Supports: <path>, "path", 'path', or bare path (no spaces).
    m = re.match(r"^\s*(<[^>]+>|\"[^\"]+\"|'[^']+'|\S+)(?P<rest>.*)$", inner)
    if not m:
        return None
    token = m.group(1)
    rest = m.group("rest") or ""
    return token, rest


def _resolve_image_source(md_path: Path, link_path: str) -> Path:
    # link_path is already stripped of any surrounding quotes/brackets.
    if re.match(r"^[a-zA-Z]:[\\/]", link_path):
        return Path(link_path)
    return (md_path.parent / link_path).resolve()


def _dest_rel_for_image(md_path: Path, link_path: str) -> str:
    # Prefer preserving anything after the first "images/" path segment.
    norm = _posix_path(link_path)
    parts = norm.split("/")

    idx = None
    for i, part in enumerate(parts):
        if part.lower() == "images":
            idx = i
            break

    if idx is not None and idx + 1 < len(parts):
        rel_parts = parts[idx:]
        return "/".join(rel_parts)

    # Fallback: flatten into images/ to avoid collisions.
    name = Path(parts[-1]).name
    return f"images/{md_path.stem}__{name}"


def combine_md_files(md_files: list[Path], combined_path: Path) -> None:
    combined_path.parent.mkdir(parents=True, exist_ok=True)
    combined_dir = combined_path.parent
    images_copied: set[Path] = set()
    image_pattern = re.compile(
        r"!\[(?P<alt>[^\]]*)\]\((?P<inner>[^)]+)\)"
    )
    with combined_path.open("w", encoding="utf-8") as out:
        for i, md_file in enumerate(md_files):
            text = md_file.read_text(encoding="utf-8", errors="ignore").rstrip()

            # Copy referenced images into the combined folder and rewrite links.
            def repl(m: Match[str]) -> str:
                alt = m.group("alt")
                inner = m.group("inner")
                inner_stripped = inner.strip()

                # Skip empty.
                if not inner_stripped:
                    return m.group(0)

                tok_rest = _first_token_with_rest(inner_stripped)
                if not tok_rest:
                    return m.group(0)

                token, rest = tok_rest
                link_path = _strip_md_image_path(token)
                lp_lower = link_path.lower()
                if lp_lower.startswith(("http://", "https://", "data:")):
                    return m.group(0)

                src_img = _resolve_image_source(md_file, link_path)
                if not src_img.exists() or not src_img.is_file():
                    # Leave the link untouched if we can't find the source image.
                    return m.group(0)

                dest_rel = _dest_rel_for_image(md_file, link_path)
                dest_img = (combined_dir / dest_rel).resolve()
                if dest_img not in images_copied:
                    dest_img.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src_img, dest_img)
                    images_copied.add(dest_img)

                new_inner = dest_rel
                rest2 = rest.rstrip()
                if rest2:
                    new_inner = f"{dest_rel}{rest2}"

                return f"![{alt}]({new_inner})"

            text = image_pattern.sub(repl, text)

            out.write(text)
            if i != len(md_files) - 1:
                out.write("\n\n")


def main() -> None:
    if len(sys.argv) not in (2, 3):
        print("Usage: python combine-from-links.py <links_file> [output_root]")
        sys.exit(1)

    links_file = Path(sys.argv[1]).resolve()
    if not links_file.is_file():
        print(f"Not a file: {links_file}")
        sys.exit(1)

    output_root = Path(sys.argv[2]).resolve() if len(sys.argv) == 3 else links_file.parent
    if not output_root.exists():
        output_root.mkdir(parents=True, exist_ok=True)
    if not output_root.is_dir():
        print(f"Not a folder: {output_root}")
        sys.exit(1)

    base_dir = links_file.parent
    md_files: list[Path] = []
    seen: set[Path] = set()

    if links_file.suffix.lower() == ".csv":
        with links_file.open("r", encoding="utf-8", errors="ignore", newline="") as f:
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
    else:
        for raw in links_file.read_text(encoding="utf-8", errors="ignore").splitlines():
            cand = extract_first_md_path(raw)
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
                raise FileNotFoundError(f"[links] Missing file: {p}")
            if p in seen:
                continue
            seen.add(p)
            md_files.append(p)

    if not md_files:
        print("No *.md entries found in links file.")
        sys.exit(0)

    first = md_files[0]
    first_stem = first.stem

    out_dir = output_root / first_stem
    out_md = out_dir / f"{first_stem}-combined.md"

    combine_md_files(md_files, out_md)
    print(f"Done. Wrote: {out_md}")

    # Best-effort manifest CSV.
    manifest_csv = out_md.with_suffix(".csv")
    try:
        with manifest_csv.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["index", "source_md_path"])
            for i, p in enumerate(md_files, start=1):
                writer.writerow([i, str(p)])
    except Exception as e:
        print(f"[warn] Could not write manifest CSV: {e}")


if __name__ == "__main__":
    main()
