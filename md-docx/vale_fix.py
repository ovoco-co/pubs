"""Run Vale on a folder of Markdown files and apply the mechanical HM style fixes.

Usage:
    python vale_fix.py <md_folder> <log_file.md|output_folder> [--config PATH] [--dry-run]

Edits the .md files in place, only at the exact text Vale flagged, and writes a
log of every change (default name: vale-fix-log.md). A copy of the folder is
saved first as <md_folder>-before-vale-fix-<date-time>, unless --dry-run is given.

Fixed automatically: 37.4% -> 37.4 %, "44 percent" -> 44 %, underlined citations,
"the island" -> Puerto Rico,
healthcare/toward/gray, place-name accents, date ordinals, Federal Government,
US -> U.S./United States, "Section 3.1" -> "Sec. 3.1" mid-sentence, i.e., a few
wordy phrases (in order to -> to), and unit spacing (130mph -> 130 mph).

Left for a person: acronyms, spelled-out numbers, list punctuation, captions,
and anything in reference-list entries, title translations, or table rows.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from output_paths import resolve_output_file

DEFAULT_CONFIG = r"C:\Users\Artist\projects\HM\.vale.ini"

REF_LINE = re.compile(
    r"^\s*(?:\[\^\d+\]:\s*)?[^\s(][^\n]{0,200}?\((?:\d{4}[a-z]?|n\.d\.)(?:, [A-Z][a-z]+(?: \d+)?)?\)\.\s"  # Author. (2017, May 4).
    r"|^\s*(?:\[\^\d+\]:\s*)?[A-Z][^.\n]{0,120}\. \d{4}[a-z]?\. "  # NIST. 2022. Title
    r"|\bRetrieved (?:from|[A-Z][a-z]+ \d)"
)
# Rows of Pandoc simple tables: indented, with wide gaps between columns.
TABLE_ROW = re.compile(r"^ {2,}\S.*\S {3,}\S")
DETERMINERS = {"the", "these", "those", "their", "its", "them", "our", "this", "all", "his", "her", "which", "whom", "us"}


def find_vale() -> str:
    vale = shutil.which("vale")
    if vale:
        return vale
    winget = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WinGet" / "Links" / "vale.exe"
    if winget.exists():
        return str(winget)
    raise SystemExit("Vale not found. Install it with: winget install errata-ai.Vale")


def cap_like(src: str, new: str) -> str:
    return new[0].upper() + new[1:] if src[:1].isupper() else new


def fix(rule: str, m: str, before: str, after: str):
    """Return the replacement for match m, or None to leave it.

    May return (new_text, n) to also replace the n characters before the match.
    """
    if rule == "HM.PercentSpacing":
        return re.sub(r"\s*%$", " %", m)
    if rule == "HM.Percent":
        return re.sub(r"\s+percent$", " %", m)
    if rule == "HM.Underline":
        return re.sub(r"\[([^\[\]]*)\]\{\.underline\}", r"\1", m)
    if rule == "HM.Island":
        if "wide" in m.lower() or m.lower() == "an island" or after.startswith("-"):
            return None
        if before.endswith("on "):
            return ("in Puerto Rico", 3)  # "on the island" -> "in Puerto Rico"
        return "Puerto Rico"
    if rule == "HM.Spelling":
        return {"health care": "healthcare", "towards": "toward", "Towards": "Toward", "grey": "gray"}.get(m)
    if rule == "HM.Diacritics":
        return {"Mayaguez": "Mayagüez", "Catano": "Cataño", "San German": "San Germán", "Loiza": "Loíza",
                "Rio Grande": "Río Grande", "Bayamon": "Bayamón", "Rincon": "Rincón", "Guanica": "Guánica",
                "Penuelas": "Peñuelas", "Anasco": "Añasco", "Comerio": "Comerío", "Canovanas": "Canóvanas",
                "San Sebastian": "San Sebastián"}.get(m)
    if rule == "HM.DateOrdinal":
        return re.sub(r"(\d+)(?:st|nd|rd|th)$", r"\1", m)
    if rule == "HM.Ordinals":
        return re.sub(r"(\d+)\^(st|nd|rd|th)\^", r"\1\2", m)
    if rule == "HM.Government":
        return {"the federal government": "the Federal Government",
                "U.S. government": "U.S. Government"}.get(m)
    if rule == "HM.US":
        if re.match(r"\s*[A-Za-z]", after):
            return "U.S."
        if re.match(r"\s*[,.;)]", after) or after == "":
            return "United States"
        return None
    if rule == "HM.USNoun":
        return "the United States" if m == "the U.S." else None
    if rule == "HM.SectionRefs":
        return "Sec. " if re.match(r"\d+\.\d", after) else None
    if rule == "HM.LatinAbbreviations":
        return "that is" if before.endswith(", ") and after.startswith(",") else None
    if rule == "HM.Wordiness":
        low = m.lower()
        if low == "the majority of":
            nxt = after.split(None, 1)[0].lower() if after.strip() else ""
            return cap_like(m, "most of" if nxt in DETERMINERS else "most")
        simple = {"in order to": "to", "in excess of": "more than", "in the absence of": "without",
                  "with the exception of": "except", "in conjunction with": "with", "whether or not": "whether"}
        return cap_like(m, simple[low]) if low in simple else None
    if rule == "HM.UnitSpacing":
        mm = re.fullmatch(r"([\d.,]+)(in|mph|km)", m)
        return f"{mm.group(1)} {mm.group(2)}" if mm else None
    return None


def fix_file(path: Path, alerts: list, dry_run: bool) -> list:
    """Apply fixes to one file; return [(line, rule, old, new)]."""
    lines = path.read_text(encoding="utf-8").split("\n")
    by_line = defaultdict(list)
    for a in alerts:
        by_line[a["Line"]].append(a)

    changes = []
    for ln, items in by_line.items():
        text = lines[ln - 1]
        if TABLE_ROW.match(text):
            continue
        is_ref = bool(REF_LINE.search(text))
        applied = []  # spans already changed on this line, 0-based, end exclusive
        # Work right to left so earlier spans keep their positions.
        for a in sorted(items, key=lambda a: -a["Span"][0]):
            s, e = a["Span"][0] - 1, a["Span"][1]
            if any(not (e <= ps or s >= pe) for ps, pe in applied):
                continue
            if is_ref and a["Check"] != "HM.Underline":
                continue
            m, before, after = text[s:e], text[:s], text[e:]
            # Leave bracketed translations of titles alone: \[ ... \]
            if before.rfind("\\[") > before.rfind("\\]"):
                continue
            new = fix(a["Check"], m, before, after)
            if new is None:
                continue
            extra = 0
            if isinstance(new, tuple):
                new, extra = new
            if new == m and not extra:
                continue
            s -= extra
            changes.append((ln, a["Check"], text[s:e], new))
            text = text[:s] + new + text[e:]
            applied.append((s, e))
        lines[ln - 1] = text

    if changes and not dry_run:
        path.write_text("\n".join(lines), encoding="utf-8", newline="")
    return changes


def write_log(log_path: Path, results: dict, dry_run: bool) -> None:
    counts = defaultdict(int)
    for changes in results.values():
        for _, rule, _, _ in changes:
            counts[rule] += 1
    esc = lambda t: t.replace("|", "\\|")
    with log_path.open("w", encoding="utf-8") as out:
        out.write("# Vale fixes " + ("(dry run, files not changed)" if dry_run else "applied") + "\n\n")
        out.write("| Rule | Changes |\n|---|---|\n")
        for rule, n in sorted(counts.items(), key=lambda x: -x[1]):
            out.write(f"| {rule} | {n} |\n")
        for name in sorted(results):
            if not results[name]:
                continue
            out.write(f"\n## {name}\n\n| Line | Rule | Was | Now |\n|---|---|---|---|\n")
            for ln, rule, old, new in sorted(results[name]):
                out.write(f"| {ln} | {rule} | `{esc(old)}` | `{esc(new)}` |\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Vale and apply the mechanical HM fixes to Markdown files.")
    parser.add_argument("md_folder", type=Path, help="Folder of .md files (edited in place)")
    parser.add_argument("log", help="Log file (.md), or a folder for vale-fix-log.md")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help=f"Vale config (default: {DEFAULT_CONFIG})")
    parser.add_argument("--dry-run", action="store_true", help="Write the log without changing any files")
    args = parser.parse_args()

    folder = args.md_folder.resolve()
    md_files = sorted(folder.glob("*.md"))
    if not md_files:
        print(f"No .md files in: {folder}")
        sys.exit(1)
    log_path = resolve_output_file(args.log, "vale-fix-log.md", (".md",))
    if log_path.parent == folder:
        sys.exit("Put the log outside the md folder, or Vale will check it next time.")

    result = subprocess.run(
        [find_vale(), "--config", args.config, "--output", "JSON", *[str(p) for p in md_files]],
        capture_output=True, text=True, encoding="utf-8",
    )
    try:
        alerts = json.loads(result.stdout or "{}")
    except json.JSONDecodeError:
        sys.exit(f"Vale failed:\n{result.stderr.strip() or result.stdout.strip()}")

    if not args.dry_run:
        backup = folder.with_name(f"{folder.name}-before-vale-fix-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copytree(folder, backup)
        print(f"Backup: {backup}")

    results = {}
    for fname, file_alerts in alerts.items():
        path = Path(fname)
        results[path.name] = fix_file(path, file_alerts, args.dry_run)

    write_log(log_path, results, args.dry_run)
    total = sum(len(c) for c in results.values())
    print(f"{'Would make' if args.dry_run else 'Made'} {total} changes in {len(md_files)} files. Log: {log_path}")


if __name__ == "__main__":
    main()
