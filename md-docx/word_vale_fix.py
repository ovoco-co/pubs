"""Apply the mechanical HM Vale fixes directly in Word documents, as tracked changes.

Usage:
    python word_vale_fix.py <file.docx | folder> [...] --log <log_file.md|folder> --author "Name" [--dry-run]

Opens each .docx in Word (Word must be installed), turns on Track Changes, and makes the
same fixes as vale_fix.py, plus removing parentheses around a percentage that is part of
the sentence ("approximately (66%) were" -> "approximately 66 % were"). Every edit shows as
a tracked change by --author. Existing tracked changes and comments are left as they are.

Checked: body paragraphs and footnotes, using the text as it reads with all tracked changes
shown as accepted. Skipped: tables, indented quotes, reference-list entries, the table of
contents, and captions. The files are saved in place; --dry-run only writes the log.
"""
import argparse
import json
import re
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import win32com.client as win32

from output_paths import resolve_output_file
from vale_fix import DEFAULT_CONFIG, REF_LINE, find_vale, fix

# Tracked text that disappears when changes are accepted: deletions, the old spot of a moved
# passage ("moved from"), and conflict deletions.
WD_REMOVED_TYPES = {2, 14, 21}
WD_STORIES = (1, 2, 3)  # main text, footnotes, endnotes
SKIP_STYLE = re.compile(r"caption|toc|table of|quote|bibliograph", re.IGNORECASE)
PCT_PAREN = re.compile(r"\(\s*([\d.,]+\s?%)\s*\)")
PAREN_BAD_BEFORE = re.compile(
    r"(?:\b(?:of|approximately|about|nearly|roughly|almost|and|with|than|to|was|were|is|are|by|only|just)"
    r"|[;,:]|[.!?]|^)\s*$",
    re.IGNORECASE,
)
CONTEXT = 20


class Para:
    """One paragraph: its raw text (tracked deletions included) and its final text."""

    def __init__(self, para, raw, deleted_texts):
        self.para = para
        self.raw = raw
        deleted = [False] * len(raw)
        cursor = 0
        for text in deleted_texts:
            i = raw.find(text, cursor)
            if i < 0:
                continue
            for k in range(i, i + len(text)):
                deleted[k] = True
            cursor = i + len(text)
        self.deleted = deleted
        self.final_to_raw = [i for i, d in enumerate(deleted) if not d]
        self.final = "".join(raw[i] for i in self.final_to_raw).rstrip("\r\x07")


def collect_paragraphs(doc):
    tables = [(t.Range.Start, t.Range.End) for t in doc.Tables]
    paras = []
    for story_type in WD_STORIES:
        try:
            story = doc.StoryRanges(story_type)
        except Exception:
            continue
        deletions = [(r.Range.Start, r.Range.End, r.Range.Text)
                     for r in story.Revisions if r.Type in WD_REMOVED_TYPES]
        for p in story.Paragraphs:
            rng = p.Range
            s, e = rng.Start, rng.End
            if story_type == 1 and any(ts <= s < te for ts, te in tables):
                continue
            try:
                style = p.Style.NameLocal
            except Exception:
                style = ""
            if SKIP_STYLE.search(style):
                continue
            if p.LeftIndent > 0 and rng.ListFormat.ListType == 0:
                continue  # indented block: interview quote
            raw = rng.Text
            dels = sorted((ds, dt) for ds, de, dt in deletions if s <= ds < e)
            item = Para(p, raw, [dt for _, dt in dels])
            if item.final.strip():
                paras.append(item)
    return paras


def vale_alerts(paras, config):
    """Run Vale on the final text, one paragraph per Markdown paragraph."""
    with tempfile.TemporaryDirectory() as tmp:
        md = Path(tmp) / "doc.md"
        md.write_text("\n\n".join(p.final for p in paras) + "\n", encoding="utf-8")
        result = subprocess.run([find_vale(), "--config", config, "--output", "JSON", str(md)],
                                capture_output=True, text=True, encoding="utf-8")
        try:
            data = json.loads(result.stdout or "{}")
        except json.JSONDecodeError:
            sys.exit(f"Vale failed:\n{result.stderr.strip() or result.stdout.strip()}")
    by_para = defaultdict(list)
    for alerts in data.values():
        for a in alerts:
            by_para[(a["Line"] - 1) // 2].append(a)
    return by_para


def plan_edits(item, alerts):
    """Return [(start, end, old, new, rule)] in final-text positions, non-overlapping."""
    text = item.final
    edits = []
    for m in PCT_PAREN.finditer(text):
        if PAREN_BAD_BEFORE.search(text[:m.start()]):
            edits.append((m.start(), m.end(), m.group(0), re.sub(r"\s*%$", " %", m.group(1)), "Parentheses"))
    if not REF_LINE.search(text):
        for a in sorted(alerts, key=lambda a: -a["Span"][0]):
            s, e = a["Span"][0] - 1, a["Span"][1]
            new = fix(a["Check"], text[s:e], text[:s], text[e:])
            if new is None:
                continue
            if isinstance(new, tuple):
                new, extra = new
                s -= extra
            if text[s:e] == new:
                continue
            edits.append((s, e, text[s:e], new, a["Check"]))
    edits.sort(key=lambda x: x[0])
    kept = []
    for ed in edits:
        if kept and ed[0] < kept[-1][1]:
            continue  # overlaps an earlier edit (the parentheses fix already spaces the %)
        kept.append(ed)
    return kept


def apply_edit(doc, item, edit, limit_raw):
    """Make one tracked edit. Returns None on success, or a reason it was skipped."""
    s, e, old, new, _ = edit
    rs, re_ = item.final_to_raw[s], item.final_to_raw[e - 1] + 1
    if any(item.deleted[rs:re_]):
        return "crosses another tracked change"
    raw = item.raw
    before = raw[max(0, rs - CONTEXT):rs]
    after = raw[re_:min(limit_raw, re_ + CONTEXT)]
    before = re.split(r"[\x00-\x1f]", before)[-1]
    after = re.split(r"[\x00-\x1f]", after)[0]
    token = raw[rs:re_]
    search = before + token + after
    occurrence = raw[:rs - len(before)].count(search) if search else 0
    if not search or len(search) > 250:
        return "no usable context"
    rng = item.para.Range
    find = rng.Find
    find.ClearFormatting()
    for _ in range(occurrence + 1):
        if not find.Execute(search.replace("^", "^^"), True, False, False, False, False, True, 0):
            return "context not found"
    inner = rng.Duplicate  # same story as the paragraph (main text or footnotes)
    skip = before.count(token)
    f2 = inner.Find
    for _ in range(skip + 1):
        if not f2.Execute(token.replace("^", "^^"), True, False, False, False, False, True, 0):
            return "token not found"
    para = item.para.Range
    if (inner.Text != token or inner.StoryType != para.StoryType
            or inner.Start < para.Start or inner.End > para.End):
        return "found text did not match"
    if any(r.Type in WD_REMOVED_TYPES for r in inner.Revisions):
        return "inside deleted or moved text"
    inner.Text = new
    return None


def process(path, author, config, dry_run, log):
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    old_name, old_initials = word.UserName, word.UserInitials
    old_local = word.Options.UseLocalUserInfo
    try:
        word.UserName = author
        word.UserInitials = "".join(w[0] for w in author.split() if w).upper()
        word.Options.UseLocalUserInfo = True
        doc = word.Documents.Open(str(path), False, dry_run, False)
        doc.TrackRevisions = True
        paras = collect_paragraphs(doc)
        alerts = vale_alerts(paras, config)
        made = skipped = 0
        # Work from the end of the document backward, and right to left in each paragraph,
        # so positions and context not yet edited stay valid.
        for idx in range(len(paras) - 1, -1, -1):
            item = paras[idx]
            edits = plan_edits(item, alerts.get(idx, []))
            limit = len(item.raw)
            for edit in reversed(edits):
                reason = "dry run" if dry_run else apply_edit(doc, item, edit, limit)
                limit = item.final_to_raw[edit[0]]
                ok = reason in (None, "dry run")
                made += ok
                skipped += not ok
                log.append((path.name, idx + 1, edit[4], edit[2], edit[3], reason or "", item.final))
        if not dry_run:
            doc.Save()
        doc.Close(False)
        print(f"{path.name}: {made} edits{' planned' if dry_run else ''}, {skipped} skipped")
    finally:
        word.UserName, word.UserInitials = old_name, old_initials
        word.Options.UseLocalUserInfo = old_local
        word.Quit()


def write_log(log_path, log, dry_run):
    esc = lambda t: t.replace("|", "\\|").replace("\n", " ")
    counts = defaultdict(int)
    for row in log:
        if row[5] in ("", "dry run"):
            counts[row[2]] += 1
    with log_path.open("w", encoding="utf-8") as out:
        out.write("# Word tracked-change fixes " + ("(dry run)" if dry_run else "applied") + "\n\n")
        out.write("| Rule | Edits |\n|---|---|\n")
        for rule, n in sorted(counts.items(), key=lambda x: -x[1]):
            out.write(f"| {rule} | {n} |\n")
        skipped = [r for r in log if r[5] not in ("", "dry run")]
        if skipped:
            out.write("\n## Skipped (make these by hand)\n\n| File | Paragraph | Rule | Was | Now | Why |\n|---|---|---|---|---|---|\n")
            for f, p, rule, old, new, why, _ in skipped:
                out.write(f"| {f} | {p} | {rule} | `{esc(old)}` | `{esc(new)}` | {why} |\n")
        current = None
        for f, p, rule, old, new, why, text in sorted(log, key=lambda r: (r[0], r[1])):
            if f != current:
                out.write(f"\n## {f}\n\n| Paragraph | Rule | Was | Now |\n|---|---|---|---|\n")
                current = f
            out.write(f"| {p} | {rule} | `{esc(old)}` | `{esc(new)}` |\n")


def main():
    parser = argparse.ArgumentParser(description="Apply HM Vale fixes in Word as tracked changes.")
    parser.add_argument("inputs", type=Path, nargs="+", help=".docx files or folders of them")
    parser.add_argument("--log", required=True, help="Log file (.md), or a folder for word-vale-fix-log.md")
    parser.add_argument("--author", required=True, help="Name shown on the tracked changes")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help=f"Vale config (default: {DEFAULT_CONFIG})")
    parser.add_argument("--dry-run", action="store_true", help="Write the log without changing any files")
    args = parser.parse_args()

    docs = []
    for source in (p.resolve() for p in args.inputs):
        docs += ([source] if source.is_file() else
                 sorted(p for p in source.glob("*.docx") if not p.name.startswith("~$")))
    if not docs:
        sys.exit("No .docx files found.")
    log_path = resolve_output_file(args.log, "word-vale-fix-log.md", (".md",))

    log = []
    for path in docs:
        process(path, args.author, args.config, args.dry_run, log)
    write_log(log_path, log, args.dry_run)
    print(f"Log: {log_path}")


if __name__ == "__main__":
    main()
