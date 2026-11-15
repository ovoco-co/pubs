"""Run the clean-template migration, every step in order, from the originals.

Usage:
    python run_pipeline.py <originals folder> [--out <run folder>] [--settings settings.json]

Always starts from the untouched author files, so every run is the originals
plus these steps, never the originals plus whatever happened last time. Each
step writes its own folder under the run folder (00-template, 01-strip, ...)
with a copy of the chapters as that step left them, and a check against the
originals runs after every step that touches chapters. A failed check stops
the run at that step and says what changed.

Steps (README.md has the detail):
  00 template       style architecture -> empty NIST-Template.docx
  01 strip          hand formatting out, tracked changes kept         XML
  02 remove toc     chapters lose their contents and lists            XML
  03 move           chapters into the template                        Word
  04 table text     data-table cells onto Table Text                  XML
  05 table style    one table style and options; widths; callouts     XML
  06 table notes    trailing note rows -> one table-footnote row      XML
  07 group rows     full-width group-label rows shaded and bold       XML
  08 decimal tabs   numbers aligned on the decimal point              XML
  09 chapter number each chapter's own number back on Heading 1       XML
  10 captions       caption check; figure descriptions -> Caption_Figure_2
  11 update fields  F9, after a guard against runaway REF bookmarks   Word
  12 cross-refs     typed Table/Figure/Section mentions -> live refs  XML+Word
  13 update fields  F9 again, so references show current numbers      Word
  14 three columns  optional: 3-column tables at fixed widths         XML
  15 accent bar     optional: colored bar by caption                  XML
"""
import argparse
import fnmatch
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
WORD = HERE.parent / 'word'
PY = sys.executable
sys.path.insert(0, str(HERE))
from compare_content import compare  # noqa: E402

# Captions hold en dashes and other non-ASCII text: the step scripts write
# UTF-8 (read back as UTF-8 below), and this console prints it.
# Unbuffered, so a step that is ended on timeout has still passed on what it
# printed; this console line-buffered, so progress shows while it runs.
CHILD_ENV = {**os.environ, 'PYTHONIOENCODING': 'utf-8', 'PYTHONUNBUFFERED': '1'}
sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)


def run(cmd, log, timeout=None):
    """Run one command, log its output, stop the pipeline if it fails. A
    command that drives Word gets a timeout: hidden Word can wait forever on
    a dialog nobody sees, so on timeout the automation Word is ended and the
    run stops with what the step had printed (the last file named is the
    one it finished; the next one is where it stuck)."""
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=HERE,
                             env=CHILD_ENV, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        end_automation_word()
        partial = e.stdout.decode('utf-8', 'replace') if isinstance(e.stdout, bytes) else (e.stdout or '')
        log.write(f'$ {" ".join(str(c) for c in cmd)}\n{partial}\nTIMED OUT after {timeout} s\n')
        sys.exit(f'TIMED OUT after {timeout / 60:.0f} min, hidden Word ended: '
                 f'{" ".join(str(c) for c in cmd)}\n{partial[-2000:]}')
    log.write(f'$ {" ".join(str(c) for c in cmd)}\n{out.stdout}{out.stderr}\n')
    if out.returncode:
        sys.exit(f'FAILED: {" ".join(str(c) for c in cmd)}\n{out.stdout[-2000:]}{out.stderr[-2000:]}')
    return out.stdout


def ps(script, *args):
    return ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(WORD / script), *args]


def chapters(folder):
    return sorted(p for p in folder.glob('*.docx') if not p.name.startswith('~$') and 'Template' not in p.name)


def copy_step(prev, cur):
    cur.mkdir(parents=True, exist_ok=True)
    for f in chapters(prev):
        shutil.copyfile(f, cur / f.name)


def check(originals, folder, step, keep_toc=None, **opts):
    """Compare each chapter in folder with the same chapter in originals
    (the source files, or the previous step for a field update). keep_toc,
    once step 02 has run: the originals' tables of contents are set aside,
    except in the files matching these patterns, which keep theirs."""
    failed = []
    for orig in chapters(originals):
        new = folder / orig.name
        if not new.exists():
            failed.append(f'{orig.name}: missing')
            continue
        drop = keep_toc is not None and not any(fnmatch.fnmatch(orig.name, k) for k in keep_toc)
        problems, _ = compare(orig, new, drop_toc=drop, **opts)
        if problems:
            failed.append(f'{orig.name}: ' + '; '.join(problems)[:600])
    if failed:
        sys.exit(f'CHECK FAILED after step {step}:\n  ' + '\n  '.join(failed))
    what = ('the previous step, field results aside' if opts.get('field_update')
            else 'the previous step' if originals.name[:2].isdigit() else 'the originals')
    print(f'  check: all {len(chapters(originals))} chapters match {what}')


def chapter_number(original):
    """The chapter number the original document carries, as
    Set-NISTChapterNumbering reads it (HMList level 1 start)."""
    out = subprocess.run(ps('Set-NISTChapterNumbering.ps1', '-File', str(original), '-ListOnly', '-Force'),
                         capture_output=True, text=True, encoding='utf-8', errors='replace')
    m = re.search(r'Chapter number (\d+) read', out.stdout)
    return m.group(1) if m else None


def end_automation_word():
    """End the Word started by a step (command line /Automation). The run
    refuses to start while any Word is open, so no one else's Word has it."""
    subprocess.run(['powershell', '-NoProfile', '-Command',
                    "Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE'\" | "
                    "Where-Object { $_.CommandLine -match '/Automation' } | "
                    "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"], capture_output=True)


def word_running():
    out = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE'], capture_output=True, text=True)
    return 'WINWORD.EXE' in out.stdout.upper()


def main():
    ap = argparse.ArgumentParser(description='Run the clean-template migration from the originals.')
    ap.add_argument('originals', type=Path)
    ap.add_argument('--out', type=Path, help='run folder (default: <originals>/run)')
    ap.add_argument('--settings', type=Path, default=HERE / 'settings.json')
    a = ap.parse_args()
    s = json.loads(a.settings.read_text(encoding='utf-8'))
    source = a.originals.resolve()
    out = (a.out or source / 'run').resolve()
    wait = s['word_step_timeout_minutes'] * 60
    included = [f for f in chapters(source) if not any(fnmatch.fnmatch(f.name, x) for x in s['exclude'])]
    if not included:
        sys.exit(f'No chapters in {source}')
    locked = [f.name for f in included if f.with_name('~$' + f.name[2:]).exists()]
    if locked:
        sys.exit('Open in Word, close first: ' + ', '.join(locked))
    if word_running():
        sys.exit('Word is running. Close it first: steps 03, 11, 12 and 13 drive Word and would attach to the open copy.')
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    # Every step reads this snapshot of the included originals, never the
    # source folder itself.
    orig = out / 'originals'
    orig.mkdir()
    for f in included:
        shutil.copyfile(f, orig / f.name)
    for f in sorted(set(chapters(source)) - set(included)):
        print(f'  excluded by settings: {f.name}')
    log = (out / 'run-log.txt').open('w', encoding='utf-8')
    started = time.time()
    steps = []

    def step(n, name):
        d = out / f'{n:02d}-{name}'
        print(f'{n:02d} {name}')
        steps.append(d.name)
        return d

    # 00 template
    d = step(0, 'template')
    d.mkdir()
    src_tpl = d / 'template-source.docx'
    shutil.copyfile(s['template_source'], src_tpl)
    run(ps('Set-NISTStyleArchitecture.ps1', '-File', str(src_tpl), '-Force'), log)
    template = d / 'NIST-Template.docx'
    run([PY, 'make_empty_template.py', str(src_tpl), str(template)], log)

    # 01 strip
    d = step(1, 'strip')
    maps = [f'{k}={v}' for k, v in s['style_map'].items() if not k.startswith('_')]
    run([PY, 'strip_formatting.py', str(orig), str(d), '--template', str(template),
         *(['--map', *maps] if maps else [])], log)
    check(orig, d, '01')

    # 02 remove toc: chapters do not get a table of contents or lists of
    # tables/figures; front matter files keep theirs
    prev, d = d, step(2, 'remove-toc')
    copy_step(prev, d)
    kt = s['keep_toc']
    for line in run([PY, 'remove_toc.py', str(d), '--keep', *kt], log).splitlines():
        if 'removed' in line:
            print('  ' + line)
    check(orig, d, '02', keep_toc=kt)

    # 03 move
    prev, d = d, step(3, 'move')
    run([PY, 'move_into_template.py', str(prev), str(template), str(d)], log, wait)
    check(orig, d, '03', keep_toc=kt)

    # 04 table text
    prev, d = d, step(4, 'table-text')
    copy_step(prev, d)
    run([PY, 'set_table_text.py', str(d)], log)
    check(orig, d, '04', keep_toc=kt)

    # 05 table style
    prev, d = d, step(5, 'table-style')
    copy_step(prev, d)
    run([PY, 'apply_table_style.py', str(d), '--style', s['table_style']], log)
    check(orig, d, '05', keep_toc=kt)

    # 06 table notes (reshapes the note rows, so the shape test is skipped here)
    prev, d = d, step(6, 'table-notes')
    copy_step(prev, d)
    run(ps('Merge-TableFootnotes.ps1', '-Folder', str(d), '-Force', '-NoBackup'), log)
    check(orig, d, '06', keep_toc=kt, allow_table_shape=True)

    # 07 group rows
    prev, d = d, step(7, 'group-rows')
    copy_step(prev, d)
    run(ps('Set-TableBannerRows.ps1', '-Folder', str(d), '-Fill', s['group_row_fill'], '-Force', '-NoBackup'), log)
    check(orig, d, '07', keep_toc=kt, allow_table_shape=True)

    # 08 decimal tabs
    prev, d = d, step(8, 'decimal-tabs')
    copy_step(prev, d)
    run(ps('Set-TableDecimalTabs.ps1', '-Folder', str(d), '-Force', '-NoBackup'), log)
    check(orig, d, '08', keep_toc=kt, allow_table_shape=True)

    # 09 chapter numbers: the move put every chapter on the template's
    # numbering, which starts at 1; each gets back the number its original had
    prev, d = d, step(9, 'chapter-numbers')
    copy_step(prev, d)
    for f in chapters(d):
        n = chapter_number(orig / f.name)
        if not n:
            sys.exit(f'No chapter number in the original {f.name}')
        run(ps('Set-NISTChapterNumbering.ps1', '-File', str(f), '-ChapterNumber', n, '-Force', '-NoBackup'), log)
        print(f'  {f.name}: chapter {n}')
    check(orig, d, '09', keep_toc=kt, allow_table_shape=True)

    # 10 captions
    prev, d = d, step(10, 'captions')
    copy_step(prev, d)
    for line in run([PY, 'caption_check.py', str(d)], log).splitlines():
        if line.startswith('    '):
            print('  ' + line)
    check(orig, d, '10', keep_toc=kt, allow_table_shape=True)

    # 11 update fields
    prev, d = d, step(11, 'update-fields')
    copy_step(prev, d)
    run([PY, 'update_fields.py', str(d)], log, wait)
    check(prev, d, '11', allow_table_shape=True, field_update=True)

    # 12 cross-references: tables, figures (existing script), sections (Word)
    prev, d = d, step(12, 'cross-refs')
    copy_step(prev, d)
    for label in ('Table', 'Figure'):
        run(ps('Set-TableCrossReferences.ps1', '-Folder', str(d), '-Label', label, '-Force', '-NoBackup'), log)
    for line in run([PY, 'section_xrefs.py', str(d)], log, wait).splitlines():
        print('  ' + line)
    # against step 11: from the field update on, field results are current,
    # and this step only turns typed text into fields showing that same text
    check(prev, d, '12', allow_table_shape=True, new_fields=True)

    # 13 update fields again, so the new references show current numbers
    prev, d = d, step(13, 'update-fields')
    copy_step(prev, d)
    run([PY, 'update_fields.py', str(d)], log, wait)
    check(prev, d, '13', allow_table_shape=True, field_update=True)

    # 14 optional: three-column widths
    if s['optional_steps'].get('three_column_widths'):
        prev, d = d, step(14, 'three-columns')
        copy_step(prev, d)
        w1, w2, w3 = s['three_column_widths_inches']
        run(ps('Set-NISTThreeColumnTableWidths.ps1', '-Folder', str(d), '-FirstColumnInches', str(w1),
               '-SecondColumnInches', str(w2), '-ThirdColumnInches', str(w3), '-Force', '-NoBackup'), log)
        check(orig, d, '14', keep_toc=kt, allow_table_shape=True, new_fields=True)

    # 15 optional: accent bar
    if s['optional_steps'].get('accent_bar'):
        prev, d = d, step(15, 'accent-bar')
        copy_step(prev, d)
        ab = s['accent_bar']
        rules = ['-ColorRules', *ab['color_rules']] if ab['color_rules'] else []
        run(ps('Set-TableAccentBar.ps1', '-Folder', str(d), '-Color', ab['color'], *rules, '-Force', '-NoBackup'), log)
        check(orig, d, '15', keep_toc=kt, allow_table_shape=True, new_fields=True)

    final = out / 'final'
    copy_step(d, final)
    shutil.copyfile(template, final / 'NIST-Template.docx')
    log.close()
    (out / 'run-report.md').write_text(
        f'# Clean-template run\n\n- Originals: `{orig}`\n- Settings: `{a.settings}`\n'
        f'- Steps: {", ".join(steps)}\n- Finished in {time.time() - started:.0f} s; every step checked against the originals.\n'
        f'- Finished chapters: `{final}`\n- Step 01 report: `{out / "01-strip" / "strip-report.md"}`\n'
        f'- Full command output: `{out / "run-log.txt"}`\n', encoding='utf-8')
    print(f'Done. Finished chapters: {final}')


if __name__ == '__main__':
    main()
