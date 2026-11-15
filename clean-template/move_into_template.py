"""Step 2 of the clean-template migration: put each stripped chapter into a
fresh copy of the template, using Word.

Usage: python move_into_template.py <stripped folder> <template.docx> <out folder>

For each .docx in the stripped folder: copies the template to
<out folder>/<name>.docx, opens it, inserts the chapter at the top of the
template's empty body (Word's Insert > Text from File), turns track changes
on, and saves. Where a style name exists in both, the template's definition
wins. Tracked changes and comments arrive as they were, under their original
authors: nothing is inserted while tracking is on, so no new revision is made.

Each chapter's own document properties (title, custom properties) then
replace the template's, since the chapter banner and running heads print the
TITLE property.

The template's own empty paragraph stays at the end. Deleting it would merge
its paragraph mark into the chapter's last paragraph and restyle it.

Needs Word, and Word must not already be running: an automation session can
attach to a copy someone has open.
"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import win32com.client as win32


def word_running():
    out = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE'], capture_output=True, text=True)
    return 'WINWORD.EXE' in out.stdout.upper()


def main():
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    src, template, out = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
    files = sorted(p for p in src.glob('*.docx') if not p.name.startswith('~$'))
    if not files:
        sys.exit(f'No .docx files in {src}')
    if word_running():
        sys.exit('Word is running. Close it first: this script would attach to the open copy.')
    out.mkdir(parents=True, exist_ok=True)

    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for f in files:
            dst = out / f.name
            shutil.copyfile(template, dst)
            doc = word.Documents.Open(str(dst.resolve()), False, False, False)
            try:
                doc.TrackRevisions = False
                doc.Range(0, 0).InsertFile(str(f.resolve()))
                doc.TrackRevisions = True
                doc.Save()
                print(f'{f.name}: moved into the template')
            finally:
                doc.Close(0)
    finally:
        word.Quit(0)
    for f in files:
        copy_properties(f, out / f.name)


def copy_properties(src, dst):
    """The chapter's own document properties (title, subject, custom
    properties) replace the template's. Insert-from-file brings the content
    but not the properties, and the banner and running heads print the TITLE
    property: without this every chapter is titled like the template."""
    zs = zipfile.ZipFile(src)
    zd = zipfile.ZipFile(dst)
    carry = {p: zs.read(p) for p in ('docProps/core.xml', 'docProps/custom.xml')
             if p in zs.namelist() and p in zd.namelist()}
    tmp = dst.with_suffix('.tmp')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zd.infolist():
            zout.writestr(info, carry.get(info.filename, zd.read(info.filename)))
    zd.close()
    tmp.replace(dst)


if __name__ == '__main__':
    main()
