# Word Template Toolkit

PowerShell scripts for working with Word templates at the OOXML layer. Reads files directly as zip archives of XML using `System.IO.Compression` and `System.Xml`, so Word does not need to be installed. Works on Windows, WSL, Linux, and macOS with PowerShell 5.1+.

Most scripts are client-agnostic and work on any Word document. The `Set-NIST*` / `Audit-NIST*` scripts are NIST-specific: they are thin wrappers that supply the NIST palette, style list, and margins and delegate to the generic engines in this same folder. The `ncstar/` and `nwirp/` subfolders hold the report-assembly tools for those two NIST document sets.

## Running the scripts

`cd` into this folder and run each script with `.\`. Do NOT copy the scripts into the document folder - the `Set-NIST*` wrappers find their generic engine and the `OoXml/` module by their own location, so they must run from here. Point `-Folder` at wherever the `.docx` files live.

```powershell
# cd into the toolkit folder (WSL path from Windows; adjust the distro name if yours differs)
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word

$docs = "$env:USERPROFILE\Documents\260726"
.\Set-PageLayout.ps1 -Folder $docs -InnerInches 1.25 -OuterInches 0.75
```

(From another directory without `cd`, call by full path with the `&` operator instead: `& "\\wsl$\...\word\Set-PageLayout.ps1" -Folder $docs`.)

- Each script lists its targets and asks `Proceed? (y/N)`. Type `y`, or add `-Force` to skip.
- Swap `-Folder` for `-File "one.docx"` to test on a single file first.
- Scripts that modify files write a `.bak` backup before overwriting.

## NIST template build

Apply the NIST template to a folder of chapter files. Order matters: step 1 installs the palette that step 2's styles reference.

```powershell
cd \\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word
$docs = "$env:USERPROFILE\Documents\260726"

.\Set-NISTTheme.ps1               -Folder $docs
.\Set-NISTStyleArchitecture.ps1   -Folder $docs
.\Set-PageLayout.ps1              -Folder $docs -InnerInches 1.25 -OuterInches 0.75
.\Set-TableLook.ps1               -Folder $docs
.\Set-DocTitle.ps1                -Folder $docs   # no -Title, so it prompts per file

# Header/footer: chapters get a first-page footer; front matter stays blank.
# Run the front-matter file with -FrontMatter, the chapters without it.
.\Set-NISTHeaderFooterContent.ps1 -File "$docs\NCSTAR_4-3_Chapter-00_Front-Matter.docx" -FrontMatter
Get-ChildItem "$docs\*.docx" -Exclude '*Front-Matter*' | ForEach-Object {
    .\Set-NISTHeaderFooterContent.ps1 -File $_.FullName -Force
}

.\Remove-Shading.ps1              -Folder $docs   # drop stray text/paragraph backgrounds
```

**Entering the title.** Leave `-Title` off and `Set-DocTitle.ps1` walks each file one at a time, showing `[n/total] filename`, and prompts you for that file's title:

```
[1/7] NCSTAR_4-3_Chapter-00_Front-Matter.docx
Title (blank to skip):
```

Type the title and press Enter; it sets it and moves to the next file. Leaving a title blank skips that file. This is the usual mode when each chapter carries its own title. Pass `-Title "..."` to apply one title to every file instead (with a single y/N confirmation), or `-File "$docs\Chapter-00.docx"` to target just one.

Notes:

- **NIST margins.** `Set-PageLayout.ps1` defaults to 1.0" all around; `-InnerInches 1.25 -OuterInches 0.75` are the NIST mirror-margin values.
- **Style colors are fixed hex**, not theme-linked, so they survive Word's Import Styles into a document with a different theme. Changing Design > Colors will not recolor them.
- `Reapply-Styles.ps1` (strip direct character formatting) is a separate post-authoring cleanup step, not part of this build.

## Scripts

### NIST template (this folder)

| Script | What it does |
|--------|-------------|
| `Set-NISTTheme.ps1` | Thin wrapper: passes the NIST NCSTAR palette to `Set-ThemeColors.ps1`. |
| `Set-NISTStyleArchitecture.ps1` | Imports `OoXml/OoXml.psd1` and applies NIST style config. Rebases Heading1-9, BodyText, Caption, List Bullet/Number onto `xx*`, and AppendixHead 1-4 and Front_matter_head 1-4 onto Heading 1-4 (numbering suppressed) so appendix and front-matter titles match the headings. Deletes junk styles. Sets style colors as fixed hex (headings Resolution Blue `#00257A`; Heading 5 neutral gray `#404040`). Sets the bullet-glyph ladder, the List Bullet indent (left 0.25" / hanging 0.2" / right 0), and the HM-Data table's first-column shading (`#F2F2F2`). Manages the Sidebar / Sidebar Bullet callout paragraph styles and installs a Sidebar Table style (whole-table amber `#FFF2CC` fill). |
| `Set-NISTHeaderFooterContent.ps1` | Installs even/odd/first header and footer parts with STYLEREF chapter number/title and TITLE+PAGE fields. Chapter first pages get a TITLE+PAGE footer; pass `-FrontMatter` to leave the front-matter first-page footer blank. |
| `Install-NISTThemeColors.ps1` | Thin wrapper: installs the NIST palette into the user's Word Theme Colors folder. Per-user; does not touch documents. |
| `Audit-NISTStyleSet.ps1` | Thin wrapper: passes the required/optional NIST style ids to `Audit-StyleSet.ps1`. Pre-assembly gate for `ncstar/Build-NISTReport.ps1`. |

### Inspection (read-only)

| Script | What it does |
|--------|-------------|
| `Get-WordStyles.ps1` | Lists every style in a Word file as Markdown. One section per style with Identity, Font, Paragraph, and Table attributes. Optional HTML view. |
| `Get-WordTableStyles.ps1` | Table styles only, with full conditional formatting blocks (header row, banded rows, first/last column, corner cells). |
| `Get-WordParts.ps1` | Lists Building Blocks (Quick Parts, AutoText, cover pages) from the glossary with metadata and content preview. |
| `Compare-WordTemplates.ps1` | Compares every Word file in a folder for drift. Produces CSV + HTML showing where they disagree on styles, headers, footers, building blocks, content controls, fields, bookmarks, or custom XML parts. |
| `Audit-WordTheme.ps1` | Reports the theme palette and lists every place a style or run sets an explicit hex color instead of a theme reference. |

### Modification

| Script | What it does |
|--------|-------------|
| `Expand-WordPackage.ps1` | Unpacks a Word file into a folder of XML files for hand editing. Creates a backup first. |
| `Compress-WordPackage.ps1` | Repacks an unpacked folder back into a valid Word file. Validates the result before overwriting. |
| `Reapply-Styles.ps1` | Strips run-level direct character formatting so paragraph styles render cleanly. Preserves italic, superscript/subscript, and OLE/MathType equations. |
| `Remove-Shading.ps1` | Removes locally applied (direct) shading from runs and paragraphs, so text and bullets show no background. Keeps style-defined shading (styles.xml is never touched) and table-cell fills (unless `-IncludeTables`). |
| `Repair-Typography.ps1` | Converts straight quotes to curly (typographic) quotes, deciding open/close by context across run and paragraph boundaries. Only `<w:t>` text is changed - field codes and markup stay straight. Idempotent. `-Ellipses` also converts `...` to a single ellipsis. |
| `Set-DocTitle.ps1` | Sets `<dc:title>` in docProps/core.xml. Works for any title string in any Word document. |
| `Merge-WordDocuments.ps1` | Combines several Word files into one with a section break between each. Uses the first file as the host (theme, styles, headers), renumbers rIds/bookmarks/footnotes, copies and renames media, accepts revisions, and strips comments. Originals untouched. |
| `Set-PageLayout.ps1` | Sets page margins with mirror margins. Configurable via -InnerInches, -OuterInches, -TopInches, -BottomInches, -HeaderInches, -FooterInches. |
| `Set-TableLook.ps1` | Normalizes `<w:tblLook>` to formal report defaults (Header Row + Repeat Header + Banding) and adds `<w:tblHeader/>` to first rows. Tables using the Sidebar Table style are instead floated as right-margin callout boxes (text wraps around, 3" wide, no repeat header) since the float can't live in a table style. |
| `Set-ThemeColors.ps1` | Replaces the `<a:clrScheme>` in theme1.xml with a caller-supplied color scheme XML block. Validates all 12 theme slots are present. |
| `Install-ThemeColors.ps1` | Installs a theme color scheme file into the user's Word Theme Colors folder so Design > Colors displays it by name. |
| `Audit-StyleSet.ps1` | Generic style audit engine. Checks that caller-supplied required style ids exist in styles.xml. Returns non-zero on drift. JSON or table output. |

### Removing stray shading

`Remove-Shading.ps1` targets shading applied as direct formatting - the "Pattern: Clear (White)" white box (and any other pale fill) that text and bullets pick up from pasting or template imports. Shading that comes from a style belongs and is kept.

```powershell
# Whole folder (backs up each file to <name>.shading.bak)
.\Remove-Shading.ps1 -Folder $docs

# One file
.\Remove-Shading.ps1 -File "$docs\Chapter-02.docx"

# Also strip direct table-cell fills
.\Remove-Shading.ps1 -Folder $docs -IncludeTables
```

### Fixing quotes and typography

`Repair-Typography.ps1` converts straight quotes (`"` `'`) to curly typographic quotes, choosing opening vs closing by context - tracked across run and paragraph boundaries so a quote that starts a run still opens correctly. Apostrophes (`it's`, `teachers'`, `the '90s`) resolve to the right form. Only `<w:t>` text is touched, so field codes (like `STYLEREF "Heading 1"`) stay straight. Already-curly quotes are left alone, so it is safe to re-run.

```powershell
# Whole folder (backs up each file to <name>.typography.bak)
.\Repair-Typography.ps1 -Folder $docs

# One file
.\Repair-Typography.ps1 -File "$docs\Chapter-02.docx"

# Also convert ... to a single ellipsis
.\Repair-Typography.ps1 -Folder $docs -Ellipses
```

### Sidebar callout boxes

A sidebar is a shaded box that floats at the right margin with the body text wrapping around it. It is two pieces because the float cannot live in a table style:

1. In Word, put the sidebar content in a table and apply the **Sidebar Table** style (Table Design gallery). Amber `#FFF2CC` fill, thin borders, sans text, 3" wide.
2. Run the pipeline (`Set-NISTStyleArchitecture.ps1` then `Set-TableLook.ps1`). `Set-TableLook.ps1` stamps the right-margin float + text wrapping onto every Sidebar-Table-styled table and forces the 3" width; the box is the same for every sidebar and only its height changes with the content.

For sidebar *text* (not in a table), the **Sidebar** and **Sidebar Bullet** paragraph styles are managed by `Set-NISTStyleArchitecture.ps1` (sans 9pt, framed).

### Combining files into one document

`Merge-WordDocuments.ps1` combines several Word files into a single document with a section break between each. The first file is the host: its theme, styles, numbering, and header/footer parts apply to the whole output. Each following file contributes body content; its own page geometry is kept as its section, but headers/footers are unified to the host's. Source files are never modified.

```powershell
# Every .docx in the folder, in file-name order, -> <folder>\Combined.docx
.\Merge-WordDocuments.ps1 -Folder $docs

# Choose the output path
.\Merge-WordDocuments.ps1 -Folder $docs -Output "$docs\Full-Draft.docx" -Force

# Explicit order / subset (first file is the host)
.\Merge-WordDocuments.ps1 -Files a.docx,b.docx,c.docx -Output out.docx

# Start each document on an odd (right-hand) page
.\Merge-WordDocuments.ps1 -Folder $docs -SectionBreak oddPage
```

It renumbers relationship, bookmark, and footnote/endnote ids, copies and renames each file's media, accepts tracked changes, and strips comments. Open the result and press **Ctrl+A** then **F9** to refresh the table of contents and other fields. For a full NIST deliverable with roman/decimal page numbering and a style-audit gate, use `ncstar/Build-NISTReport.ps1` instead.

### Reapply styles (Word macro)

After Word's Import Styles, some paragraphs keep a stale look until you re-click the style, even with no direct formatting to clear. A file edit cannot fix this - the file is already correct and it is Word's rendering that is stale. `ReapplyAllStyles.bas` re-applies each paragraph's own style across the whole document, the same as clicking the style, but on every paragraph.

To use: in Word press **Alt+F11**, Insert > Module, paste the contents of `ReapplyAllStyles.bas`, press **F5**. It does not clear direct formatting (neither does clicking a style); use `Remove-Shading.ps1` or `Reapply-Styles.ps1` for that.

### Report builds

| Folder | What it holds |
|--------|---------------|
| `ncstar/` | NCSTAR report assembly and front-matter validation (`Build-NISTReport.ps1`, `Test-NISTFrontMatter.ps1`). See `ncstar/README.md`. |
| `nwirp/` | NWIRP building-block and snippet builders. See `nwirp/README.md`. |

### OoXml Module

The `OoXml/` subdirectory is a PowerShell module for generic OOXML style and numbering manipulation. The NIST style script imports this module and passes its configuration data.

```powershell
Import-Module ./OoXml/OoXml.psd1
```

**Style functions**: `Edit-StyleBlock`, `Remove-StyleBlock`, `Get-AllStyleIdsByType`, `Set-StyleBasedOn`, `Set-StyleColor`, `Edit-RunRPrInStyle`, `Set-StyleRecommended`, `Set-StyleHidden`

**Numbering functions**: `Get-HeadingAbstractNumIds`, `Set-NumberingSuffix`, `Sync-ListIndents`, `Set-BulletGlyph`

**Scriptblock variables**: `$ListSingleSpacingEdit`, `$BodyTextSpacingEdit`, `$SuppressNumbering` (passed to `Edit-StyleBlock` as mutators)

## Workflows

### Find drift across templates

```powershell
Compare-WordTemplates.ps1 -Folder "C:\path\to\templates"
# Or with a reference file:
Compare-WordTemplates.ps1 -Folder "C:\templates" -Reference "master.dotx"
```

### Edit template XML by hand

```powershell
# 1. Unpack (backs up original)
Expand-WordPackage.ps1 -File "master.dotx" -OpenEditor
# 2. Edit XML in VSCode
# 3. Repack (validates before overwriting)
Compress-WordPackage.ps1 -SourceDir "master_unpacked"
```

## Word Wildcard Find and Replace

Text fixes done in Word itself (not at the OOXML layer). In Find and Replace
(`Ctrl+H`), click **More** and check **Use wildcards**, then use these patterns.

### Space between a number and the percent sign

Inserts a space where a digit sits directly against `%`.

| Field | Value |
|-------|-------|
| Find what | `([0-9])%` |
| Replace with (regular space) | `\1 %` |
| Replace with (non-breaking space) | `\1^s%` |

To fix values that already have a regular space (`5 %`), put the space in the Find
pattern. Use the optional-space variant to normalize both spaced and unspaced values in
one pass.

| Purpose | Find what | Replace with |
|---------|-----------|--------------|
| Convert an existing regular space to non-breaking | `([0-9]) %` | `\1^s%` |
| Catch spaced and unspaced together, normalize to non-breaking | `([0-9]) {0,1}%` | `\1^s%` |

- `([0-9])` captures the digit into group 1; `\1` puts it back. `%` is literal in wildcard mode.
- `^s` is a non-breaking space and keeps the number and `%` on the same line. It works only in the Replace field, never in Find.
- The space in Find is invisible: put exactly one regular space between `)` and `%` (or before `{0,1}`). `{0,1}` means zero or one of the preceding space.
- A regular space in Find does not match a non-breaking space, so values already using a non-breaking space are skipped. The patterns are idempotent - re-running changes nothing already correct.

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word, no COM automation (except the optional PDF export in `ncstar/Build-NISTReport.ps1` and the `ReapplyAllStyles.bas` macro, which run inside Word)

Network paths (UNC like `\\server\share`) work directly. The scripts resolve them via `ProviderPath`.

## OOXML Reference

### Key files in an unpacked Word document

| Path | What it is |
|------|------------|
| `word/styles.xml` | Paragraph, character, table, and list style definitions |
| `word/theme/theme1.xml` | Theme color palette and theme fonts |
| `word/document.xml` | Main document body content |
| `word/numbering.xml` | List, bullet, and outline numbering definitions |
| `word/header*.xml`, `footer*.xml` | Header and footer content per section |
| `word/settings.xml` | Document-wide settings |
| `word/glossary/document.xml` | Building Blocks |

### Units cheat sheet

| Property | OOXML unit | Conversion |
|----------|-----------|------------|
| Font size (`<w:sz>`) | half-points | 24 = 12pt, 28 = 14pt |
| Spacing (`<w:spacing>`) | twips (1/20 pt) | 240 twips = 12pt |
| Indent (`<w:ind>`) | twips (1/1440 in) | 720 = 0.5 in |
| Line spacing (auto) | 240ths of a line | 240 = single, 480 = double |
| Color (`<w:color>`) | hex, no `#` | `000000` = black |
| Toggle (`<w:b/>`) | presence = true | absent = inherit from base |

### Theme color slots

| Slot | Use |
|------|-----|
| `dk1`, `lt1` | Text dark, Background light |
| `dk2`, `lt2` | Text dark 2, Background light 2 |
| `accent1` through `accent6` | Accent colors |
| `hlink`, `folHlink` | Hyperlink, followed hyperlink |

Use `<w:color w:themeColor="accent1"/>` instead of `<w:color w:val="HEXVAL"/>` to make styles follow the theme.
