# Word Template Toolkit

Generic PowerShell scripts for working with Word templates at the OOXML layer. Reads files directly as zip archives of XML using `System.IO.Compression` and `System.Xml`, so Word does not need to be installed. Works on Windows, WSL, Linux, and macOS with PowerShell 5.1+.

These tools are client-agnostic. They work on any Word document from any organization. Client-specific template tools live in `../clients/<client>/`.

## Scripts

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
| `Set-DocTitle.ps1` | Sets `<dc:title>` in docProps/core.xml. Works for any title string in any Word document. |
| `Set-PageLayout.ps1` | Sets page margins with mirror margins. Configurable via -InnerInches, -OuterInches, -TopInches, -BottomInches, -HeaderInches, -FooterInches. |
| `Set-TableLook.ps1` | Normalizes `<w:tblLook>` to formal report defaults (Header Row + Repeat Header + Banding) and adds `<w:tblHeader/>` to first rows. |
| `Set-ThemeColors.ps1` | Replaces the `<a:clrScheme>` in theme1.xml with a caller-supplied color scheme XML block. Validates all 12 theme slots are present. |
| `Install-ThemeColors.ps1` | Installs a theme color scheme file into the user's Word Theme Colors folder so Design > Colors displays it by name. |
| `Audit-StyleSet.ps1` | Generic style audit engine. Checks that caller-supplied required style ids exist in styles.xml. Returns non-zero on drift. JSON or table output. |

### OoXml Module

The `OoXml/` subdirectory is a PowerShell module for generic OOXML style and numbering manipulation. Client scripts import this module and pass their configuration data.

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

### Make theme colors permanent

```powershell
# 1. Audit current color usage
Audit-WordTheme.ps1 -File "master.dotx"
# 2. Unpack and edit theme1.xml and styles.xml
# 3. Convert explicit colors to theme references
# 4. Repack and verify
```

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word, no COM automation

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
