# NIST Template Tools

Shared PowerShell scripts for applying NIST template standards to Word documents. These tools apply to both NCSTAR and NWIRP document sets. Document-set-specific build tools live in `ncstar/` and `nwirp/`.

## Pipeline Order

Order matters. Run in this sequence:

1. `Set-NISTTheme.ps1` - installs the Resolution Blue color palette so styles can reference theme slots
2. `Set-NISTStyleArchitecture.ps1` - rebases styles onto the `xx*` family, deletes junk styles, applies theme-aligned colors, sets canonical bullet glyphs
3. `Set-NISTPageLayout.ps1` - sets NIST page margins (1.25" inner, 0.75" outer, 1" top/bottom)
4. `Set-NISTDocTitle.ps1` - sets document Title metadata for the TITLE field in headers/footers
5. `Set-NISTHeaderFooterContent.ps1` - installs even/odd headers and footers with STYLEREF and TITLE+PAGE fields

## Scripts

### Template Setup

| Script | What it does |
|--------|-------------|
| `Set-NISTTheme.ps1` | Replaces the color scheme with the NIST NCSTAR palette (Resolution Blue, Tawny, Shamrock Green, Claret, Picton Blue, Mulberry as accent1-6). Single file or folder. |
| `Set-NISTStyleArchitecture.ps1` | Rebases Heading1-9, BodyText, Caption, List Bullet, List Number onto the `xx*` base family. Deletes junk styles. Fixes Appendix_Head4 to parallel-sibling. Applies theme colors to headers. Sets bullet-glyph ladder. |
| `Set-NISTPageLayout.ps1` | Sets page margins with mirror margins. Configurable via -InnerInches, -OuterInches, -TopInches, -BottomInches, -HeaderInches, -FooterInches. |
| `Set-NISTDocTitle.ps1` | Sets `<dc:title>` in docProps/core.xml. Fully generic - works for any title string. |
| `Set-NISTHeaderFooterContent.ps1` | Installs 6 header/footer parts (even/odd/first). Even header: chapter number via STYLEREF. Odd header: chapter title. Footers: TITLE + PAGE with mirror layout. Adds evenAndOddHeaders and mirrorMargins to settings.xml. |
| `Set-NISTTableLook.ps1` | Normalizes `<w:tblLook>` to NIST shape (Header Row + Repeat Header + Banding) and adds `<w:tblHeader/>` to first rows. |
| `Install-NISTThemeColors.ps1` | Installs the NIST NCSTAR palette into the user's Word Theme Colors folder. Per-user; does not touch documents. |

### Validation

| Script | What it does |
|--------|-------------|
| `Audit-NISTStyleSet.ps1` | Pre-flight check that every required style id exists in styles.xml. Returns non-zero on drift. Consumed by Build-NISTReport.ps1 as a pre-assembly gate. |

## Example: Apply template architecture to a chapter folder

```powershell
$dir   = "C:\HM\NCSTAR-4-1"
$title = "NIST NCSTAR 4-1, Hurricane Maria Investigation"

Set-NISTTheme.ps1 -Folder $dir
Set-NISTStyleArchitecture.ps1 -Folder $dir
Set-NISTPageLayout.ps1 -Folder $dir
Set-NISTDocTitle.ps1 -Folder $dir -Title $title
Set-NISTHeaderFooterContent.ps1 -Folder $dir
```

## Adapting for a new client

To create tools for a different organization, create a new folder under `../` (e.g., `../jpl/`), copy these scripts, and modify:

- `Set-NISTTheme.ps1` - replace hex values in the clrScheme block
- `Set-NISTStyleArchitecture.ps1` - review the deletion list and style rebasing
- `Set-NISTPageLayout.ps1` - change margin defaults
- `Set-NISTHeaderFooterContent.ps1` - change header/footer layout and border color
- `Set-NISTDocTitle.ps1` - fully generic, works as-is

See `../../word/` for client-agnostic OOXML tools that work with any document.
