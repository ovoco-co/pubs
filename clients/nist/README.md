# NIST Template Tools

PowerShell scripts for applying NIST template standards to Word documents. These tools apply to both NCSTAR and NWIRP document sets. Document-set-specific build tools live in `ncstar/` and `nwirp/`.

Generic OOXML tools live in `../../word/`. The NIST scripts are thin wrappers that supply NIST-specific configuration (color palettes, style lists, margin defaults) and delegate to the generic tools.

## Pipeline Order

Order matters. Run in this sequence:

1. `Set-NISTTheme.ps1` - installs the Resolution Blue color palette so styles can reference theme slots
2. `Set-NISTStyleArchitecture.ps1` - rebases styles onto the `xx*` family, deletes junk styles, applies theme-aligned colors, sets canonical bullet glyphs
3. `../../word/Set-PageLayout.ps1` - sets page margins (generic, moved from this directory)
4. `../../word/Set-DocTitle.ps1` - sets document Title metadata (generic, moved from this directory)
5. `Set-NISTHeaderFooterContent.ps1` - installs even/odd headers and footers with STYLEREF and TITLE+PAGE fields

## Scripts

### Template Setup

| Script | What it does |
|--------|-------------|
| `Set-NISTTheme.ps1` | Thin wrapper: passes the NIST NCSTAR palette to `../../word/Set-ThemeColors.ps1`. |
| `Set-NISTStyleArchitecture.ps1` | Imports `../../word/OoXml/OoXml.psd1` and passes NIST style configuration. Rebases Heading1-9, BodyText, Caption, List Bullet, List Number onto `xx*`. Deletes junk styles. Applies theme colors. Sets bullet-glyph ladder. |
| `Set-NISTHeaderFooterContent.ps1` | Installs 6 header/footer parts (even/odd/first). Even header: chapter number via STYLEREF. Odd header: chapter title. Footers: TITLE + PAGE with mirror layout. |
| `Install-NISTThemeColors.ps1` | Thin wrapper: passes the NIST palette to `../../word/Install-ThemeColors.ps1`. Per-user; does not touch documents. |

### Validation

| Script | What it does |
|--------|-------------|
| `Audit-NISTStyleSet.ps1` | Thin wrapper: passes 48 required + 24 optional NIST style ids to `../../word/Audit-StyleSet.ps1`. Pre-assembly gate for Build-NISTReport. |

### Moved to Generic (../../word/)

These scripts were previously in this directory but are now generic tools:

| Former name | New location | Notes |
|-------------|-------------|-------|
| `Set-NISTDocTitle.ps1` | `../../word/Set-DocTitle.ps1` | Fully generic, no NIST logic |
| `Set-NISTPageLayout.ps1` | `../../word/Set-PageLayout.ps1` | Parameterized margins, no NIST defaults |
| `Set-NISTTableLook.ps1` | `../../word/Set-TableLook.ps1` | Standard formal report defaults |

## Example: Apply template architecture to a chapter folder

```powershell
$dir   = "C:\HM\NCSTAR-4-1"
$title = "NIST NCSTAR 4-1, Hurricane Maria Investigation"

Set-NISTTheme.ps1 -Folder $dir
Set-NISTStyleArchitecture.ps1 -Folder $dir
..\..\word\Set-PageLayout.ps1 -Folder $dir
..\..\word\Set-DocTitle.ps1 -Folder $dir -Title $title
Set-NISTHeaderFooterContent.ps1 -Folder $dir
```

## Adapting for a new client

Create a new folder under `../` (e.g., `../jpl/`) and follow the thin-wrapper pattern:

- Theme: call `../../word/Set-ThemeColors.ps1` with your palette
- Style architecture: import `../../word/OoXml/OoXml.psd1` and pass your style config
- Audit: call `../../word/Audit-StyleSet.ps1` with your required/optional style lists
- Page layout, doc title, table look: use the generic scripts directly

See `../../word/README.md` for the full list of generic OOXML tools.
