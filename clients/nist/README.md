# NIST Template Tools

PowerShell scripts for applying NIST template standards to Word documents. These tools apply to both NCSTAR and NWIRP document sets. Document-set-specific build tools live in `ncstar/` and `nwirp/`.

Generic OOXML tools live in `../../word/`. The NIST scripts are thin wrappers that supply NIST-specific configuration (color palettes, style lists, margin defaults) and delegate to the generic tools.

## Pipeline Order

Order matters. Run in this sequence:

1. `Set-NISTTheme.ps1` - installs the Resolution Blue color palette so styles can reference theme slots
2. `Set-NISTStyleArchitecture.ps1` - rebases styles onto the `xx*` family, deletes junk styles, applies theme-aligned colors, sets canonical bullet glyphs
3. `../../word/Set-PageLayout.ps1` - sets page margins (generic; pass NIST margins as parameters)
4. `../../word/Set-TableLook.ps1` - normalizes table looks to Header Row + Repeat Header + Banding
5. `../../word/Set-DocTitle.ps1` - sets document Title metadata (generic, moved from this directory)
6. `Set-NISTHeaderFooterContent.ps1` - installs even/odd headers and footers with STYLEREF and TITLE+PAGE fields

Step 1 must run before step 2: the styles reference the theme slots the palette creates.

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

## Workflow: apply the NIST template to a chapter folder

Run the scripts in place, from their repo location. The `Set-NIST*` wrappers find their
generic engine in `../../word/` by walking up from their own location, so do NOT copy them
into the document folder - that breaks the delegation. Point `-Folder` at wherever the
`.docx` files live (a local Windows folder is fine).

The `&` call operator runs each script from its full path without changing directory, so
`$PSScriptRoot` still resolves correctly.

```powershell
# Script locations (WSL path from Windows; adjust the distro name if yours differs)
$nist = "\\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\clients\nist"
$word = "\\wsl$\Ubuntu-24.04\home\artist\ovoco\tools\publishing\word"

# Folder holding the chapter .docx files
$docs = "$env:USERPROFILE\Documents\260726"

& "$nist\Set-NISTTheme.ps1"               -Folder $docs
& "$nist\Set-NISTStyleArchitecture.ps1"   -Folder $docs
& "$word\Set-PageLayout.ps1"              -Folder $docs -InnerInches 1.25 -OuterInches 0.75
& "$word\Set-TableLook.ps1"               -Folder $docs
& "$word\Set-DocTitle.ps1"                -Folder $docs   # no -Title, so PowerShell prompts
& "$nist\Set-NISTHeaderFooterContent.ps1" -Folder $docs
```

Notes:

- **Title prompt.** `Set-DocTitle.ps1` declares `-Title` mandatory. Leave it off and PowerShell pauses to ask (`Title:`); type the report title (for example `NIST NCSTAR 4-3, Hurricane Maria Investigation`) and it applies to every file in the folder. Pass `-Title "..."` to skip the prompt, or `-File "$docs\Chapter-00.docx"` for a per-file title.
- **NIST margins.** `Set-PageLayout.ps1` defaults to 1.0" all around; the `-InnerInches 1.25 -OuterInches 0.75` above are the NIST mirror-margin values.
- **Confirmation.** Each script lists its targets and asks `Proceed? (y/N)`. Add `-Force` to any of them to skip it.
- `Set-TableLook.ps1` is included here because the NIST template repeats table header rows; `Reapply-Styles.ps1` (`../../word/`) is a separate post-authoring cleanup step, not part of this build.

## Adapting for a new client

Create a new folder under `../` (e.g., `../jpl/`) and follow the thin-wrapper pattern:

- Theme: call `../../word/Set-ThemeColors.ps1` with your palette
- Style architecture: import `../../word/OoXml/OoXml.psd1` and pass your style config
- Audit: call `../../word/Audit-StyleSet.ps1` with your required/optional style lists
- Page layout, doc title, table look: use the generic scripts directly

See `../../word/README.md` for the full list of generic OOXML tools.
