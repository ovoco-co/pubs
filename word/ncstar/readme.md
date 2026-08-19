# NCSTAR Report Tools

PowerShell scripts for building and validating NIST NCSTAR deliverables. These are specific to the NCSTAR document set and depend on the shared NIST template tools in the parent directory (`../`).

## Scripts

| Script | What it does |
|--------|-------------|
| `Build-NISTReport.ps1` | Combines chapter files (Chapter-00 front matter, Chapter-01..NN, Appendix-A..Z) into a single deliverable docx. Optionally exports PDF via Word COM. Renumbers image/hyperlink rels, footnote/endnote/bookmark ids per chapter. Strips comments and accepts revisions. Marks TOC, LOF, LOT fields dirty. Runs Audit-NISTStyleSet.ps1 as a pre-assembly gate. |
| `Test-NISTFrontMatter.ps1` | Audits a Chapter-00 front-matter file against the canonical NIST front-matter sequence for either a Project Report or a Summary Report. Reports missing sections, extras, and out-of-order items. Returns non-zero on drift. |

## Example: Assemble a multi-chapter report

```powershell
$dir = "C:\Users\Admin\Documents\260512"

# 1. Validate front matter sequence
Test-NISTFrontMatter.ps1 -Folder $dir -Series NCSTAR_4-3 -ReportType Project

# 2. Assemble chapters into a single deliverable + PDF
Build-NISTReport.ps1 -Folder $dir -Series NCSTAR_4-3 `
    -Output "$dir\_assembled" -Pdf -Force
```

## Known Dependencies

- `Build-NISTReport.ps1` references `Audit-NISTStyleSet.ps1` in the parent `../` (now `publishing/word/`) via `$PSScriptRoot`. Resolves correctly after the consolidation into `word/`.

## Prerequisites

Run the shared NIST pipeline first (theme, styles, layout, title, headers/footers) on all chapter files before assembling.
