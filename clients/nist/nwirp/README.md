# NWIRP Report Tools

PowerShell scripts for building NIST NWIRP deliverables (Hurricane Maria / Puerto Rico schools-and-hospitals work). These are specific to the NWIRP document set and depend on the shared NIST template tools in the parent directory (`../`).

## Scripts

| Script | What it does |
|--------|-------------|
| `Build-NWIRPBuildingBlocks.ps1` | Adds 22 NIST NCSTAR/NWIRP-tailored building blocks to an unpacked Word template. Creates glossary scaffolding if missing. Registers content types. Adds FrontMatterHeader and AppendixHeader styles. Idempotent - re-running adds only blocks not already present. |
| `Build-NWIRPSnippetsDoc.ps1` | Builds a standalone reference docx containing the 22 snippets as labeled sections with page breaks. Intended as a copy-and-paste source rather than a Quick Parts gallery. Alternative to Build-NWIRPBuildingBlocks.ps1, not a companion. |

## Building Block Categories

The 22 blocks are distributed across six phases:

- **Front matter**: NCSTAR Cover Page, NIST Standard Disclaimer, Abstract Block, Acknowledgments Block
- **Chapter structure**: Chapter Opener, Section Divider
- **Recurring artifacts**: Figure Caption Block, Table Caption Block, Methodology Box, Key Findings Sidebar, Interview Quote Block
- **Schools/Hospitals comparison**: Schools-Hospitals Comparison, Demographic Summary Table, Recovery Status Legend
- **Recovery model results**: Recovery Model Results Table
- **End matter**: References Section Opener, Bibliography Entry, Appendix Opener, Acronyms Appendix Table, Stats Inline Result, Effect Size Citation, Sample Size Statement

## Example: Add building blocks to a template

```powershell
# 1. Unpack the template
../../word/Expand-WordPackage.ps1 -File "HM-NWIRP-Template.dotx"

# 2. Add building blocks
Build-NWIRPBuildingBlocks.ps1 -SourceDir "HM-NWIRP-Template_unpacked"

# 3. Repack
../../word/Compress-WordPackage.ps1 -SourceDir "HM-NWIRP-Template_unpacked"
```

## Prerequisites

Run the shared NIST pipeline first (theme, styles, layout) on the template before adding building blocks. The blocks reference existing custom styles (BodyText, Caption, Disclaimer, References, TableandFigureFootnotes, Heading1).
