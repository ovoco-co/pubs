# Publishing Toolkit

PowerShell scripts and tools for documentation engineering across multiple publishing workflows and clients.

## Publishing Paths

### Word (`word/`)

OOXML inspection and manipulation tools. The generic scripts work on any Word document from any client: inspect styles, compare templates for drift, unpack/repack OOXML packages, audit theme colors, strip direct formatting and shading. No Word installation required. The NIST template wrappers (`Set-NIST*`, `Audit-NISTStyleSet`) and the NCSTAR/NWIRP report builders (`word/ncstar/`, `word/nwirp/`) live here too. See `word/readme.md`.

Source conversion into the template lives here as well: `word/Build-NISTReferenceDoc.ps1` builds the pandoc reference document from the NIST template, and `word/Convert-ToNISTWord.ps1` turns HTML, LaTeX, and Markdown into NIST-formatted Word files. These two are the only scripts in the toolkit that need an external binary (pandoc).

### Charts (`charts/`)

Chart styling VBA and the guide it implements: `StyleCharts_PowerPoint.bas` and `StyleCharts_Excel.bas` apply the NWIRP publication look - palette, axes, gridlines, legends, labels - to charts that already exist in a deck or workbook. They restyle; they do not build charts from data. The guide also covers getting a chart into Word as an EMF and the 508 alt-text checklist. See `charts/readme.md`.

### Markdown and pandoc utilities (`md-docx/`, `scrape/`)

Generic conversion and capture scripts with no template awareness: Markdown to Word, Word to Markdown, PDF to Markdown, HTML to Markdown, and documentation site scrapers. Use these for throwaway or interchange work. Anything that has to land in a client template goes through the Word path instead.

### DITA (`dita/`)

Word-to-DITA conversion pipeline. Extracts content from Word documents, maps paragraph styles to DITA elements via a configurable style map, and produces topic files, a ditamap, and metadata extracts (numbering, headers/footers, tables, images).

### FrameMaker (`frame/`)

Adobe FrameMaker template tooling for both modes:

- `frame/unstructured/` - Traditional FrameMaker: master pages, paragraph catalogs, import format mappings
- `frame/structured/dita/` - DITA publishing through FrameMaker
- `frame/structured/s1000d/` - S1000D military/aerospace technical publications
- `frame/structured/docbook/` - DocBook book-oriented XML publishing

Note: `dita/` converts Word content INTO DITA format. `frame/structured/dita/` publishes DITA content THROUGH FrameMaker. These are different workflows.

## Client Implementations (`clients/`)

Per-client notes. The NIST template tools were consolidated into `word/` (wrappers alongside the generic engines; `word/ncstar/` and `word/nwirp/` for the report builders), so all Word scripts live in one folder.

- NIST - now in `word/` (see `word/readme.md`)
- `clients/navy/` - Navy-specific tools (placeholder)

## Onboarding a New Client

Client scripts are thin wrappers that supply client configuration (palette, style list, margins) and delegate to the generic engines in `word/`:

1. Add `Set-<Client>Theme.ps1` calling `word/Set-ThemeColors.ps1` with the client palette
2. Add `Set-<Client>StyleArchitecture.ps1` importing `word/OoXml/OoXml.psd1` with the client style config
3. Add thin wrappers for page layout, audit, and headers/footers as needed
4. Keep wrappers in `word/` so they resolve the engines and `OoXml/` module by their own location
5. The generic engines in `word/` remain unchanged

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word installation, no COM automation
- All OOXML tools use System.IO.Compression and System.Xml
- Pandoc, for the source-conversion scripts only (`word/Convert-ToNISTWord.ps1` and the `md-docx/` wrappers). Every other script runs without it

## Pipeline Independence

Each publishing path (Word, DITA, FrameMaker) is independent. A team that only needs DITA conversion does not need FrameMaker tools. A team that only needs Word template cleanup does not need DITA tools.
