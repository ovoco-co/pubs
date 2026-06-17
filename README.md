# Publishing Toolkit

PowerShell scripts and tools for documentation engineering across multiple publishing workflows and clients.

## Publishing Paths

### Word (`word/`)

Generic OOXML inspection and manipulation tools. Work on any Word document from any client or organization. Inspect styles, compare templates for drift, unpack/repack OOXML packages, audit theme colors, strip direct formatting. No Word installation required.

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

Client-specific template tools that customize the generic toolkit for a particular organization's standards:

- `clients/nist/` - Shared NIST template tools (theme, style architecture, layout, headers/footers)
  - `clients/nist/ncstar/` - NCSTAR report builds and Excel VBA
  - `clients/nist/nwirp/` - NWIRP report builds and Excel VBA
- `clients/navy/` - Navy-specific tools

## Onboarding a New Client

1. Create a folder under `clients/` named for the client (e.g., `clients/jpl/`)
2. Copy the NIST shared scripts as a starting point
3. Rename scripts to use the new client prefix (e.g., `Set-JPLTheme.ps1`)
4. Modify the color palette, style names, page layout, and header/footer content
5. Add document-set-specific build tools in subfolders as needed
6. The generic tools in `word/` remain unchanged

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word installation, no COM automation
- All OOXML tools use System.IO.Compression and System.Xml

## Pipeline Independence

Each publishing path (Word, DITA, FrameMaker) is independent. A team that only needs DITA conversion does not need FrameMaker tools. A team that only needs Word template cleanup does not need DITA tools.
