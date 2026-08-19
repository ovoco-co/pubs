# Word-to-DITA Conversion Pipeline

PowerShell scripts for converting Word documents into DITA XML. Extracts content, maps Word paragraph styles to DITA elements via a configurable style map, and produces topic files, a ditamap, and metadata extracts.

This pipeline converts content INTO DITA format. For publishing DITA content through FrameMaker, see `../frame/structured/dita/`.

## Scripts

| Script | What it does |
|--------|-------------|
| `Run-All.ps1` | Orchestrator that runs the full pipeline. Calls the four scripts below in sequence. |
| `Convert-WordToDita.ps1` | Reads word/document.xml, splits on H1 headings into separate topic files, maps styles to DITA elements via stylemap.json, and generates a ditamap. |
| `Extract-Numbering.ps1` | Extracts numbering definitions from word/numbering.xml into a JSON file (abstract numbering, level formats, level text patterns). |
| `Extract-HeadersFooters.ps1` | Extracts text content from all header*.xml and footer*.xml parts into a JSON file. |
| `Extract-TablesImages.ps1` | Extracts all tables as JSON (rows and cells as text) and copies all images from word/media/ to an output folder. |

## Style Map

`stylemap.json` maps Word paragraph style names to DITA element types:

| Style Name | DITA Output |
|------------|-------------|
| H1, Heading 1 | Topic split point (new topic per H1) |
| H2-H5, Heading 2-5 | `<section>` nesting |
| Normal, Body, Body Text | `<p>` |
| Note | `<note>` |
| Warning, Caution | `<note type="warning">`, `<note type="caution">` |
| Code, Code Block | `<codeblock>` |
| Bullet, List Bullet | `<ul><li>` |
| Numbered, List Number | `<ol><li>` |
| Caption | `<fig><title>` |

Unmapped styles default to `<p>`.

To support a client's custom styles (e.g., NIST's `xxbase`, `xxheader` families), add entries to stylemap.json mapping those style names to the appropriate DITA types.

## Usage

```powershell
# Run the full pipeline
Run-All.ps1 -InputDocx "source.docx" -WorkingDir "output"

# Or run individual scripts
Convert-WordToDita.ps1 -InputDocx "source.docx" -OutputDir "output/dita" -StyleMapJson "stylemap.json"
Extract-Numbering.ps1 -InputDocx "source.docx" -OutputJson "output/numbering.json"
Extract-HeadersFooters.ps1 -InputDocx "source.docx" -OutputJson "output/headers-footers.json"
Extract-TablesImages.ps1 -InputDocx "source.docx" -OutputDir "output/assets"
```

## Output Structure

```
output/
  dita/
    topic.dita           One .dita file per H1 section
    map.ditamap          References all topic files
  numbering.json         Numbering definitions (list formats, levels)
  headers-footers.json   Header and footer text content
  assets/
    tables.json          All tables as row/cell text arrays
    images/              All images from word/media/
```

## Requirements

- PowerShell 5.1 or PowerShell 7+
- No Word installation required

## Limitations

- Splits on H1 only. Documents without H1 headings produce a single topic.
- Inline formatting (bold, italic, links) within runs is not preserved in the DITA output. Only paragraph-level text is extracted.
- Images are extracted but not referenced from the DITA topics. Image placement must be done manually.
- The converter reads style names from word/document.xml. If a document uses custom style names not in stylemap.json, content maps to `<p>` by default. Use `../../word/Get-WordStyles.ps1` to discover the actual style names in a document.
