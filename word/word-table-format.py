import sys
from pathlib import Path
import win32com.client as win32

def main(doc_path):
    doc_path = Path(doc_path)

    if not doc_path.exists():
        raise FileNotFoundError(f"File not found: {doc_path}")

    # Word constants
    wdStyleTypeParagraph = 1
    wdLineSpaceSingle = 0
    wdAutoFitContent = 1

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(doc_path))

        style_name = "TableText"

        # Get or create style
        try:
            style = doc.Styles(style_name)
        except Exception:
            style = doc.Styles.Add(style_name, wdStyleTypeParagraph)

        # Format the style
        style.Font.Name = "Arial Narrow"
        style.Font.Size = 10
        style.ParagraphFormat.SpaceBefore = 0
        style.ParagraphFormat.SpaceAfter = 0
        style.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

        # Loop through tables
        for tbl in doc.Tables:
            # Apply table style
            tbl.Style = "Grid Table 2"

            # Table style options
            tbl.ApplyStyleHeadingRows = True
            tbl.ApplyStyleRowBands = True
            tbl.ApplyStyleFirstColumn = False
            tbl.ApplyStyleLastColumn = False

            # Repeat first row
            if tbl.Rows.Count >= 1:
                tbl.Rows(1).HeadingFormat = True

            # Fit table to content
            tbl.AutoFitBehavior(wdAutoFitContent)

            # Apply custom text style to table text
            tbl.Range.Style = style_name

        doc.Save()
        doc.Close()

    finally:
        word.Quit()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python word-table-format.py <path-to-docx>")
        sys.exit(1)

    main(sys.argv[1])