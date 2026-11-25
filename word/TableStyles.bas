Attribute VB_Name = "TableStyles"
' Sets every table in the active document to Grid Table 1 Light, with the
' header row and banded rows on, the first column, total row, last column,
' and banded columns off, and the first row repeated as the header row.
Option Explicit

Private Const TABLE_STYLE As String = "Grid Table 1 Light"

Public Sub ApplyTableStyle()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim tracking As Boolean
    tracking = doc.TrackRevisions
    doc.TrackRevisions = False

    Dim tbl As Table, done As Long, skipped As Long
    For Each tbl In doc.Tables
        If StyleTable(tbl) Then done = done + 1 Else skipped = skipped + 1
    Next tbl

    doc.TrackRevisions = tracking
    If skipped = 0 Then
        MsgBox done & " table(s) set to " & TABLE_STYLE & ".", vbInformation, "Table style"
    Else
        MsgBox done & " table(s) set to " & TABLE_STYLE & "." & vbCr & skipped & _
               " table(s) could not have their header row set. Set Repeat Header Rows on those by hand.", vbExclamation, "Table style"
    End If
End Sub

Private Function StyleTable(tbl As Table) As Boolean
    tbl.Style = TABLE_STYLE
    tbl.ApplyStyleHeadingRows = True
    tbl.ApplyStyleFirstColumn = False
    tbl.ApplyStyleLastRow = False
    tbl.ApplyStyleLastColumn = False
    tbl.ApplyStyleRowBands = True
    tbl.ApplyStyleColumnBands = False

    ' Word cannot address single rows in a table with vertically merged cells,
    ' so those tables set the header row through a selection
    On Error GoTo MergedRows
    Dim r As Long
    tbl.Rows(1).HeadingFormat = True
    For r = 2 To tbl.Rows.Count
        If tbl.Rows(r).HeadingFormat Then tbl.Rows(r).HeadingFormat = False
    Next r
    StyleTable = True
    Exit Function
MergedRows:
    Resume BySelection
BySelection:
    On Error GoTo Failed
    tbl.Select
    Selection.Rows.HeadingFormat = False
    tbl.Cell(1, 1).Select
    Selection.SelectRow
    Selection.Rows.HeadingFormat = True
    Selection.Collapse
    StyleTable = True
    Exit Function
Failed:
    StyleTable = False
End Function
