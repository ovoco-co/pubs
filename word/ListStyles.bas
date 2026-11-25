Attribute VB_Name = "ListStyles"
' Lines up the list styles of the active document: List Bullet, List Number,
' List Continue, and List, levels 1-5. Level 1's bullet or number sits at the
' left margin plus FIRST_IN, each level below steps STEP_IN further in, the text
' sits GAP_IN to the right of the marker, and List Continue and List start at
' the text column. Change the constants to change the layout.
Option Explicit

Private Const FIRST_IN As Single = 0
Private Const STEP_IN As Single = 0.2
Private Const GAP_IN As Single = 0.15

Public Sub AlignListStyles()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim tracking As Boolean
    tracking = doc.TrackRevisions
    doc.TrackRevisions = False

    Dim lvl As Integer
    For lvl = 1 To 5
        SetBulletStyle doc, lvl
        SetNumberStyle doc, lvl
        SetPlainStyle doc, StyleName("List Continue", lvl), lvl
        SetPlainStyle doc, StyleName("List", lvl), lvl
    Next lvl

    If MsgBox("List styles are set." & vbCr & vbCr & _
              "Also clear direct indents and numbering from paragraphs in these styles, " & _
              "so they line up with the styles?", vbYesNo + vbQuestion, "List styles") = vbYes Then
        ResetListParagraphs doc
    End If

    doc.TrackRevisions = tracking
End Sub

Private Function StyleName(baseName As String, lvl As Integer) As String
    If lvl = 1 Then StyleName = baseName Else StyleName = baseName & " " & lvl
End Function

Private Function MarkerPos(lvl As Integer) As Single
    MarkerPos = InchesToPoints(FIRST_IN + STEP_IN * (lvl - 1))
End Function

Private Function TextPos(lvl As Integer) As Single
    TextPos = InchesToPoints(FIRST_IN + STEP_IN * (lvl - 1) + GAP_IN)
End Function

Private Function GetStyle(doc As Document, nm As String) As Style
    On Error Resume Next
    Set GetStyle = doc.Styles(nm)
    On Error GoTo 0
    If GetStyle Is Nothing Then Set GetStyle = doc.Styles.Add(nm, wdStyleTypeParagraph)
End Function

Private Sub SetIndents(st As Style, lvl As Integer, hanging As Boolean)
    With st.ParagraphFormat
        .LeftIndent = TextPos(lvl)
        If hanging Then .FirstLineIndent = MarkerPos(lvl) - TextPos(lvl) Else .FirstLineIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
End Sub

Private Sub LinkList(doc As Document, st As Style, lvl As Integer, fmt As String, numStyle As WdListNumberStyle, fontName As String)
    Dim lt As ListTemplate
    Set lt = doc.ListTemplates.Add(OutlineNumbered:=False)
    With lt.ListLevels(1)
        ' the format and font go first: Word refuses a bullet style before the bullet character is set
        .NumberFormat = fmt
        If fontName <> "" Then .Font.Name = fontName
        .NumberStyle = numStyle
        .TrailingCharacter = wdTrailingTab
        .NumberPosition = MarkerPos(lvl)
        .TextPosition = TextPos(lvl)
        .TabPosition = TextPos(lvl)
        .Alignment = wdListLevelAlignLeft
        .StartAt = 1
        .ResetOnHigher = 0
        .LinkedStyle = st.NameLocal
    End With
End Sub

Private Sub SetBulletStyle(doc As Document, lvl As Integer)
    Dim st As Style
    Set st = GetStyle(doc, StyleName("List Bullet", lvl))
    Select Case lvl
        Case 1: LinkList doc, st, lvl, ChrW(&HF0B7), wdListNumberStyleBullet, "Symbol"      ' filled disc
        Case 2: LinkList doc, st, lvl, "o", wdListNumberStyleBullet, "Courier New"          ' open circle
        Case 3: LinkList doc, st, lvl, ChrW(&HF06E), wdListNumberStyleBullet, "Wingdings"   ' filled square
        Case 4: LinkList doc, st, lvl, ChrW(&HF075), wdListNumberStyleBullet, "Wingdings"   ' filled diamond
        Case 5: LinkList doc, st, lvl, ChrW(&HF071), wdListNumberStyleBullet, "Wingdings"   ' open square
    End Select
    SetIndents st, lvl, True
End Sub

Private Sub SetNumberStyle(doc As Document, lvl As Integer)
    Dim st As Style
    Set st = GetStyle(doc, StyleName("List Number", lvl))
    Select Case lvl
        Case 2: LinkList doc, st, lvl, "%1)", wdListNumberStyleLowercaseLetter, ""
        Case 3: LinkList doc, st, lvl, "%1.", wdListNumberStyleLowercaseRoman, ""
        Case Else: LinkList doc, st, lvl, "%1.", wdListNumberStyleArabic, ""
    End Select
    SetIndents st, lvl, True
End Sub

Private Sub SetPlainStyle(doc As Document, nm As String, lvl As Integer)
    Dim st As Style
    Set st = GetStyle(doc, nm)
    SetIndents st, lvl, False
End Sub

Private Sub ResetListParagraphs(doc As Document)
    Dim p As Paragraph, nm As String
    For Each p In doc.Paragraphs
        nm = p.Style.NameLocal
        If nm Like "List Bullet*" Or nm Like "List Number*" Or nm Like "List Continue*" Or nm = "List" Or nm Like "List #" Then
            p.Range.ParagraphFormat.Reset
            p.Style = nm
        End If
    Next p
End Sub
