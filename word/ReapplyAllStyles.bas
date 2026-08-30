Attribute VB_Name = "ReapplyAllStyles"
' Re-applies each paragraph's assigned style across the whole document -
' body, headers, footers, footnotes, textboxes - the same as clicking the
' style in the gallery, but on every paragraph.
'
' Use this for paragraphs that keep a stale look after Word's Import Styles
' until they are manually re-clicked, even when there is NO direct formatting
' to clear. That is a Word rendering-refresh issue: the file is already
' correct, so a .docx-editing script cannot fix it - only Word re-applying
' the style can.
'
' This does NOT clear direct formatting (neither does clicking a style). For
' that, use Remove-Shading.ps1 (shading) or Reapply-Styles.ps1 (character
' formatting).
'
' To run: Alt+F11 in Word, Insert > Module, paste this, press F5.

Sub ReapplyAllStyles()
    Dim story As Range
    Dim para As Paragraph
    Dim n As Long

    Application.ScreenUpdating = False
    On Error Resume Next   ' some story ranges are empty or not present

    For Each story In ActiveDocument.StoryRanges
        Do
            For Each para In story.Paragraphs
                para.Range.Style = para.Range.Style   ' re-apply = the "click"
                n = n + 1
            Next para
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next story

    Application.ScreenUpdating = True
    MsgBox "Reapplied styles to " & n & " paragraph(s).", vbInformation
End Sub
