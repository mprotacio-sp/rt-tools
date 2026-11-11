Sub CandidateName()

  'define variables
  Dim nPageNum As Integer
  Dim objDoc As Document
  Dim strFindText As String
  Dim strReplaceText As String

  'prompt for Candidate Name
    Set objDoc = ActiveDocument
    strFindText = "[Candidate Name]"
    strReplaceText = InputBox("Please PASTE the candidate's name.", "Update Title and Header")
    
    If StrPtr(strReplaceText) = 0 Then
        MsgBox ("Canceled. Exiting script.")
        Exit Sub
    ElseIf strReplaceText = vbNullString Then
        MsgBox ("Nothing entered. Exiting script.")
        Exit Sub
    Else
        Dim response As Integer
        response = MsgBox("'" & strReplaceText & "'" & vbCrLf & vbCrLf & "The above text will be added to the top of the document, the header of page 2, and subsequent pages. Proceed?", vbOKCancel)
        If response = 2 Then
            MsgBox ("User canceled. Exiting script.")
            Exit Sub
        End If
    End If

  'go to top of document
  Selection.HomeKey unit:=wdStory
  
  'add pagebreak to ensure caret is on page 2
  Selection.InsertBreak Type:=wdPageBreak
 
  'find and replace in header
  For nPageNum = 1 To Selection.Information(wdNumberOfPagesInDocument)
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:=nPageNum
    Application.Browser.Target = wdBrowsePage
    objDoc.Bookmarks("\page").Range.Select
    With objDoc.ActiveWindow
      .ActivePane.View.SeekView = wdSeekCurrentPageHeader
      With .Selection.Find
        .ClearFormatting
        .Text = strFindText
        .Replacement.ClearFormatting
        .Replacement.Text = strReplaceText
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
      End With
    End With
  Next nPageNum
 
  'go to top of document
  objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  Selection.HomeKey unit:=wdStory
  Selection.TypeText Text:=strReplaceText
  Selection.Style = ActiveDocument.Styles("Candidate Name")
  Selection.TypeParagraph
  Selection.Style = ActiveDocument.Styles("Normal")
  Selection.MoveRight unit:=wdCharacter, Count:=2
  Selection.TypeBackspace
  Selection.TypeBackspace

End Sub
