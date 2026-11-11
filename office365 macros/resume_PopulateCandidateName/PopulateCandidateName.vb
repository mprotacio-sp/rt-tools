Sub CandidateName()

  ' Define variables
  Dim nPageNum As Integer
  Dim objDoc As Document
  Dim strFindText As String
  Dim strReplaceText As String

  ' -- Step 1: Collect candidate name from user -------------------------------

  Set objDoc = ActiveDocument
  strFindText = "[Candidate Name]"
  strReplaceText = InputBox("Please PASTE the candidate's name.", "Update Title and Header")

  ' -- Sanitize input ---------------------------------------------------------
  '
  ' Pasted text can carry hidden characters that silently corrupt output:
  '   vbCr / vbLf / vbCrLf  — carriage returns and line feeds
  '   vbTab                  — tab characters
  '   Chr(160)               — non-breaking space (common from web/PDF paste)
  ' Replace each with a regular space, then collapse any runs of multiple
  ' spaces down to one, and trim leading/trailing whitespace.
 
  strReplaceText = Replace(strReplaceText, vbCrLf, " ")  ' Windows line ending first
  strReplaceText = Replace(strReplaceText, vbCr, " ")
  strReplaceText = Replace(strReplaceText, vbLf, " ")
  strReplaceText = Replace(strReplaceText, vbTab, " ")
  strReplaceText = Replace(strReplaceText, Chr(160), " ") ' Non-breaking space
 
  ' Collapse runs of multiple spaces to a single space
  Do While InStr(strReplaceText, "  ") > 0
      strReplaceText = Replace(strReplaceText, "  ", " ")
  Loop
 
  strReplaceText = Trim(strReplaceText)
 

  ' Handle cancel (X button or Esc returns StrPtr = 0)
  If StrPtr(strReplaceText) = 0 Then
      MsgBox ("Canceled. Exiting script.")
      Exit Sub
  ' Handle empty submission
  ElseIf strReplaceText = vbNullString Then
      MsgBox ("Nothing entered. Exiting script.")
      Exit Sub
  Else
      ' Confirm the name before making any changes
      Dim response As Integer
      response = MsgBox("'" & strReplaceText & "'" & vbCrLf & vbCrLf & _
                        "The above text will be added to the top of the document, " & _
                        "the header of page 2, and subsequent pages. Proceed?", vbOKCancel)
      If response = 2 Then
          MsgBox ("User canceled. Exiting script.")
          Exit Sub
      End If
  End If

  ' -- Step 2: Insert a temporary page break ----------------------------------
  '
  ' Word only exposes the P2+ header when a second page physically exists.
  ' We insert a manual page break at the top of the document to guarantee
  ' the cursor can reach page 2. This break is removed explicitly in Step 4
  ' using Find/Replace for the ^m character — a more reliable approach than
  ' counting backspaces, which behaves differently on Windows vs. Mac.

  Selection.HomeKey unit:=wdStory
  Selection.InsertBreak Type:=wdPageBreak

  ' -- Step 3: Replace [Candidate Name] in the header of every page -----------
  '
  ' We loop page by page so Word activates each page's header context in turn,
  ' which is necessary for documents that use "Different First Page" or
  ' section-level header overrides. wdFindStop (instead of wdFindContinue)
  ' prevents the search from wrapping and re-replacing text already updated.

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
              .Wrap = wdFindStop       ' Stop at end of header; do not wrap
              .Execute Replace:=wdReplaceAll
          End With
      End With
  Next nPageNum

  ' -- Step 4: Return to the main document and remove the temporary page break -
  '
  ' We use Find/Replace targeting ^m (manual page break) with wdFindStop and
  ' wdReplaceOne to remove exactly one instance — the break we added in Step 2.
  ' This avoids the cursor-position ambiguity that caused the name to be typed
  ' twice on Windows when relying on TypeBackspace to clean up.

  objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

  With objDoc.Content.Find
      .ClearFormatting
      .Text = "^m^p"                 ' ^m = manual page break
      .Replacement.ClearFormatting
      .Replacement.Text = ""
      .Wrap = wdFindStop             ' Do not wrap; stop after first match
      .Execute Replace:=wdReplaceOne ' Remove exactly one page break
  End With

  ' -- Step 5: Write the candidate name at the top of the document ------------

  Selection.HomeKey unit:=wdStory
  Selection.TypeText Text:=strReplaceText
  Selection.Style = ActiveDocument.Styles("Candidate Name")
  Selection.TypeParagraph
  Selection.Style = ActiveDocument.Styles("Normal")

End Sub