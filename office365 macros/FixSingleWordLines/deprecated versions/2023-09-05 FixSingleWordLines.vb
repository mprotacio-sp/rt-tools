' CHANGELOG
' - works with lines at the end of a page (where line count is negative)
'      - this isn't perfect; it will add an NBSP even if the para is a single line;
'           I don't imagine this becoming an issue, but can likely fix by
'           adjusting the GetParagraphLineCount function to deal with negative line counts
' - ignores paragraphs that have already been adjusted_manually
' - ignores paragraphs that have multiple_final_words_strung_together_with_NBSPs
'      (i.e., prevents turning all spaces into NBSPs if run multiple times)
' - fixes issue where a blank bullet would be added at end of documents
'   that end in a bulleted list (i.e., ignores paragraph marker at end of paragraphs)
' - prints to console for better debugging
' - 2023-09-05: adjusted to only replace the space itself and nothing else;
'      preserves all surrounding formatting

Sub FixSingleWordLines()
    Dim regex As Object
    Dim para As Paragraph
    Dim searchRange As Range
    Dim match As Object
    Dim pos As Long

    ' Create a new regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
    End With

    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.Text) > 2 Then
            If GetParagraphLineCount(para) <> 1 Then
                Set searchRange = para.Range
                searchRange.MoveEnd unit:=wdCharacter, Count:=-1
                regex.Pattern = "(\s)(\S+\s*(?: \s*)*)$"
                
                ' Check if the paragraph contains a match of the regular expression
                Set match = regex.Execute(searchRange.Text)
                
                If match.Count > 0 Then
                    ' Check if the second submatch contains a non-breaking space
                    If InStr(match.Item(0).SubMatches(1), ChrW(160)) > 0 Then
                        Debug.Print "Paragraph already adjusted: " & searchRange.Text
                    Else
                        ' Find the position of the space to be replaced
                        pos = InStrRev(searchRange.Text, " ")
                        
                        ' Replace only the space character at the specified position
                        searchRange.Start = searchRange.Start + pos - 1
                        searchRange.End = searchRange.Start + 1
                        searchRange.Text = ChrW(160)
                    End If
                End If
            End If
        End If
    Next para
End Sub

Function GetParagraphLineCount(para As Paragraph) As Long
    Dim rngFirst As Range, rngLast As Range
    Set rngFirst = para.Range
    rngFirst.Collapse Direction:=wdCollapseStart
    Set rngLast = para.Range
    rngLast.Collapse Direction:=wdCollapseEnd
    GetParagraphLineCount = rngLast.Information(Type:=wdFirstCharacterLineNumber) _
        - rngFirst.Information(Type:=wdFirstCharacterLineNumber)
    Debug.Print "Paragraph line count: " & GetParagraphLineCount
End Function
