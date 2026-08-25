' CHANGELOG
' - works with lines at the end of a page (where line count is negative)
'       - this isn't perfect; it will add an NBSP even if the para is a single line;
'           I don't imagine this becoming an issue, but can likely fix by
'           adjusting the GetParagraphLineCount function to deal with negative line counts
' - ignores paragraphs that have already been adjusted_manually
' - ignores paragraphs that have multiple_final_words_strung_together_with_NBSPs
'       (i.e., prevents turning all spaces into NBSPs if run multiple times)
' - fixes issue where a blank bullet would be added at end of documents
'       that end in a bulleted list (i.e., ignores paragraph marker at end of paragraphs)
' - prints to console for better debugging
' - 2023-09-05: adjusted to only replace the space itself and nothing else;
'       now preserves all surrounding formatting
' - 2023-09-06:
'       - Added more robust debug statements
'       - Realized Word captures NBSPs as part of `\S` in RegEx. 
'           Modified RegEx by creating a character class that is 
'           essentially all printable characters except the space and NBSP.
'       - Passes all tests!!! (except for single last line on a page--v low priority)

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
                regex.Pattern = "(\s)([^\s ]+)\s*([ \t ]*)$"
                
                ' Check if the paragraph contains a match of the regular expression
                Set match = regex.Execute(searchRange.Text)
                
                If match.Count > 0 Then
                    
                    Dim i As Integer
                    For i = 0 To match.Item(0).SubMatches.Count - 1
                        Dim debugStr As String
                        debugStr = match.Item(0).SubMatches(i)
                        
                        ' Replace spaces and tabs with \s
                        debugStr = Replace(debugStr, " ", "\s")
                        debugStr = Replace(debugStr, vbTab, "\s")
                        
                        ' Replace NBSP with *
                        debugStr = Replace(debugStr, ChrW(160), "*")
                        
                        Debug.Print "Capture Group " & i + 1 & ": " & debugStr
                    Next i
                
                    ' Extract the last word from the 2nd capture group
                    Dim lastWord As String
                    lastWord = match.Item(0).SubMatches(1)
                
                    ' Calculate the position in the original string where this last word starts
                    pos = InStrRev(searchRange.Text, lastWord)
                
                    ' Adjust the position to point to the space before the last word
                    pos = pos - 2 ' -1 for zero-based index, -1 for the preceding space
                    
                    ' Replace the space character at the specified position
                    If pos >= 0 Then
                        searchRange.Start = searchRange.Start + pos
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