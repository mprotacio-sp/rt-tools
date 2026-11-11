' FILENAME:       FixSingleWordLines.vb
' LATEST UPDATE:  2026-06-17 
' CHANGELOG
' - works with lines at the end of a page (where line count is negative)
'     - this isn't perfect; it will add an NBSP even if the para is a single line; I don't imagine this becoming an issue, but can likely fix by adjusting the GetParagraphLineCount function to deal with negative line counts
' - ignores paragraphs that have already been adjusted_manually
' - ignores paragraphs that have multiple_final_words_strung_together_with_NBSPs (i.e., prevents turning all spaces into NBSPs if run multiple times)
' - fixes issue where a blank bullet would be added at end of documents that end in a bulleted list (i.e., ignores paragraph marker at end of paragraphs)
' - prints to console for better debugging
' - 2023-09-05: adjusted to only replace the space itself and nothing else; now preserves all surrounding formatting
' - 2023-09-06:
'     - Added more robust debug statements
'     - Realized Word captures NBSPs as part of `\S` in RegEx. Modified RegEx by creating a character class that is essentially all printable characters except the space and NBSP.
'     - Passes all tests!!! (except for single last line on a page--v low priority)
' - 2026-05-26:
'     - fixed GetParagraphLineCount logic to prevent negatives
'     - added re-run guard
' - 2026-07-16:
'     - fixed bug where skipped last para on page bug--now runs on all paragraphs
'     - fixed bug where deleted wrong character for paragraphs that have images anchored to them


Sub FixSingleWordLines()
    Dim para As Paragraph
    Dim searchRange As Range

    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.Text) > 2 Then
            If ParagraphHasMultipleLines(para) Then
                Set searchRange = para.Range
                searchRange.MoveEnd unit:=wdCharacter, Count:=-1
                If Not ReplaceLastSpaceWithNbsp(searchRange) Then
                    Debug.Print "No eligible space found"
                End If
            End If
        End If
    Next para
End Sub

Sub FixSingleWordLines_macOS()
    Dim para As Paragraph
    Dim searchRange As Range

    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.Text) > 2 Then
            If ParagraphHasMultipleLines(para) Then
                Set searchRange = para.Range
                searchRange.MoveEnd unit:=wdCharacter, Count:=-1
                If Not ReplaceLastSpaceWithNbsp(searchRange) Then
                    Debug.Print "No eligible space found"
                End If
            End If
        End If
    Next para
End Sub

Function ReplaceLastSpaceWithNbsp(ByVal searchRange As Range) As Boolean
    Dim i As Long
    Dim currentChar As String
    Dim seenWordCharacter As Boolean

    For i = searchRange.Characters.Count To 1 Step -1
        currentChar = searchRange.Characters(i).Text

        Select Case currentChar
            Case " ", vbTab
                If seenWordCharacter Then
                    searchRange.Characters(i).Text = ChrW(160)
                    ReplaceLastSpaceWithNbsp = True
                    Exit Function
                End If
            Case ChrW(160)
                If seenWordCharacter Then
                    Debug.Print "Skipping already-adjusted paragraph"
                    Exit Function
                End If
            Case vbCr, vbLf
                ' Ignore paragraph and line terminators that may appear in the range.
            Case Else
                seenWordCharacter = True
        End Select
    Next i
End Function

Function ParagraphHasMultipleLines(para As Paragraph) As Boolean
    Dim paragraphRange As Range
    Dim lineCount As Long

    Set paragraphRange = para.Range.Duplicate
    paragraphRange.MoveEnd unit:=wdCharacter, Count:=-1
    lineCount = paragraphRange.ComputeStatistics(wdStatisticLines)
    ParagraphHasMultipleLines = (lineCount > 1)
    Debug.Print "Paragraph line count: " & lineCount
End Function