Sub FixSingleWordLines_macOS()
    Dim para As Paragraph
    Dim searchRange As Range
    Dim textLine As String
    Dim pos As Long
    Dim spacePos As Long
    Dim lastWord As String
    Dim trimmedText As String

    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.Text) > 2 Then
            If GetParagraphLineCount(para) <> 1 Then
                Set searchRange = para.Range
                searchRange.MoveEnd unit:=wdCharacter, Count:=-1
                textLine = searchRange.Text

                ' Strip trailing spaces, nbsps, and tabs
                trimmedText = RTrim(textLine, " " & ChrW(160) & vbTab)

                ' Find last word using string manipulation
                pos = InStrRev(trimmedText, " ")
                If pos > 0 Then
                    lastWord = Mid(trimmedText, pos + 1)
                    spacePos = pos

                    ' Debug Print
                    Debug.Print "Space Position: " & spacePos
                    Debug.Print "Last Word: " & lastWord

                    ' Replace the space character at the specified position
                    If spacePos >= 0 Then
                        searchRange.Start = searchRange.Start + spacePos - 1
                        searchRange.End = searchRange.Start + 1
                        searchRange.Text = ChrW(160)
                    End If
                End If
            End If
        End If
    Next para
End Sub

Function RTrim(ByVal str As String, ByVal chars As String) As String
    Dim i As Integer
    For i = Len(str) To 1 Step -1
        If InStr(chars, Mid(str, i, 1)) = 0 Then Exit For
    Next i
    RTrim = Left(str, i)
End Function

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
