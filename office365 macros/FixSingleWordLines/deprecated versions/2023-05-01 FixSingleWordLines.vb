' CHANGELOG
' - works with lines at the end of a page (where line count is negative)
'   - this isn't perfect; it will add an NBSP even if the para is a single line;
'       I don't imagine this becoming an issue, but can likely fix by
'       adjusting the GetParagraphLineCount function to deal with negative line counts
' - ignores paragraphs that have already been adjusted_manually
' - ignores paragraphs that have multiple_final_words_strung_together_with_NBSPs
' - prints to console for better debugging

Sub FixSingleWordLines()
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Text = "^w^p"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With

    Dim regex As Object
    Dim para As Paragraph
    Dim searchRange As Range
    Dim foundMatch As Boolean
    Dim match As Object

    ' Create a new regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
    End With

    ' Loop through each paragraph in the active document
    For Each para In ActiveDocument.Paragraphs
        ' Skip empty paragraphs
        If Len(para.Range.Text) > 2 Then
            ' Check if the paragraph is longer than 1 line
            If GetParagraphLineCount(para) <> 1 Then
                ' Set the search range to the current paragraph
                Set searchRange = para.Range

                ' Set the regex pattern for testing (accounts for 2+ words with NBSPs)
                regex.Pattern = "\s(?:\w+\xA0)+\w+(\.\r)"

                ' Check if the paragraph contains a match of the regular expression
                foundMatch = regex.Test(searchRange.Text)
                ' Print a message to the console indicating whether the pattern was found or not
                If foundMatch Then
                    Debug.Print "Paragraph already adjusted: " & para.Range.Text
                Else
                    ' Pattern not found
                    Debug.Print "Pattern not found; let's add an NBSP: " & para.Range.Text

                    ' Update the regex pattern for the final space replacement
                    regex.Pattern = "(\s)(\S+\s*(?: \s*)*)$"

                    Set match = regex.Execute(para.Range.Text)
                    If match.Count > 0 Then
                        para.Range.Find.Execute FindText:=match.Item(0).SubMatches(0) & match.Item(0).SubMatches(1), _
                                                ReplaceWith:=ChrW(160) & match.Item(0).SubMatches(1), _
                                                Replace:=wdReplaceOne
                    End If
                End If
            Else
            ' Do nothing and move to the next paragraph
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
