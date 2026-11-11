Option Explicit

' =============================================================================
' FILENAME:       FixSingleWordLines.vb
' PURPOSE:        Prevent a single final word from appearing alone on the last
'                 rendered line of a paragraph.
'
' DETECTION CHANGE FROM THE PREVIOUS VERSION
' ------------------------------------------
' The previous version used:
'
'     paragraphRange.ComputeStatistics(wdStatisticLines) > 1
'
' That test established only that a paragraph occupied more than one rendered
' line. It did not establish that the final rendered line contained only one
' word. As a result, any multi-line paragraph could be adjusted, even when its
' final line already contained two or more words.
'
' This version identifies the final two word-like runs in each paragraph and
' compares their actual rendered positions using:
'
'     Range.Information(wdActiveEndPageNumber)
'     Range.Information(wdFirstCharacterLineNumber)
'
' The separator before the final word is replaced with a nonbreaking space only
' when the final word and the preceding word appear on different rendered lines
' or different pages.
'
' EXAMPLE
' -------
' This paragraph is not changed because "several" and "words." are on the same
' rendered line:
'
'     This paragraph spans several lines but finishes with several
'     words.
'
' This paragraph is changed because "with" and "experience." are on different
' rendered lines:
'
'     This paragraph spans several lines and finishes with
'     experience.
'
' The regular space between "with" and "experience." becomes an NBSP.
'
' SAFEGUARDS
' ----------
' - Skips single-line endings where the final two words share a rendered line.
' - Skips paragraphs already adjusted with an NBSP.
' - Skips a final word placed after a manual line break.
' - Ignores paragraph markers, table-cell markers, and trailing whitespace.
' - Skips paragraphs ending with an inline object or unsupported control mark.
' - Skips punctuation-only final runs.
' - Replaces only the separator character, preserving surrounding formatting.
' - Repaginates the document before examining rendered line positions.
' - Reports completion statistics in a message box and the Immediate window.
'
' IMPORTANT
' ---------
' Because detection depends on Word's rendered layout, run the macro after the
' document's page size, margins, fonts, styles, and paragraph formatting are
' final. Print Layout view is recommended.
'
' This version performs only NBSP correction. It does not change paragraph
' indents, margins, font scaling, or character spacing. Bounded copyfitting can
' be added later without changing the detector.
' =============================================================================


Private Enum FixResult
    frNotEvaluated = 0
    frReplaced = 1
    frSameRenderedLine = 2
    frAlreadyAdjusted = 3
    frTooShort = 4
    frNoPrecedingWord = 5
    frIntentionalLineBreak = 6
    frUnsupportedEnding = 7
    frIneligibleSeparator = 8
    frError = 9
End Enum


Public Sub FixSingleWordLines()

    Dim para As Paragraph
    Dim separatorRange As Range
    Dim result As FixResult

    Dim totalParagraphs As Long
    Dim replacedCount As Long
    Dim sameLineCount As Long
    Dim alreadyAdjustedCount As Long
    Dim intentionalBreakCount As Long
    Dim ineligibleCount As Long
    Dim errorCount As Long

    Dim documentWasSaved As Boolean
    Dim message As String

    On Error GoTo FatalError

    documentWasSaved = ActiveDocument.Saved

    Application.ScreenUpdating = False
    Application.StatusBar = _
        "Checking paragraphs for hanging final words..."

    Debug.Print String(72, "=")
    Debug.Print "FixSingleWordLines started: " & Now
    Debug.Print "Document: " & ActiveDocument.Name

    ' Force Word to update page and line-layout information before the
    ' position comparisons begin.
    ActiveDocument.Repaginate

    For Each para In ActiveDocument.Paragraphs

        totalParagraphs = totalParagraphs + 1
        Set separatorRange = Nothing
        result = frNotEvaluated

        Application.StatusBar = _
            "Checking paragraph " & totalParagraphs & _
            " of " & ActiveDocument.Paragraphs.Count & "..."

        If GetHangingFinalWordSeparator( _
            para, separatorRange, result) Then

            On Error GoTo ParagraphError

            separatorRange.Text = ChrW(160)
            replacedCount = replacedCount + 1
            result = frReplaced

            Debug.Print _
                "REPLACED: " & GetDebugParagraphText(para)

            On Error GoTo FatalError

        Else

            Select Case result

                Case frSameRenderedLine
                    sameLineCount = sameLineCount + 1

                Case frAlreadyAdjusted
                    alreadyAdjustedCount = _
                        alreadyAdjustedCount + 1

                Case frIntentionalLineBreak
                    intentionalBreakCount = _
                        intentionalBreakCount + 1

                Case frTooShort, _
                     frNoPrecedingWord, _
                     frUnsupportedEnding, _
                     frIneligibleSeparator

                    ineligibleCount = ineligibleCount + 1

                Case frError
                    errorCount = errorCount + 1

            End Select

        End If

ContinueLoop:
        On Error GoTo FatalError

    Next para

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Debug.Print "Paragraphs examined:       " & totalParagraphs
    Debug.Print "Hanging words fixed:       " & replacedCount
    Debug.Print "Final words on same line:  " & sameLineCount
    Debug.Print "Already adjusted:          " & alreadyAdjustedCount
    Debug.Print "Intentional line breaks:   " & intentionalBreakCount
    Debug.Print "Ineligible paragraphs:     " & ineligibleCount
    Debug.Print "Errors:                    " & errorCount
    Debug.Print "FixSingleWordLines ended:  " & Now
    Debug.Print String(72, "=")

    message = _
        "Fix Single-Word Lines is complete." & vbCrLf & vbCrLf & _
        "Paragraphs examined: " & _
            Format$(totalParagraphs, "#,##0") & vbCrLf & _
        "Hanging words fixed: " & _
            Format$(replacedCount, "#,##0") & vbCrLf & _
        "Already adjusted: " & _
            Format$(alreadyAdjustedCount, "#,##0") & vbCrLf & _
        "Final two words already on the same line: " & _
            Format$(sameLineCount, "#,##0") & vbCrLf & _
        "Intentional line breaks skipped: " & _
            Format$(intentionalBreakCount, "#,##0") & vbCrLf & _
        "Other ineligible paragraphs: " & _
            Format$(ineligibleCount, "#,##0")

    If errorCount > 0 Then
        message = message & vbCrLf & _
            "Paragraph errors: " & _
            Format$(errorCount, "#,##0") & vbCrLf & vbCrLf & _
            "See the VBA Immediate window for details."
    End If

    If replacedCount = 0 And documentWasSaved Then
        ' Preserve Word's saved-state indicator when no document changes
        ' were made by this macro.
        ActiveDocument.Saved = True
    End If

    MsgBox _
        Prompt:=message, _
        Buttons:=vbInformation + vbOKOnly, _
        Title:="Fix Single-Word Lines"

    Exit Sub


ParagraphError:

    errorCount = errorCount + 1

    Debug.Print _
        "PARAGRAPH ERROR " & Err.Number & ": " & Err.Description

    Err.Clear
    Resume ContinueLoop


FatalError:

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox _
        Prompt:= _
            "Fix Single-Word Lines could not complete." & _
            vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, _
        Buttons:=vbExclamation + vbOKOnly, _
        Title:="Fix Single-Word Lines"

End Sub


Private Function GetHangingFinalWordSeparator( _
    ByVal para As Paragraph, _
    ByRef separatorRange As Range, _
    ByRef result As FixResult _
) As Boolean

    Dim contentRange As Range
    Dim paragraphStart As Long

    Dim i As Long
    Dim finalRunStartIndex As Long
    Dim finalRunEndIndex As Long
    Dim previousRunEndIndex As Long

    Dim separatorStartIndex As Long
    Dim separatorEndIndex As Long
    Dim replacementIndex As Long

    Dim currentCharacter As String
    Dim finalRunText As String

    Dim finalWordProbe As Range
    Dim previousWordProbe As Range

    Dim finalWordPage As Long
    Dim previousWordPage As Long
    Dim finalWordLine As Long
    Dim previousWordLine As Long

    On Error GoTo DetectionError

    GetHangingFinalWordSeparator = False
    Set separatorRange = Nothing
    result = frNotEvaluated

    Set contentRange = para.Range.Duplicate
    TrimRangeTerminators contentRange

    If contentRange.Characters.Count < 3 Then
        result = frTooShort
        Exit Function
    End If

    paragraphStart = contentRange.Start
    i = contentRange.Characters.Count

    ' Ignore trailing spaces, tabs, and NBSPs. A manual line break is not
    ' removed because it represents an intentional rendered-line boundary.
    Do While i >= 1

        currentCharacter = contentRange.Characters(i).Text

        If IsTrailingWhitespace(currentCharacter) Then
            i = i - 1
        Else
            Exit Do
        End If

    Loop

    If i < 1 Then
        result = frTooShort
        Exit Function
    End If

    finalRunEndIndex = i

    ' Do not treat an inline picture, embedded object, or unsupported control
    ' character at the end of a paragraph as a final word.
    If IsUnsupportedObjectCharacter( _
        contentRange.Characters(finalRunEndIndex).Text) Then

        result = frUnsupportedEnding
        Exit Function
    End If

    ' Move backward over the final non-whitespace run.
    Do While i >= 1

        currentCharacter = contentRange.Characters(i).Text

        If IsRunSeparator(currentCharacter) Then
            Exit Do
        End If

        If IsUnsupportedObjectCharacter(currentCharacter) Then
            result = frUnsupportedEnding
            Exit Function
        End If

        i = i - 1

    Loop

    finalRunStartIndex = i + 1

    If finalRunStartIndex <= 1 Then
        result = frNoPrecedingWord
        Exit Function
    End If

    finalRunText = GetCharacterSlice( _
        contentRange, _
        finalRunStartIndex, _
        finalRunEndIndex)

    ' Avoid treating an isolated punctuation mark as a word.
    If Not ContainsWordCharacter(finalRunText) Then
        result = frUnsupportedEnding
        Exit Function
    End If

    separatorEndIndex = finalRunStartIndex - 1
    i = separatorEndIndex

    ' Examine the complete whitespace run between the final two text runs.
    Do While i >= 1

        currentCharacter = contentRange.Characters(i).Text

        If currentCharacter = ChrW(160) Then
            result = frAlreadyAdjusted
            Exit Function
        End If

        If IsManualLineBreak(currentCharacter) Then
            result = frIntentionalLineBreak
            Exit Function
        End If

        If currentCharacter = " " Or currentCharacter = vbTab Then
            i = i - 1
        Else
            Exit Do
        End If

    Loop

    separatorStartIndex = i + 1
    previousRunEndIndex = i

    If previousRunEndIndex < 1 Then
        result = frNoPrecedingWord
        Exit Function
    End If

    ' Replace the whitespace character immediately before the final word.
    ' This matches the prior macro while preserving all surrounding text.
    replacementIndex = separatorEndIndex
    currentCharacter = _
        contentRange.Characters(replacementIndex).Text

    If currentCharacter <> " " And _
       currentCharacter <> vbTab Then

        result = frIneligibleSeparator
        Exit Function
    End If

    If IsUnsupportedObjectCharacter( _
        contentRange.Characters(previousRunEndIndex).Text) Then

        result = frUnsupportedEnding
        Exit Function
    End If

    Set previousWordProbe = contentRange.Duplicate
    previousWordProbe.SetRange _
        Start:=paragraphStart + previousRunEndIndex - 1, _
        End:=paragraphStart + previousRunEndIndex

    Set finalWordProbe = contentRange.Duplicate
    finalWordProbe.SetRange _
        Start:=paragraphStart + finalRunStartIndex - 1, _
        End:=paragraphStart + finalRunStartIndex

    previousWordPage = previousWordProbe.Information( _
        wdActiveEndPageNumber)

    finalWordPage = finalWordProbe.Information( _
        wdActiveEndPageNumber)

    previousWordLine = previousWordProbe.Information( _
        wdFirstCharacterLineNumber)

    finalWordLine = finalWordProbe.Information( _
        wdFirstCharacterLineNumber)

    Debug.Print _
        "POSITION: previous=P" & previousWordPage & _
        "/L" & previousWordLine & _
        "; final=P" & finalWordPage & _
        "/L" & finalWordLine & _
        "; text=" & GetDebugParagraphText(para)

    If finalWordPage <> previousWordPage Or _
       finalWordLine <> previousWordLine Then

        Set separatorRange = contentRange.Duplicate
        separatorRange.SetRange _
            Start:=paragraphStart + replacementIndex - 1, _
            End:=paragraphStart + replacementIndex

        GetHangingFinalWordSeparator = True

    Else

        result = frSameRenderedLine

    End If

    Exit Function


DetectionError:

    Debug.Print _
        "DETECTION ERROR " & Err.Number & _
        ": " & Err.Description & _
        "; text=" & GetDebugParagraphText(para)

    result = frError
    Err.Clear

End Function


Private Sub TrimRangeTerminators(ByRef targetRange As Range)

    Dim finalCharacter As String

    Do While targetRange.Characters.Count > 0

        finalCharacter = _
            targetRange.Characters( _
                targetRange.Characters.Count).Text

        Select Case finalCharacter

            Case vbCr, vbLf, ChrW(7)
                targetRange.MoveEnd _
                    Unit:=wdCharacter, _
                    Count:=-1

            Case Else
                Exit Do

        End Select

    Loop

End Sub


Private Function IsTrailingWhitespace( _
    ByVal character As String _
) As Boolean

    IsTrailingWhitespace = _
        character = " " Or _
        character = vbTab Or _
        character = ChrW(160)

End Function


Private Function IsRunSeparator( _
    ByVal character As String _
) As Boolean

    IsRunSeparator = _
        character = " " Or _
        character = vbTab Or _
        character = ChrW(160) Or _
        IsManualLineBreak(character)

End Function


Private Function IsManualLineBreak( _
    ByVal character As String _
) As Boolean

    ' Word represents a manual line break inserted with Shift+Enter as
    ' character 11.
    IsManualLineBreak = _
        character = ChrW(11)

End Function


Private Function IsUnsupportedObjectCharacter( _
    ByVal character As String _
) As Boolean

    Dim characterCode As Long

    If Len(character) = 0 Then
        IsUnsupportedObjectCharacter = True
        Exit Function
    End If

    characterCode = AscW(Left$(character, 1))

    ' Character 1 commonly represents an inline picture or embedded object
    ' in Word's document text.
    IsUnsupportedObjectCharacter = _
        characterCode = 1 Or _
        characterCode = 2 Or _
        characterCode = 8

End Function


Private Function GetCharacterSlice( _
    ByVal sourceRange As Range, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long _
) As String

    Dim sliceRange As Range

    Set sliceRange = sourceRange.Duplicate

    sliceRange.SetRange _
        Start:=sourceRange.Start + firstIndex - 1, _
        End:=sourceRange.Start + lastIndex

    GetCharacterSlice = sliceRange.Text

End Function


Private Function ContainsWordCharacter( _
    ByVal text As String _
) As Boolean

    Dim i As Long
    Dim character As String
    Dim codePoint As Long

    For i = 1 To Len(text)

        character = Mid$(text, i, 1)
        codePoint = AscW(character)

        ' Include ASCII letters and numbers plus non-ASCII Unicode
        ' characters. This covers ordinary English resume text and names
        ' containing accented or other international characters.
        If character Like "[A-Za-z0-9]" Or _
           codePoint > 127 Then

            ContainsWordCharacter = True
            Exit Function

        End If

    Next i

    ContainsWordCharacter = False

End Function


Private Function GetDebugParagraphText( _
    ByVal para As Paragraph _
) As String

    Dim text As String

    On Error GoTo DebugTextError

    text = para.Range.Text
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(7), "")
    text = Replace(text, ChrW(11), "[LINE BREAK]")
    text = Replace(text, ChrW(160), "[NBSP]")
    text = Replace(text, ChrW(1), "[OBJECT]")

    If Len(text) > 160 Then
        text = Left$(text, 157) & "..."
    End If

    GetDebugParagraphText = text
    Exit Function

DebugTextError:
    GetDebugParagraphText = "[Unable to read paragraph text]"

End Function