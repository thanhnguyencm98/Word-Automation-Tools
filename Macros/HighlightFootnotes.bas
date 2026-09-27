Attribute VB_Name = "mod_HighlightFootnotes"
Option Explicit

' ==============================================================================
' Module:      HighlightFootnotes.bas
' Description: Highlights italic and true Small Caps text segments in footnotes
'              with yellow while respecting non-yellow highlights and
'              non-footnote scopes. Plain capital text is not flagged.
' ==============================================================================

Public Enum FootnoteFormat
    ffItalic = 1
    ffSmallCaps = 2
End Enum

Public Sub HighlightItalicFootnotes()
    RunFootnoteHighlight ffItalic
End Sub

Public Sub HighlightSmallCapsFootnotes()
    RunFootnoteHighlight ffSmallCaps
End Sub

Public Function HighlightFootnoteFormatting(ByVal doc As Document, _
                                            ByVal formatCheck As FootnoteFormat, _
                                            ByRef outAlreadyYellow As Long, _
                                            ByRef outSkippedOther As Long) As Long
    Dim fn As Footnote
    Dim searchRange As Range
    Dim fnEnd As Long
    Dim searchFrom As Long
    Dim nextStart As Long
    Dim hadOtherHighlight As Boolean
    Dim newlyHighlighted As Long

    newlyHighlighted = 0
    outAlreadyYellow = 0
    outSkippedOther = 0

    For Each fn In doc.Footnotes
        Set searchRange = fn.Range.Duplicate
        fnEnd = searchRange.End

        With searchRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            If formatCheck = ffItalic Then
                .Font.Italic = True
            Else
                .Font.SmallCaps = True
            End If
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False

            ' searchRange always covers only the unsearched rest of this footnote.
            ' Never execute on a collapsed range: Word would search on to the end
            ' of the footnotes story.
            Do While searchRange.Start < fnEnd
                searchFrom = searchRange.Start
                If Not .Execute Then Exit Do
                If searchRange.Start >= fnEnd Then Exit Do

                ' A formatted run can continue into the next footnote; keep this footnote's part.
                If searchRange.End > fnEnd Then searchRange.End = fnEnd

                If HasVisibleText(searchRange.Text) Then
                    Select Case searchRange.HighlightColorIndex
                        Case wdNoHighlight
                            searchRange.HighlightColorIndex = TARGET_HIGHLIGHT
                            newlyHighlighted = newlyHighlighted + 1
                        Case TARGET_HIGHLIGHT
                            outAlreadyYellow = outAlreadyYellow + 1
                        Case wdUndefined
                            ' Mixed highlighting: fill in only the unhighlighted parts.
                            If HighlightUnmarkedParts(searchRange, hadOtherHighlight) Then
                                newlyHighlighted = newlyHighlighted + 1
                            ElseIf hadOtherHighlight Then
                                outSkippedOther = outSkippedOther + 1
                            Else
                                outAlreadyYellow = outAlreadyYellow + 1
                            End If
                        Case Else
                            outSkippedOther = outSkippedOther + 1
                    End Select
                End If

                nextStart = searchRange.End
                If nextStart <= searchFrom Then nextStart = searchFrom + 1
                If nextStart >= fnEnd Then Exit Do
                searchRange.SetRange nextStart, fnEnd
            Loop
        End With
    Next fn

    HighlightFootnoteFormatting = newlyHighlighted
End Function

Private Sub RunFootnoteHighlight(ByVal formatCheck As FootnoteFormat)
    Dim doc As Document
    Dim reviewName As String
    Dim segmentName As String
    Dim newCount As Long, existingYellow As Long, skippedOther As Long
    Dim errNumber As Long, errDescription As String

    If formatCheck = ffItalic Then
        reviewName = "Italic"
        segmentName = "Italic"
    Else
        reviewName = "Small Caps"
        segmentName = "Small-caps"
    End If

    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub
    If doc.Footnotes.Count = 0 Then
        MsgBox "No footnotes found in '" & doc.Name & "'.", vbInformation, APP_TITLE
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    BeginDocumentChanges "Highlight Footnote " & reviewName
    newCount = HighlightFootnoteFormatting(doc, formatCheck, existingYellow, skippedOther)
    EndDocumentChanges

    MsgBox "Footnote " & reviewName & " review for '" & doc.Name & "':" & vbCrLf & vbCrLf & _
           segmentName & " segments highlighted: " & newCount & vbCrLf & _
           "Already yellow: " & existingYellow & vbCrLf & _
           "Skipped (other highlight): " & skippedOther & vbCrLf & vbCrLf & _
           "Press Ctrl+Z once to undo this whole review.", vbInformation, APP_TITLE
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    EndDocumentChanges
    MsgBox "An error occurred: " & errDescription & " (Error " & errNumber & ")" & vbCrLf & vbCrLf & _
           "Press Ctrl+Z once to undo any highlights already added.", vbCritical, APP_TITLE
End Sub

' Highlights the unhighlighted stretches of a range whose highlighting is mixed,
' leaving existing yellow and other colors untouched. Returns True when anything
' was highlighted; outHadOther reports whether a non-yellow highlight was seen.
Private Function HighlightUnmarkedParts(ByVal target As Range, ByRef outHadOther As Boolean) As Boolean
    Dim probe As Range
    Dim pos As Long
    Dim gapStart As Long
    Dim colorIndex As Long

    outHadOther = False
    gapStart = -1
    Set probe = target.Duplicate

    For pos = target.Start To target.End - 1
        probe.SetRange pos, pos + 1
        colorIndex = probe.HighlightColorIndex
        If colorIndex = wdNoHighlight Then
            If gapStart < 0 Then gapStart = pos
        Else
            If colorIndex <> TARGET_HIGHLIGHT Then outHadOther = True
            If gapStart >= 0 Then
                probe.SetRange gapStart, pos
                probe.HighlightColorIndex = TARGET_HIGHLIGHT
                HighlightUnmarkedParts = True
                gapStart = -1
            End If
        End If
    Next pos

    If gapStart >= 0 Then
        probe.SetRange gapStart, target.End
        probe.HighlightColorIndex = TARGET_HIGHLIGHT
        HighlightUnmarkedParts = True
    End If
End Function

' True when segmentText contains anything besides spaces, tabs, line/paragraph/page
' breaks, non-breaking spaces, and cell markers.
Private Function HasVisibleText(ByVal segmentText As String) As Boolean
    Dim i As Long

    For i = 1 To Len(segmentText)
        Select Case AscW(Mid$(segmentText, i, 1))
            Case 7, 9, 10, 11, 12, 13, 32, 160
            Case Else
                HasVisibleText = True
                Exit Function
        End Select
    Next i
End Function
