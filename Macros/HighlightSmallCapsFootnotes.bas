Attribute VB_Name = "mod_HighlightSmallCapsFootnotes"
Option Explicit

' ==============================================================================
' Module:      HighlightSmallCapsFootnotes.bas
' Description: Detects and highlights true Small Caps font property in footnotes
'              without flagging plain capital text.
' ==============================================================================

Public Function ExecuteHighlightSmallCapsFootnotes(ByVal doc As Document, _
                                                  ByRef outAlreadyYellow As Long, _
                                                  ByRef outSkippedOther As Long) As Long
    Dim fn As Footnote
    Dim searchRange As Range
    Dim fnEnd As Long
    Dim prevStart As Long
    Dim newlyHighlighted As Long
    
    newlyHighlighted = 0
    outAlreadyYellow = 0
    outSkippedOther = 0
    
    For Each fn In doc.Footnotes
        Set searchRange = fn.Range.Duplicate
        fnEnd = fn.Range.End
        
        With searchRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Font.SmallCaps = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            
            Do While .Execute
                If searchRange.End > fnEnd Or searchRange.Start >= fnEnd Then Exit Do
                
                If Len(Trim(searchRange.Text)) > 0 Then
                    Select Case searchRange.HighlightColorIndex
                        Case wdNoHighlight
                            searchRange.HighlightColorIndex = TARGET_HIGHLIGHT
                            newlyHighlighted = newlyHighlighted + 1
                        Case TARGET_HIGHLIGHT
                            outAlreadyYellow = outAlreadyYellow + 1
                        Case Else
                            outSkippedOther = outSkippedOther + 1
                    End Select
                End If
                
                prevStart = searchRange.Start
                searchRange.Collapse wdCollapseEnd
                If searchRange.Start <= prevStart Then
                    searchRange.Start = prevStart + 1
                    searchRange.End = searchRange.Start
                End If
                If searchRange.Start >= fnEnd Then Exit Do
            Loop
        End With
    Next fn
    
    ExecuteHighlightSmallCapsFootnotes = newlyHighlighted
End Function

Public Sub HighlightSmallCapsFootnotes()
    Dim doc As Document
    Dim newCount As Long, existingYellow As Long, skippedOther As Long
    
    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub
    If doc.Footnotes.Count = 0 Then
        MsgBox "No footnotes found in '" & doc.Name & "'.", vbInformation, APP_TITLE
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    SetAppState True
    newCount = ExecuteHighlightSmallCapsFootnotes(doc, existingYellow, skippedOther)
CleanUp:
    SetAppState False
    MsgBox "Footnote Small Caps review for '" & doc.Name & "':" & vbCrLf & vbCrLf & _
           "Small-caps segments highlighted: " & newCount & vbCrLf & _
           "Already yellow: " & existingYellow & vbCrLf & _
           "Skipped (other highlight): " & skippedOther, vbInformation, APP_TITLE
    Exit Sub
ErrorHandler:
    SetAppState False
    MsgBox "An error occurred: " & Err.Description, vbCritical, APP_TITLE
End Sub