Attribute VB_Name = "mod_RemoveAllHyperlinks"
Option Explicit

' ==============================================================================
' Module:      RemoveAllHyperlinks.bas
' Description: Removes all hyperlinks across all story ranges and linked chains
'              while strictly preserving visible display text and unrelated fields.
' ==============================================================================

Public Function ExecuteRemoveHyperlinks(ByVal doc As Document) As Long
    Dim currentStory As Range
    Dim linkedStory As Range
    Dim removedCount As Long
    Dim i As Long
    
    removedCount = 0
    
    For Each currentStory In doc.StoryRanges
        Set linkedStory = currentStory
        Do While Not (linkedStory Is Nothing)
            If linkedStory.Hyperlinks.Count > 0 Then
                For i = linkedStory.Hyperlinks.Count To 1 Step -1
                    linkedStory.Hyperlinks(i).Delete
                    removedCount = removedCount + 1
                Next i
            End If
            Set linkedStory = linkedStory.NextStoryRange
        Loop
    Next currentStory
    
    ExecuteRemoveHyperlinks = removedCount
End Function

Public Sub RemoveAllHyperlinks()
    Dim doc As Document
    Dim removedCount As Long
    
    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub
    
    If MsgBox("Are you sure you want to remove all hyperlinks from '" & doc.Name & "'?" & vbCrLf & _
              "Visible text will be preserved, but web links will be stripped." & vbCrLf & vbCrLf & _
              "Ensure you have a backup copy before proceeding.", _
              vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    SetAppState True
    removedCount = ExecuteRemoveHyperlinks(doc)
    
CleanUp:
    SetAppState False
    MsgBox "Hyperlinks removed from '" & doc.Name & "': " & removedCount, vbInformation, APP_TITLE
    Exit Sub
ErrorHandler:
    SetAppState False
    MsgBox "An error occurred during hyperlink removal: " & Err.Description, vbCritical, APP_TITLE
End Sub