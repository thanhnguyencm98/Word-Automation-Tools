Attribute VB_Name = "mod_RemoveAllHyperlinks"
Option Explicit

' ==============================================================================
' Module:      RemoveAllHyperlinks.bas
' Description: Removes all hyperlinks across all story ranges and linked chains
'              while strictly preserving visible display text and unrelated fields.
'              Table of contents and table of figures entry links are kept.
' ==============================================================================

Public Function ExecuteRemoveHyperlinks(ByVal doc As Document, ByRef outTocLinksKept As Long) As Long
    Dim currentStory As Range
    Dim linkedStory As Range
    Dim link As Hyperlink
    Dim removedCount As Long
    Dim i As Long

    removedCount = 0
    outTocLinksKept = 0

    For Each currentStory In doc.StoryRanges
        Set linkedStory = currentStory
        Do While Not (linkedStory Is Nothing)
            For i = linkedStory.Hyperlinks.Count To 1 Step -1
                Set link = linkedStory.Hyperlinks(i)
                If IsTableOfContentsLink(link) Then
                    outTocLinksKept = outTocLinksKept + 1
                Else
                    link.Delete
                    removedCount = removedCount + 1
                End If
            Next i
            Set linkedStory = linkedStory.NextStoryRange
        Loop
    Next currentStory

    ExecuteRemoveHyperlinks = removedCount
End Function

Public Sub RemoveAllHyperlinks()
    Dim doc As Document
    Dim removedCount As Long
    Dim tocLinksKept As Long
    Dim summary As String
    Dim errNumber As Long, errDescription As String

    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub

    If MsgBox("Are you sure you want to remove all hyperlinks from '" & doc.Name & "'?" & vbCrLf & _
              "Visible text will be preserved, but the links will be stripped." & vbCrLf & _
              "Table of contents links are kept." & vbCrLf & vbCrLf & _
              "Ensure you have a backup copy before proceeding.", _
              vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    BeginDocumentChanges "Remove All Hyperlinks"
    removedCount = ExecuteRemoveHyperlinks(doc, tocLinksKept)
    EndDocumentChanges

    summary = "Hyperlinks removed from '" & doc.Name & "': " & removedCount
    If tocLinksKept > 0 Then
        summary = summary & vbCrLf & "Table of contents links kept: " & tocLinksKept
    End If
    MsgBox summary & vbCrLf & vbCrLf & "Press Ctrl+Z once to undo this whole operation.", _
           vbInformation, APP_TITLE
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    EndDocumentChanges
    MsgBox "An error occurred during hyperlink removal: " & errDescription & " (Error " & errNumber & ")" & vbCrLf & vbCrLf & _
           "Press Ctrl+Z once to restore any hyperlinks already removed.", vbCritical, APP_TITLE
End Sub

' Table of contents and table of figures entries are generated links to hidden
' "_Toc" bookmarks in the same document. Removing them would break navigation
' from the table until it is updated.
Private Function IsTableOfContentsLink(ByVal link As Hyperlink) As Boolean
    On Error Resume Next
    IsTableOfContentsLink = (Len(link.Address) = 0 And LCase$(Left$(link.SubAddress, 4)) = "_toc")
End Function
