Attribute VB_Name = "mod_RunAllChecks"
Option Explicit

' ==============================================================================
' Module:      RunAllChecks.bas
' Description: Master macro coordinating hyperlink removal and footnote citation
'              highlighting with aggregated statistical reporting.
' ==============================================================================

Public Sub RunAllChecks()
    Dim doc As Document
    Dim linksRemoved As Long
    Dim italicNew As Long, italicYellow As Long, italicSkipped As Long
    Dim scNew As Long, scYellow As Long, scSkipped As Long
    Dim confirmMsg As String
    
    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub
    
    confirmMsg = "This operation will:" & vbCrLf & _
                 "1. Remove all hyperlinks." & vbCrLf & _
                 "2. Highlight italic text in footnotes." & vbCrLf & _
                 "3. Highlight small-caps text in footnotes." & vbCrLf & vbCrLf & _
                 "Target document: " & doc.Name & vbCrLf & vbCrLf & _
                 "Save a backup before continuing." & vbCrLf & vbCrLf & _
                 "Do you wish to proceed?"
                 
    If MsgBox(confirmMsg, vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    SetAppState True
    
    ' 1. Remove Hyperlinks
    linksRemoved = ExecuteRemoveHyperlinks(doc)
    
    ' 2. Highlight Footnote Italics
    italicNew = ExecuteHighlightItalicFootnotes(doc, italicYellow, italicSkipped)
    
    ' 3. Highlight Footnote Small Caps
    scNew = ExecuteHighlightSmallCapsFootnotes(doc, scYellow, scSkipped)
    
CleanUp:
    SetAppState False
    MsgBox "Review complete for: " & doc.Name & vbCrLf & vbCrLf & _
           "Hyperlinks removed: " & linksRemoved & vbCrLf & _
           "Italic segments highlighted: " & italicNew & _
           " (Already yellow: " & italicYellow & ", Skipped: " & italicSkipped & ")" & vbCrLf & _
           "Small-caps segments highlighted: " & scNew & _
           " (Already yellow: " & scYellow & ", Skipped: " & scSkipped & ")", _
           vbInformation, APP_TITLE
    Exit Sub

ErrorHandler:
    SetAppState False
    MsgBox "Execution failed on '" & doc.Name & "': " & Err.Description & " (Error " & Err.Number & ")", _
           vbCritical, APP_TITLE
End Sub