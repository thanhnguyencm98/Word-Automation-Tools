Attribute VB_Name = "mod_RunAllChecks"
Option Explicit

' ==============================================================================
' Module:      RunAllChecks.bas
' Description: Master macro coordinating hyperlink removal and footnote citation
'              highlighting with aggregated statistical reporting.
' ==============================================================================

Public Sub RunAllChecks()
    Dim doc As Document
    Dim linksRemoved As Long, tocLinksKept As Long
    Dim italicNew As Long, italicYellow As Long, italicSkipped As Long
    Dim scNew As Long, scYellow As Long, scSkipped As Long
    Dim confirmMsg As String
    Dim currentStep As String
    Dim errNumber As Long, errDescription As String

    Set doc = GetTargetDocument()
    If doc Is Nothing Then Exit Sub

    confirmMsg = "This operation will:" & vbCrLf & _
                 "1. Remove all hyperlinks (table of contents links are kept)." & vbCrLf & _
                 "2. Highlight italic text in footnotes." & vbCrLf & _
                 "3. Highlight small-caps text in footnotes." & vbCrLf & vbCrLf & _
                 "Target document: " & doc.Name & vbCrLf & vbCrLf & _
                 "Save a backup before continuing." & vbCrLf & vbCrLf & _
                 "Do you wish to proceed?"

    If MsgBox(confirmMsg, vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    BeginDocumentChanges "Run All Citation Checks"

    currentStep = "1. Remove hyperlinks"
    linksRemoved = ExecuteRemoveHyperlinks(doc, tocLinksKept)

    currentStep = "2. Highlight footnote italics"
    italicNew = HighlightFootnoteFormatting(doc, ffItalic, italicYellow, italicSkipped)

    currentStep = "3. Highlight footnote small caps"
    scNew = HighlightFootnoteFormatting(doc, ffSmallCaps, scYellow, scSkipped)

    EndDocumentChanges

    MsgBox "Review complete for: " & doc.Name & vbCrLf & vbCrLf & _
           "Hyperlinks removed: " & linksRemoved & _
           " (Table of contents links kept: " & tocLinksKept & ")" & vbCrLf & _
           "Italic segments highlighted: " & italicNew & _
           " (Already yellow: " & italicYellow & ", Skipped: " & italicSkipped & ")" & vbCrLf & _
           "Small-caps segments highlighted: " & scNew & _
           " (Already yellow: " & scYellow & ", Skipped: " & scSkipped & ")" & vbCrLf & vbCrLf & _
           "Press Ctrl+Z once to undo the whole run.", _
           vbInformation, APP_TITLE
    Exit Sub

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description
    EndDocumentChanges
    MsgBox "Execution failed on '" & doc.Name & "' at step " & currentStep & ":" & vbCrLf & _
           errDescription & " (Error " & errNumber & ")" & vbCrLf & vbCrLf & _
           "Changes from earlier steps were already applied." & vbCrLf & _
           "Press Ctrl+Z once to undo the whole run.", _
           vbCritical, APP_TITLE
End Sub
