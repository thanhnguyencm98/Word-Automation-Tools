Attribute VB_Name = "mod_CommonUtilities"
Option Explicit

' ==============================================================================
' Module:      CommonUtilities.bas
' Description: Shared helper routines, document safety checks, and application
'              state management for Word Citation Review Automation tools.
' ==============================================================================

Public Const APP_TITLE As String = "Word Citation Review Automation"
Public Const TARGET_HIGHLIGHT As Long = wdYellow

Private m_savedScreenUpdating As Boolean
Private m_savedAlerts As WdAlertLevel
Private m_stateSaved As Boolean

' Call before modifying the target document. Suppresses screen updates and alerts,
' and groups every change into a single undo step named actionName (Word 2010+).
' The target document must be the active document.
Public Sub BeginDocumentChanges(ByVal actionName As String)
    On Error Resume Next
    If Not m_stateSaved Then
        m_savedScreenUpdating = Application.ScreenUpdating
        m_savedAlerts = Application.DisplayAlerts
        m_stateSaved = True
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.UndoRecord.StartCustomRecord Left$(actionName, 64)
    On Error GoTo 0
End Sub

' Call after BeginDocumentChanges, on both the success and the error path.
' Clears the Err object, so read Err.Number/Err.Description before calling it.
Public Sub EndDocumentChanges()
    On Error Resume Next
    If Application.UndoRecord.IsRecordingCustomRecord Then
        Application.UndoRecord.EndCustomRecord
    End If
    If m_stateSaved Then
        Application.ScreenUpdating = m_savedScreenUpdating
        Application.DisplayAlerts = m_savedAlerts
        m_stateSaved = False
    End If
    Application.ScreenRefresh
    On Error GoTo 0
End Sub

Public Function GetTargetDocument() As Document
    Dim activeDoc As Document
    Dim docItem As Document
    Dim candidateDoc As Document
    Dim otherDocsCount As Long
    Dim docArray() As Document
    Dim docListPrompt As String
    Dim userInput As String
    Dim selectedIndex As Long
    Dim docIndex As Long

    If Application.Documents.Count = 0 Then
        MsgBox "No open documents found.", vbCritical, APP_TITLE
        Set GetTargetDocument = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set activeDoc = ActiveDocument
    On Error GoTo 0

    If activeDoc Is Nothing Then
        MsgBox "No active document found.", vbCritical, APP_TITLE
        Set GetTargetDocument = Nothing
        Exit Function
    End If

    ' If the active window is the macro container itself (WordAutomationTools.docm)
    If activeDoc Is ThisDocument Then
        otherDocsCount = 0
        For Each docItem In Application.Documents
            If Not (docItem Is ThisDocument) Then
                otherDocsCount = otherDocsCount + 1
                Set candidateDoc = docItem
            End If
        Next docItem

        If otherDocsCount = 0 Then
            MsgBox "The active window is currently the macro tool ('" & ThisDocument.Name & "')." & vbCrLf & vbCrLf & _
                   "Please open the target document you wish to review and try again.", _
                   vbExclamation, APP_TITLE
            Set GetTargetDocument = Nothing
            Exit Function
        ElseIf otherDocsCount = 1 Then
            If MsgBox("The active window is currently the macro tool ('" & ThisDocument.Name & "')." & vbCrLf & vbCrLf & _
                      "Do you want to process the open target document:" & vbCrLf & _
                      "'" & candidateDoc.Name & "'?", _
                      vbQuestion + vbYesNo, APP_TITLE) = vbYes Then
                candidateDoc.Activate
                Set activeDoc = candidateDoc
            Else
                Set GetTargetDocument = Nothing
                Exit Function
            End If
        Else
            ' Multiple open documents: display a numbered selection list
            ReDim docArray(1 To otherDocsCount)
            docIndex = 1
            docListPrompt = "Multiple target documents are open." & vbCrLf & _
                            "Enter the number corresponding to the document you want to process:" & vbCrLf & vbCrLf

            For Each docItem In Application.Documents
                If Not (docItem Is ThisDocument) Then
                    Set docArray(docIndex) = docItem
                    docListPrompt = docListPrompt & docIndex & ". " & docItem.Name & vbCrLf
                    docIndex = docIndex + 1
                End If
            Next docItem

            userInput = Trim$(InputBox(docListPrompt, APP_TITLE, "1"))
            If Len(userInput) = 0 Then
                Set GetTargetDocument = Nothing
                Exit Function
            End If

            ' Accept whole numbers only; the length limit also keeps CLng from overflowing.
            If userInput Like "*[!0-9]*" Or Len(userInput) > 4 Then
                MsgBox "Input must be a whole number from the list. Operation canceled.", vbExclamation, APP_TITLE
                Set GetTargetDocument = Nothing
                Exit Function
            End If

            selectedIndex = CLng(userInput)
            If selectedIndex < 1 Or selectedIndex > otherDocsCount Then
                MsgBox "Invalid selection number. Operation canceled.", vbExclamation, APP_TITLE
                Set GetTargetDocument = Nothing
                Exit Function
            End If

            Set candidateDoc = docArray(selectedIndex)
            candidateDoc.Activate
            Set activeDoc = candidateDoc
        End If
    End If

    If Not CanModifyDocument(activeDoc) Then
        Set GetTargetDocument = Nothing
        Exit Function
    End If

    Set GetTargetDocument = activeDoc
End Function

Public Function CanModifyDocument(ByVal doc As Document) As Boolean
    If doc Is Nothing Then
        MsgBox "No active document found.", vbCritical, APP_TITLE
        CanModifyDocument = False
        Exit Function
    End If

    If doc Is ThisDocument Then
        MsgBox "Cannot modify the macro container document ('" & doc.Name & "').", _
               vbExclamation, APP_TITLE
        CanModifyDocument = False
        Exit Function
    End If

    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "The active document is protected (" & doc.Name & "). Modifications cannot be made.", _
               vbExclamation, APP_TITLE
        CanModifyDocument = False
        Exit Function
    End If

    If doc.ReadOnly Then
        MsgBox "The active document is read-only (" & doc.Name & "). Modifications cannot be saved.", _
               vbExclamation, APP_TITLE
        CanModifyDocument = False
        Exit Function
    End If

    If doc.TrackRevisions Then
        If MsgBox("Track Changes is active in '" & doc.Name & "'." & vbCrLf & vbCrLf & _
                  "Every hyperlink removal and highlight will be recorded as a tracked revision." & vbCrLf & _
                  "To avoid this, choose No, turn off Track Changes, and run the macro again." & vbCrLf & vbCrLf & _
                  "Continue with Track Changes on?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, APP_TITLE) = vbNo Then
            CanModifyDocument = False
            Exit Function
        End If
    End If

    CanModifyDocument = True
End Function
