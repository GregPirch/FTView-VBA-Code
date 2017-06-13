
Dim WithEvents oGroup As TagGroup

'------------------------------------------------------

Sub SetUpTagGroup()

    On Error Resume Next
    Err.Clear

    If oGroup Is Nothing Then
        Set oGroup = Application.CreateTagGroup(Me.AreaName, 500)
        If Err.Number Then
            LogDiagnosticsMessage "Error creating TagGroup. Error: " & Err.Description, ftDiagSeverityError
            Exit Sub
        End If
        oGroup.Add "System\Second"
        oGroup.Add "System\Minute"
        oGroup.Add "Batch_Status"
        oGroup.Active = True
    End If

End Sub
    
'------------------------------------------------------
    
Sub SetTagValue_Second()

    On Error Resume Next
    Dim oTag As Tag
    If Not oGroup Is Nothing Then
        Set oTag = oGroup.Item("System\Second")
        Err.Clear
        oTag.Value = 10
        ' Test the Error number for the result.
        Select Case Err.Number

        Case 0:
            ' Write completed successfully... log a message
            LogDiagnosticsMessage "Write to tag " & oTag.Name & " was successful."
            MsgBox "write to seconds successful"
        Case tagErrorReadOnlyAccess:
            MsgBox "Unable to write tag value. Client is read-only."
        Case tagErrorWriteValue:
            If oTag.LastErrorNumber = tagErrorInvalidSecurity Then
                MsgBox "Unable to write tag value. The current user does not have security rights."
            Else
                MsgBox "Error writing tag value. Error: " & oTag.LastErrorString
            End If
        Case tagErrorOperationFailed:
            MsgBox "Failed to write to seconds tag. Error: " & Err.Description
        End Select

    End If

End Sub
    
'------------------------------------------------------
    
Sub SetTagValue_Batch_Status()

    On Error Resume Next
    Dim oTag As Tag
    If Not oGroup Is Nothing Then
        Set oTag = oGroup.Item("Batch_Status")
        Err.Clear
        oTag.Value = "TestGreg"
        ' Test the Error number for the result.
        Select Case Err.Number

        Case 0:
            ' Write completed successfully... log a message
            LogDiagnosticsMessage "Write to tag " & oTag.Name & " was successful."
            MsgBox "write to batch status successful"
        Case tagErrorReadOnlyAccess:
            MsgBox "Unable to write tag value. Client is read-only."
        Case tagErrorWriteValue:
            If oTag.LastErrorNumber = tagErrorInvalidSecurity Then
                MsgBox "Unable to write tag value. The current user does not have security rights."
            Else
                MsgBox "Error writing tag value. Error: " & oTag.LastErrorString
            End If
        Case tagErrorOperationFailed:
            MsgBox "Failed to write to " & oTag.Name & " batch status tag. Error: " & Err.Description
        End Select

    End If

End Sub
'------------------------------------------------------

Private Sub Button_GetBatchState_Released()

    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim BatchCreateID As String
    Dim BatchState As String
    Dim currentBatchNum As String
    
'    If BatchSev.GetItem("BatchListCt") <> "0" Then
'        BatchCreateID = BatchSev.GetItem("BLCreateID_1")
'        BatchState = BatchSev.GetItem("BLState_1")
'        MsgBox (BatchState)
'        'MsgBox ("START Batch Status: " & BatchResult & " (Create ID: " & BatchCreateID & ")")
'    Else
'        MsgBox ("Cannot get Batch STATE. No Batches are Created!")
'    End If


'Public WithEvents MyGroup As TagGroup
'Public oTag As Tag

'Private Sub Display_AnimationStart()
'Set MyGroup = Application.CreateTagGroup(Me.AreaName)
'MyGroup.Add ("Batch_Status")
'Set oTag = MyGroup.Item("Batch_Status")
'oTag.Value = 0
'End Sub

    Call SetUpTagGroup
   
    Call SetTagValue_Batch_Status
    Call SetTagValue_Second
    
End Sub
    
'------------------------------------------------------


