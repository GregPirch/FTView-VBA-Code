
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


Private Sub Button_CreateBatch_Released()

    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim currentBatchNum As String
    
    If BatchSev.GetItem("BatchListCt") = "0" Then
    
        BatchResult = BatchSev.Execute("BATCH(Item,ADV2\Admin,BULK_COMPLEX_BATCH.UOP,BATCH_ID,100,None,PARMS)")
        MsgBox ("CREATE Batch Status and Batch Number: " & BatchResult)
        currentBatchNum = CInt(Replace(Item1, "SUCCESS:", ""))
    Else
        MsgBox ("Cannot ADD Another Batch. A Batch Already Exist!")
    End If
    
End Sub

Private Sub Button_StartBatch_Released()

    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim BatchCreateID As String
    Dim currentBatchNum As String
    
    If BatchSev.GetItem("BatchListCt") <> "0" Then
        BatchCreateID = BatchSev.GetItem("BLCreateID_1")
        'MsgBox (BatchCreateID)
        BatchResult = BatchSev.Execute("COMMAND(Item,ADV2\Admin," & BatchCreateID & ",START)")
        MsgBox ("START Batch Status: " & BatchResult & " (Create ID: " & BatchCreateID & ")")
    Else
        MsgBox ("Cannot START Batch. No Batches are Created!")
    End If
    
End Sub

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

Private Sub CarmelosTest_Released()

    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim currentBatchNum As String
    
    
    BatchResult = BatchSev.GetItem("TimerSteps")
    MsgBox BatchResult
    Dim StringTemp() As String
    StringTemp = Split(BatchResult, vbTab)
    BatchResult = BatchSev.GetItem(StringTemp(0) & "TimerData")
    MsgBox BatchResult
    Dim TestString4 As String
    TestString4 = BatchSev.GetItem(StringTemp(0) & "TimerStatus")
    MsgBox TestString4
    Dim TestSplit() As String
    TestSplit = Split(TestString4, vbTab)
    If TestString4 Like "*RUNNING*" Then
    Dim StringTemp2() As String
    StringTemp2 = Split(BatchResult, "REMAINING_TIME")
    Dim StringTemp3() As String
    StringTemp3 = Split(StringTemp2(1), vbTab)
    MsgBox TestSplit(7)
    Me.BatchTime.Caption = TestSplit(7)
    Else
    Me.BatchTime.Caption = ""
    End If
    
    

End Sub



Private Sub Display_AnimationStart()

End Sub


Private Sub Group24_Click()

End Sub

Private Sub NumericDisplay6_Change()

    
'''''Numeric display is an onscreen object (maybe hidden) that uses
'''''PC clock seconds to activate this event every second
If NumericDisplay6.Value Mod 5 = 0 Then '''''Every 5 seconds run code below

    ''''' In TOOLS -> References, BatchControl BatchSvr 1.0 type library
    '''''must be checked to create Batch comm object below
    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim currentBatchNum As String
    Dim BatchCreateID As String
    Dim currentBatchState As String
    
    
    BatchResult = BatchSev.GetItem("BatchList")
    '''''Read whats in the batch list
    'MsgBox (BatchResult)
    If BatchResult Like "*FORMULA*" Or BatchResult Like "*CHECK*" Or BatchResult Like "*PROMPT*" Then
    '''''If the batch is named with any of these above words... do code below to change color of nav button
        Me.SideNavButton22.BackStyle = 0
        Me.SideNavButton22.BackColor = 65280
    ElseIf BatchResult = "" Then
    '''''If its empty dont do anything and escape out the end if
    '''''Might have to add code to revert color to grey
    Else
    '''''Otherwise that means a batch is running a process, so lets get the process timer and make the nav button grey
        '''''Revert to grey nav button
        Me.SideNavButton22.BackStyle = 1
        Me.SideNavButton22.BackColor = 15790320
    
        BatchResult = BatchSev.GetItem("TimerSteps")
        'MsgBox BatchResult
        Dim StringTemp() As String
        StringTemp = Split(BatchResult, vbTab)
        Dim TestString4 As String
        TestString4 = BatchSev.GetItem(StringTemp(0) & "TimerStatus")
        'MsgBox TestString4
        If TestString4 Like "*RUNNING*" Then
            BatchResult = BatchSev.GetItem(StringTemp(0) & "TimerData")
            'MsgBox BatchResult
            If BatchResult Like "*REMAINING_TIME*" Then
                Dim TestSplit() As String
                TestSplit = Split(TestString4, vbTab)
                Dim StringTemp2() As String
                StringTemp2 = Split(BatchResult, "REMAINING_TIME")
                Dim StringTemp3() As String
                StringTemp3 = Split(StringTemp2(1), vbTab)
                'MsgBox TestSplit(7)
                Dim TimeRemainNonFormated As String
                Dim TimeRemainFormated As String
                Dim DecimalPos As Integer
                DecimalPos = InStr(1, TimeRemainNonFormated, ".")
                TimeRemainNonFormated = TestSplit(7)
                TimeRemainFormated = Left(TimeRemainNonFormated, (InStr(1, TimeRemainNonFormated, ".")) + 1)
                'If DecimalPos = 0 Then
                '    TimeRemainFormated = TimeRemainFormated & ".0"
                'End If
                               
                
                Me.BatchTime.Caption = TimeRemainFormated
                
            End If
        Else
        Me.BatchTime.Caption = ""
        End If
    End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
   ' If BatchSev.GetItem("BatchListCt") <> "0" Then
   '     BatchCreateID = BatchSev.GetItem("BLCreateID_1")
   '     currentBatchState = BatchSev.GetItem("BLState_1")
   '     'MsgBox (BatchState)
   '     Batch_Status = currentBatchState
   ' End If

    
End If


End Sub
Private Sub Button_HoldBatch_Released()

    Dim BatchSev As New BatchServer
    Dim BatchResult As String
    Dim BatchCreateID As String
    Dim currentBatchNum As String
    
 
    If BatchSev.GetItem("BatchListCt") <> "0" Then
        BatchCreateID = BatchSev.GetItem("BLCreateID_1")
        BatchResult = BatchSev.Execute("COMMAND(Item,ADV2\Admin," & BatchCreateID & ",HOLD)")
        MsgBox ("PAUSED/HELD Batch Status: " & BatchResult & " (Create ID: " & BatchCreateID & ")")
    Else
        MsgBox ("Cannot PAUSE Batch. No Batches are Created OR Running!")
    End If
    
    
    
End Sub

Private Sub StringDisplay_BatchState_Change()

End Sub
