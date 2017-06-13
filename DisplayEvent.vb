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
