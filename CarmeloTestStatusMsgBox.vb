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
