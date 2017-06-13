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
