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
