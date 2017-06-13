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
