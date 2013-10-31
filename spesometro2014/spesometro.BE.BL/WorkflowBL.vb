Module WorkflowBL
    Dim exc As List(Of Exception)
    Public Sub mainXls()
        'connette al db
    End Sub
    Public Sub Err(ByVal ex As Exception)
        exc.Add(ex)
    End Sub
End Module
