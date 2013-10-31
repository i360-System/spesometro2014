Module WorkflowBL

    Dim exc As List(Of Exception)

    ''' <summary>
    ''' MEtodo main di elaborazione del tracciato excell.
    ''' val = 0 --> select
    ''' val = 1 --> insert
    ''' val = 2 --> update
    ''' val = 3 --> delete
    ''' </summary>
    ''' <param name="val"></param>
    ''' <remarks></remarks>
    Public Sub mainXls(ByVal val As String)

        Select Case val.ToString

            Case 0 ' select
                Try
                    If ConnectionState.Open = 0 Then DASL.OleDBcommandConn(QueryBuilder.selec).Open()

                Catch ex As Exception
                    MsgBox(ex.ToString())
                End Try
                
            Case "insert"
            Case "update"
            Case "delete"

        End Select


    End Sub

    Public Sub Err(ByVal ex As Exception)

        exc.Add(ex)

    End Sub

    Private Structure QueryBuilder
        Shared insert = 0
        Shared selec = "Select * from @table where @field1 = @param1"
        Shared update = 0
        Shared delete = 0
        Shared FieldX = "@fieldX"
        Shared ParamX = "@paramX"
    End Structure

End Module
