Module ObjectTableMovimentiIvaTestata

    Public arr(,) As String
    Dim con As Integer = 0
    Dim tabella As New DataTable

    Public Property tableMovimentiIvaTestata As New List(Of System.Data.DataTable)


    Public Sub CreateMatrice(ByVal numAnag As String, ByVal conti As String)

        ReDim Preserve arr(1, con)
        arr(0, con) = numAnag
        arr(1, con) = conti
        con += 1

    End Sub

    Public Function RitornaMatrice() As String(,)

        Return arr

    End Function

    Public Sub Nullable()

        ReDim arr(1, 0)
        con = 0
        _tableMovimentiIvaTestata = New List(Of System.Data.DataTable)

    End Sub

    Public Property tabellaLista As New List(Of List(Of String))


End Module
