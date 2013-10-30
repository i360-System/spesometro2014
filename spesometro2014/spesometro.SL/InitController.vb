Module InitController

    ''' <summary>
    ''' Funzione di service layer per il controllo di esistenza dei percorsi di database e di output di xls
    ''' </summary>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Function OpzioniGeneraliXls() As Boolean

        Return (Not My.Settings.PercorsoDB.ToString = "") And (Not My.Settings.OutPutXls = "") _
            And (Not My.Settings.tipodb = "")

    End Function

    Public Function NotCampiVuotiAll(ByRef obj As Object) As Boolean
        Dim cvbool As Boolean = True
        For Each controls In obj
            If TypeOf obj Is TextBox Then
                cvbool = (Not obj.text = "") And cvbool
            End If
        Next
        Return cvbool
    End Function

End Module
