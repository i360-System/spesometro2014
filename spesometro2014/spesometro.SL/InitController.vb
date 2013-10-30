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

End Module
