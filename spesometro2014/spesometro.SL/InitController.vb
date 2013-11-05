


Module InitController

#Region "Controlli sostanziali FE"

    ''' <summary>
    ''' Funzione di SL per il controllo dei campi vuoti del form in questione, viene passata una list of controls
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NotCampiVuotiAll(ByRef obj As Object) As Boolean
        Dim cvbool As Boolean = True
        For Each ctrl As Control In obj
            If TypeOf ctrl Is Global.System.Windows.Forms.TextBox Then
                cvbool = (Not ctrl.Text = "") And cvbool
            End If
        Next
        Return cvbool
    End Function

#End Region

#Region "Controlli di convalida DASL.Database e di coerenza"

    ''' <summary>
    ''' Funzione di SL chiamante la funzione di DASL per la convalida di accesso credenziali
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Credenziali(ByVal obj As List(Of String)) As Boolean
        Return DASL.Credenziali(obj)
    End Function

    ''' <summary>
    ''' Funzione di service layer per il controllo di esistenza dei percorsi di database e di output di xls
    ''' </summary>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Function OpzioniGeneraliXls() As Boolean

        Return (Not My.Settings.PercorsoDB.ToString = "") And ((Not My.Settings.FlussoQuadro1 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro2 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro3 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro4 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro5 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro6 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro7 = "") OrElse _
                                                               (Not My.Settings.FlussoQuadro8 = ""))

    End Function

#End Region

End Module
