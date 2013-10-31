Public NotInheritable Class ElaborazioneExcell

    'TODO: questo form può essere facilmente impostato come schermata iniziale per l'applicazione dalla scheda "Applicazione"
    '  di Progettazione progetti (scegliere "Proprietà" dal menu "Progetto").



    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        If InitController.OpzioniGeneraliXls Then
            Me.Labelattendere.Visible = False
            Me.Labelconnessione.Visible = True
            WorkflowBL.mainXls(QuerySelector.selec)
            Me.Labelconnessione.Visible = False
        Else
            MsgBox("Impossibile eseguire l'elaborazione richiesta." & vbCrLf & "Controllare nelle opzioni del software," _
                   & vbCrLf & "che tutti campi e le funzionalità siano valorizzati.")
        End If


    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Enum QuerySelector
        selec = 0
        insert = 1
        update = 2
        delete = 3
    End Enum

End Class
