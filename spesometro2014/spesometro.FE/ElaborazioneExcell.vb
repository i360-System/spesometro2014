Public NotInheritable Class ElaborazioneExcell

    'TODO: questo form può essere facilmente impostato come schermata iniziale per l'applicazione dalla scheda "Applicazione"
    '  di Progettazione progetti (scegliere "Proprietà" dal menu "Progetto").



    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Labelattendere.Visible = False
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
