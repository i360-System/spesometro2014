Public NotInheritable Class ElaborazioneExcell
    Public telematico As Boolean = False
    'TODO: questo form può essere facilmente impostato come schermata iniziale per l'applicazione dalla scheda "Applicazione"
    '  di Progettazione progetti (scegliere "Proprietà" dal menu "Progetto").

    Public Tipocomunicazione As Byte

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Labelattendere.Visible = False
        'If MainForm.Telematico = True Then
        '    telematico = True
        'Else
        '    telematico = False
        'End If

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
