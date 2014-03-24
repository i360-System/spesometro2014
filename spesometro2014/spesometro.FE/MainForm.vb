Public Class MainForm
    Private flgCloseAllowed As Boolean = False
    Public Telematico As Boolean = False
    Private Sub OpzioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpzioniToolStripMenuItem.Click
        Opzioni.ShowDialog()
    End Sub

    Private Sub EsciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsciToolStripMenuItem.Click
        Me.Dispose()
        LoginForm1.Dispose()
    End Sub

    Private Sub GestioneUtenteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GestioneUtenteToolStripMenuItem.Click
        'GestioneUtente.ShowDialog()
        MsgBox("Funzione disabilitata in questa versione.")
    End Sub

    Private Sub RiduciInSystemTrayToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RiduciInSystemTrayToolStripMenuItem.Click
        flgCloseAllowed = True
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(4)
        Me.Hide()
    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        NotifyIcon1.Visible = False
        mnMenuContestuale = New ContextMenu()
        mnMostra = New System.Windows.Forms.MenuItem()
        mnEsci = New System.Windows.Forms.MenuItem()
        mnOpzioni = New System.Windows.Forms.MenuItem()

        mnMostra.Text = "Mostra spesometro 2014"
        mnEsci.Text = "&Esci"
        mnOpzioni.Text = "&Opzioni..."
        mnMenuContestuale.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnMostra, mnOpzioni, mnEsci})
        NotifyIcon1.ContextMenu = mnMenuContestuale

    End Sub

    Public Sub mnMenuContestuale_click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnMostra.Click, mnOpzioni.Click, mnEsci.Click
        Select Case CType(sender, System.Windows.Forms.MenuItem).Text
            Case "Mostra spesometro 2014"
                NotifyIcon1.Visible = False
                Me.Show()
                'Shell("Notepad.exe", AppWinStyle.NormalFocus)
            Case "&Opzioni..."
                Opzioni.ShowDialog()
                'Shell("Calc.exe", AppWinStyle.NormalFocus)
            Case "&Esci"
                NotifyIcon1.Visible = False
                LoginForm1.Dispose()
        End Select
    End Sub


    Private Sub InformazioniSulSoftwareToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InformazioniSulSoftwareToolStripMenuItem.Click
        Info.ShowDialog()
    End Sub

   

    Private Sub TracciatoTelematicoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TracciatoTelematicoToolStripMenuItem.Click
        If Not My.Settings.txtMod Then
            MsgBox("La funzione di elaborazione tracciato txt telematico," & vbCrLf & "non è stata attivata dal pannello Opzioni.")
            Telematico = False
        Else
            Telematico = True
            ElaborazioneExcell.ShowDialog()
            Telematico = False
        End If
    End Sub

    Private Sub ExcellToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcellToolStripMenuItem.Click

        ElaborazioneExcell.ShowDialog()
    End Sub

   
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Process.Start(Application.StartupPath & "\Manual\Manuale di utilizzo software Spesometro 2013.pdf")
    End Sub
End Class