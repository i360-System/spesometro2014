Public Class Form1

    Private Sub OpzioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpzioniToolStripMenuItem.Click
        Opzioni.ShowDialog()
    End Sub

    Private Sub EsciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsciToolStripMenuItem.Click
        Me.Dispose()
    End Sub

    Private Sub GestioneUtenteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GestioneUtenteToolStripMenuItem.Click
        GestioneUtente.ShowDialog()
    End Sub

    Private Sub RiduciInSystemTrayToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RiduciInSystemTrayToolStripMenuItem.Click
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(4)
        Me.Hide()
    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        NotifyIcon1.Visible = False
        Dim mnMenuContestuale = New ContextMenu()
        Dim mnMostra = New MenuItem()
        Dim mnEsci = New MenuItem()
        Dim mnOpzioni = New MenuItem()

        mnMostra.Text = "Mostra spesometro 2014"
        mnEsci.Text = "&Esci"
        mnOpzioni.Text = "&Opzioni..."
        mnMenuContestuale.MenuItems.AddRange(New MenuItem() {mnMostra, mnOpzioni, mnEsci})
        NotifyIcon1.ContextMenu = mnMenuContestuale

    End Sub

    Public Sub StMenuContestuale(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnMostra.Click, mnOpzioni.Click, mnEsci.Click
        Select Case CType(sender, MenuItem).Text
            Case "Mostra spesometro 2014"
                Me.Show()
                'Shell("Notepad.exe", AppWinStyle.NormalFocus)
            Case "&Opzioni..."
                Opzioni.ShowDialog()
                'Shell("Calc.exe", AppWinStyle.NormalFocus)
            Case "&Esci"
                NotifyIcon1.Visible = False
                Application.Exit()
        End Select
    End Sub
End Class
