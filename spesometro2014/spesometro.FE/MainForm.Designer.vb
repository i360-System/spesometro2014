<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpzioniToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RiduciInSystemTrayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EsciToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StrumentiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CreazioneFlussoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcellToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TracciatoTelematicoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GestioneUtenteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.InformazioniSulSoftwareToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.StrumentiToolStripMenuItem, Me.ToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(967, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpzioniToolStripMenuItem, Me.RiduciInSystemTrayToolStripMenuItem, Me.EsciToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpzioniToolStripMenuItem
        '
        Me.OpzioniToolStripMenuItem.Name = "OpzioniToolStripMenuItem"
        Me.OpzioniToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.OpzioniToolStripMenuItem.Text = "Opzioni..."
        '
        'RiduciInSystemTrayToolStripMenuItem
        '
        Me.RiduciInSystemTrayToolStripMenuItem.Name = "RiduciInSystemTrayToolStripMenuItem"
        Me.RiduciInSystemTrayToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.RiduciInSystemTrayToolStripMenuItem.Text = "Riduci in System Tray"
        '
        'EsciToolStripMenuItem
        '
        Me.EsciToolStripMenuItem.Name = "EsciToolStripMenuItem"
        Me.EsciToolStripMenuItem.Size = New System.Drawing.Size(187, 22)
        Me.EsciToolStripMenuItem.Text = "Esci"
        '
        'StrumentiToolStripMenuItem
        '
        Me.StrumentiToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreazioneFlussoToolStripMenuItem, Me.GestioneUtenteToolStripMenuItem})
        Me.StrumentiToolStripMenuItem.Name = "StrumentiToolStripMenuItem"
        Me.StrumentiToolStripMenuItem.Size = New System.Drawing.Size(71, 20)
        Me.StrumentiToolStripMenuItem.Text = "Strumenti"
        '
        'CreazioneFlussoToolStripMenuItem
        '
        Me.CreazioneFlussoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcellToolStripMenuItem, Me.TracciatoTelematicoToolStripMenuItem})
        Me.CreazioneFlussoToolStripMenuItem.Name = "CreazioneFlussoToolStripMenuItem"
        Me.CreazioneFlussoToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.CreazioneFlussoToolStripMenuItem.Text = "Creazione flusso"
        '
        'ExcellToolStripMenuItem
        '
        Me.ExcellToolStripMenuItem.Name = "ExcellToolStripMenuItem"
        Me.ExcellToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.ExcellToolStripMenuItem.Text = "Excell"
        '
        'TracciatoTelematicoToolStripMenuItem
        '
        Me.TracciatoTelematicoToolStripMenuItem.Name = "TracciatoTelematicoToolStripMenuItem"
        Me.TracciatoTelematicoToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.TracciatoTelematicoToolStripMenuItem.Text = "Tracciato Telematico"
        '
        'GestioneUtenteToolStripMenuItem
        '
        Me.GestioneUtenteToolStripMenuItem.Name = "GestioneUtenteToolStripMenuItem"
        Me.GestioneUtenteToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.GestioneUtenteToolStripMenuItem.Text = "Gestione Utente"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InformazioniSulSoftwareToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(24, 20)
        Me.ToolStripMenuItem1.Text = "?"
        '
        'InformazioniSulSoftwareToolStripMenuItem
        '
        Me.InformazioniSulSoftwareToolStripMenuItem.Name = "InformazioniSulSoftwareToolStripMenuItem"
        Me.InformazioniSulSoftwareToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.InformazioniSulSoftwareToolStripMenuItem.Text = "Informazioni sul software"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipText = "Spesometro 2014 verrà ridotto a ciona sulla barra di system tray"
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Spesometro 2014"
        Me.NotifyIcon1.Visible = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(967, 687)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.Text = "Spesometro 2014"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpzioniToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EsciToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StrumentiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CreazioneFlussoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExcellToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TracciatoTelematicoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InformazioniSulSoftwareToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GestioneUtenteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RiduciInSystemTrayToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon

    'Friend WithEvents nfiIcona As NotifyIcon
    Friend WithEvents mnMenuContestuale As System.Windows.Forms.ContextMenu
    Friend WithEvents mnEsci As System.Windows.Forms.MenuItem
    Friend WithEvents mnOpzioni As System.Windows.Forms.MenuItem
    Friend WithEvents mnMostra As System.Windows.Forms.MenuItem

End Class
