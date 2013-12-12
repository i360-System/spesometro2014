<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ElaborazioneExcell
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
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Labelcompletato = New System.Windows.Forms.Label()
        Me.Labelcontrollo = New System.Windows.Forms.Label()
        Me.Labelxls = New System.Windows.Forms.Label()
        Me.Labelattendere = New System.Windows.Forms.Label()
        Me.Labelelaborazione = New System.Windows.Forms.Label()
        Me.Labelraccoltainfo = New System.Windows.Forms.Label()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.UserControlMenuXLS1 = New spesometro2014.UserControlMenuXLS()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 164)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(989, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.UserControlMenuXLS1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(989, 66)
        Me.Panel1.TabIndex = 5
        '
        'Labelcompletato
        '
        Me.Labelcompletato.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelcompletato.AutoSize = True
        Me.Labelcompletato.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelcompletato.Location = New System.Drawing.Point(309, 81)
        Me.Labelcompletato.Name = "Labelcompletato"
        Me.Labelcompletato.Size = New System.Drawing.Size(125, 23)
        Me.Labelcompletato.TabIndex = 9
        Me.Labelcompletato.Text = "Completato..."
        Me.Labelcompletato.Visible = False
        '
        'Labelcontrollo
        '
        Me.Labelcontrollo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelcontrollo.AutoSize = True
        Me.Labelcontrollo.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelcontrollo.Location = New System.Drawing.Point(309, 81)
        Me.Labelcontrollo.Name = "Labelcontrollo"
        Me.Labelcontrollo.Size = New System.Drawing.Size(227, 23)
        Me.Labelcontrollo.TabIndex = 8
        Me.Labelcontrollo.Text = "Controllo coerenza dati..."
        Me.Labelcontrollo.Visible = False
        '
        'Labelxls
        '
        Me.Labelxls.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelxls.AutoSize = True
        Me.Labelxls.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelxls.Location = New System.Drawing.Point(309, 81)
        Me.Labelxls.Name = "Labelxls"
        Me.Labelxls.Size = New System.Drawing.Size(287, 23)
        Me.Labelxls.TabIndex = 7
        Me.Labelxls.Text = "Elaborazione Excell, attendere..."
        Me.Labelxls.Visible = False
        '
        'Labelattendere
        '
        Me.Labelattendere.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelattendere.AutoSize = True
        Me.Labelattendere.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelattendere.Location = New System.Drawing.Point(309, 81)
        Me.Labelattendere.Name = "Labelattendere"
        Me.Labelattendere.Size = New System.Drawing.Size(301, 23)
        Me.Labelattendere.TabIndex = 6
        Me.Labelattendere.Text = "Attendere, elaborazione in corso..."
        '
        'Labelelaborazione
        '
        Me.Labelelaborazione.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelelaborazione.AutoSize = True
        Me.Labelelaborazione.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelelaborazione.Location = New System.Drawing.Point(309, 81)
        Me.Labelelaborazione.Name = "Labelelaborazione"
        Me.Labelelaborazione.Size = New System.Drawing.Size(215, 23)
        Me.Labelelaborazione.TabIndex = 10
        Me.Labelelaborazione.Text = "Elaborazione in corso..."
        Me.Labelelaborazione.Visible = False
        '
        'Labelraccoltainfo
        '
        Me.Labelraccoltainfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Labelraccoltainfo.AutoSize = True
        Me.Labelraccoltainfo.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelraccoltainfo.Location = New System.Drawing.Point(309, 106)
        Me.Labelraccoltainfo.Name = "Labelraccoltainfo"
        Me.Labelraccoltainfo.Size = New System.Drawing.Size(216, 23)
        Me.Labelraccoltainfo.TabIndex = 11
        Me.Labelraccoltainfo.Text = "Raccolta informazioni..."
        Me.Labelraccoltainfo.Visible = False
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(306, 135)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(344, 23)
        Me.ProgressBar2.TabIndex = 12
        Me.ProgressBar2.Visible = False
        '
        'UserControlMenuXLS1
        '
        Me.UserControlMenuXLS1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UserControlMenuXLS1.Location = New System.Drawing.Point(0, 0)
        Me.UserControlMenuXLS1.Name = "UserControlMenuXLS1"
        Me.UserControlMenuXLS1.Size = New System.Drawing.Size(989, 66)
        Me.UserControlMenuXLS1.TabIndex = 0
        '
        'ElaborazioneExcell
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 187)
        Me.ControlBox = False
        Me.Controls.Add(Me.ProgressBar2)
        Me.Controls.Add(Me.Labelraccoltainfo)
        Me.Controls.Add(Me.Labelelaborazione)
        Me.Controls.Add(Me.Labelcompletato)
        Me.Controls.Add(Me.Labelcontrollo)
        Me.Controls.Add(Me.Labelxls)
        Me.Controls.Add(Me.Labelattendere)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ElaborazioneExcell"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents UserControlMenuXLS1 As spesometro2014.UserControlMenuXLS
    Friend WithEvents Labelcompletato As System.Windows.Forms.Label
    Friend WithEvents Labelcontrollo As System.Windows.Forms.Label
    Friend WithEvents Labelxls As System.Windows.Forms.Label
    Friend WithEvents Labelattendere As System.Windows.Forms.Label
    Friend WithEvents Labelelaborazione As System.Windows.Forms.Label
    Friend WithEvents Labelraccoltainfo As System.Windows.Forms.Label
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar

End Class
