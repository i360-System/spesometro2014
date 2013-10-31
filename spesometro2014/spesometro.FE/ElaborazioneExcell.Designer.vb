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
        Me.Labelattendere = New System.Windows.Forms.Label()
        Me.Labelconnessione = New System.Windows.Forms.Label()
        Me.Labelcontrollo = New System.Windows.Forms.Label()
        Me.Labelcompletato = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 164)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(581, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'Labelattendere
        '
        Me.Labelattendere.AutoSize = True
        Me.Labelattendere.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelattendere.Location = New System.Drawing.Point(57, 69)
        Me.Labelattendere.Name = "Labelattendere"
        Me.Labelattendere.Size = New System.Drawing.Size(341, 25)
        Me.Labelattendere.TabIndex = 1
        Me.Labelattendere.Text = "Attendere, elaborazione in corso..."
        '
        'Labelconnessione
        '
        Me.Labelconnessione.AutoSize = True
        Me.Labelconnessione.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelconnessione.Location = New System.Drawing.Point(57, 69)
        Me.Labelconnessione.Name = "Labelconnessione"
        Me.Labelconnessione.Size = New System.Drawing.Size(459, 25)
        Me.Labelconnessione.TabIndex = 2
        Me.Labelconnessione.Text = "Connessione al Database e controllo dei dati..."
        Me.Labelconnessione.Visible = False
        '
        'Labelcontrollo
        '
        Me.Labelcontrollo.AutoSize = True
        Me.Labelcontrollo.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelcontrollo.Location = New System.Drawing.Point(57, 69)
        Me.Labelcontrollo.Name = "Labelcontrollo"
        Me.Labelcontrollo.Size = New System.Drawing.Size(252, 25)
        Me.Labelcontrollo.TabIndex = 3
        Me.Labelcontrollo.Text = "Controllo coerenza dati..."
        Me.Labelcontrollo.Visible = False
        '
        'Labelcompletato
        '
        Me.Labelcompletato.AutoSize = True
        Me.Labelcompletato.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelcompletato.Location = New System.Drawing.Point(57, 69)
        Me.Labelcompletato.Name = "Labelcompletato"
        Me.Labelcompletato.Size = New System.Drawing.Size(139, 25)
        Me.Labelcompletato.TabIndex = 4
        Me.Labelcompletato.Text = "Completato..."
        Me.Labelcompletato.Visible = False
        '
        'ElaborazioneExcell
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(581, 187)
        Me.ControlBox = False
        Me.Controls.Add(Me.Labelcompletato)
        Me.Controls.Add(Me.Labelcontrollo)
        Me.Controls.Add(Me.Labelconnessione)
        Me.Controls.Add(Me.Labelattendere)
        Me.Controls.Add(Me.ProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ElaborazioneExcell"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Labelattendere As System.Windows.Forms.Label
    Friend WithEvents Labelconnessione As System.Windows.Forms.Label
    Friend WithEvents Labelcontrollo As System.Windows.Forms.Label
    Friend WithEvents Labelcompletato As System.Windows.Forms.Label

End Class
