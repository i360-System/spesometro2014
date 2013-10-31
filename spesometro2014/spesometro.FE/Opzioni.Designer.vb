<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Opzioni
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog2 = New System.Windows.Forms.FolderBrowserDialog()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(486, 211)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.CheckBox1)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.Button3)
        Me.TabPage2.Controls.Add(Me.TextBox3)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.Button2)
        Me.TabPage2.Controls.Add(Me.TextBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(478, 185)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Generale"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Location = New System.Drawing.Point(327, 96)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(141, 17)
        Me.CheckBox1.TabIndex = 16
        Me.CheckBox1.Text = "Attiva Flusso Telematico"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Location = New System.Drawing.Point(6, 55)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Output Tracciato:"
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(440, 50)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(28, 23)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "..."
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(113, 52)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(321, 20)
        Me.TextBox3.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Output XLS:"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(440, 21)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(28, 23)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(113, 23)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(321, 20)
        Me.TextBox2.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CheckBox2)
        Me.TabPage1.Controls.Add(Me.TextBox5)
        Me.TabPage1.Controls.Add(Me.TextBox4)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.ComboBox1)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.TextBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(478, 185)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Database"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(267, 55)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(131, 17)
        Me.CheckBox2.TabIndex = 27
        Me.CheckBox2.Text = "Credenziali di accesso"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Enabled = False
        Me.TextBox5.Location = New System.Drawing.Point(326, 103)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(131, 20)
        Me.TextBox5.TabIndex = 26
        '
        'TextBox4
        '
        Me.TextBox4.Enabled = False
        Me.TextBox4.Location = New System.Drawing.Point(326, 77)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(131, 20)
        Me.TextBox4.TabIndex = 25
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Enabled = False
        Me.Label6.Location = New System.Drawing.Point(264, 106)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Password:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Enabled = False
        Me.Label5.Location = New System.Drawing.Point(247, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Nome Utente:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Tipo Database:"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(113, 52)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Percorso Database:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(440, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(28, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(113, 26)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(321, 20)
        Me.TextBox1.TabIndex = 0
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(339, 229)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 1
        Me.Button5.Text = "&Salva"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(423, 229)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(75, 23)
        Me.Button7.TabIndex = 3
        Me.Button7.Text = "&Esci"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Opzioni
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(510, 257)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Opzioni"
        Me.ShowIcon = False
        Me.Text = "Opzioni"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialog2 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
