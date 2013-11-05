Public Class Opzioni

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        MsgBox("Funzione disabilitata")
    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        With ComboBox1
            .Items.Add("Access 97 - 2003")
            .Items.Add("Access 2007 - 2013")
        End With

        If Not IsNothing(My.Settings.TipoOleDb) Then
            ComboBox1.SelectedIndex = My.Settings.TipoOleDb
        End If
        If Not My.Settings.PercorsoDB.ToString = "" Then TextBox1.Text = My.Settings.PercorsoDB
        If Not My.Settings.OutPutXls = "" Then TextBox2.Text = My.Settings.OutPutXls
        If Not My.Settings.OutPutTxt = "" Then TextBox3.Text = My.Settings.OutPutTxt
        If Not IsNothing(My.Settings.txtMod) Then CheckBox1.CheckState = My.Settings.txtMod
        If Not IsNothing(My.Settings.conCredenziali) Then CheckBox2.Checked = My.Settings.conCredenziali
        'aggiungere tab 3
        TextBox6.Text = My.Settings.FlussoQuadro1
        'todo
        If CheckBox2.Checked = True Then
            Label5.Enabled = True
            Label6.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox4.Text = My.Settings.NomeCred
            TextBox5.Text = My.Settings.PassCred
        End If
    End Sub

    Private Sub CheckBox2_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckStateChanged
        If CheckBox2.Checked = False Then
            Label5.Enabled = False
            Label6.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
        Else
            Label5.Enabled = True
            Label6.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox4.Text = My.Settings.NomeCred
            TextBox5.Text = My.Settings.PassCred
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Dispose()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        With My.Settings
            .PercorsoDB = Trim(TextBox1.Text)
            .OutPutXls = Trim(TextBox2.Text)
            .OutPutTxt = Trim(TextBox3.Text)
            .TipoOleDb = ComboBox1.SelectedIndex
            .conCredenziali = CheckBox2.CheckState
            .txtMod = CheckBox1.CheckState
            'aggiungere tab 3
            .FlussoQuadro1 = Trim(TextBox6.Text)
            .FlussoQuadro2 = Trim(TextBox7.Text)
            .FlussoQuadro3 = Trim(TextBox8.Text)
            .FlussoQuadro4 = Trim(TextBox9.Text)
            .FlussoQuadro5 = Trim(TextBox10.Text)
            .FlussoQuadro6 = Trim(TextBox11.Text)
            .FlussoQuadro7 = Trim(TextBox12.Text)
            .FlussoQuadro8 = Trim(TextBox13.Text)
            .NomeCred = TextBox4.Text
            .PassCred = TextBox5.Text
            .Save()
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        TextBox2.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox6.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox7.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox8.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox9.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox10.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox11.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox12.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        TextBox13.Text = OpenFileDialog1.FileName
    End Sub
End Class