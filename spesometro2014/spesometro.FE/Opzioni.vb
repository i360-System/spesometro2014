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
        If Not IsNothing(My.Settings.conCredenziali) Then CheckBox2.CheckState = My.Settings.conCredenziali
        If CheckBox2.CheckState = True Then
            Label5.Enabled = True
            Label6.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox4.Text = My.Settings.NomeCred
            TextBox5.Text = My.Settings.PassCred
        End If
    End Sub

    Private Sub CheckBox2_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckStateChanged
        If CheckBox2.CheckState = False Then
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
            .PercorsoDB = TextBox1.Text
            .OutPutXls = TextBox2.Text
            .OutPutTxt = TextBox3.Text
            .TipoOleDb = ComboBox1.SelectedIndex
            .conCredenziali = CheckBox2.CheckState
            .txtMod = CheckBox1.CheckState
            .NomeCred = TextBox4.Text
            .PassCred = TextBox5.Text
            .Save()
        End With
    End Sub
End Class