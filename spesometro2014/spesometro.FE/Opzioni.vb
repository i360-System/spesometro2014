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
        If Not IsNothing(My.Settings.txtMod) Then CheckBox1.Checked = My.Settings.txtMod
        If Not IsNothing(My.Settings.conCredenziali) Then CheckBox2.Checked = My.Settings.conCredenziali
        'aggiungere tab 3
        TextBox6.Text = My.Settings.FlussoQuadro1
        TextBox7.Text = My.Settings.FlussoQuadro2
        TextBox8.Text = My.Settings.FlussoQuadro3
        TextBox9.Text = My.Settings.FlussoQuadro4
        TextBox10.Text = My.Settings.FlussoQuadro5
        TextBox11.Text = My.Settings.FlussoQuadro6
        TextBox12.Text = My.Settings.FlussoQuadro7
        TextBox13.Text = My.Settings.FlussoQuadro8
        TextBox14.Text = My.Settings.TipoFornitore
        TextBox15.Text = My.Settings.CodiceFisacaleFornitore
        TextBox16.Text = My.Settings.CodiceFiscaleProduttoreSW
        TextBox17.Text = My.Settings.CodiceCarica
        TextBox18.Text = My.Settings.DataInizioProcedura
        TextBox19.Text = My.Settings.DataFineProcedura
        TextBox20.Text = My.Settings.NumeroCAF
        TextBox21.Text = My.Settings.ImpegnoATrasmettere
        TextBox22.Text = My.Settings.DataImpegno
        If Not IsNothing(My.Settings.MostraExcel) Then CheckBox3.CheckState = My.Settings.MostraExcel
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

    ''' <summary>
    ''' Chiudi
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Dispose()
    End Sub

    ''' <summary>
    ''' Salva
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        With My.Settings
            .PercorsoDB = Trim(TextBox1.Text)
            .OutPutXls = Trim(TextBox2.Text)
            .OutPutTxt = Trim(TextBox3.Text)
            .TipoOleDb = ComboBox1.SelectedIndex
            .conCredenziali = CheckBox2.CheckState
            .txtMod = CheckBox1.Checked
            .MostraExcel = CheckBox3.Checked
            'aggiungere funzionalita di versione differenti
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
            .TipoFornitore = Trim(TextBox14.Text)
            .CodiceFisacaleFornitore = Trim(TextBox15.Text)
            .CodiceFiscaleProduttoreSW = Trim(TextBox16.Text)
            .CodiceCarica = Trim(TextBox17.Text)
            .DataInizioProcedura = Trim(TextBox18.Text)
            .DataFineProcedura = Trim(TextBox19.Text)
            .NumeroCAF = Trim(TextBox20.Text)
            .ImpegnoATrasmettere = Trim(TextBox21.Text)
            .DataImpegno = Trim(TextBox22.Text)
            .Save()
        End With
        MsgBox("Impostazioni salvate con successo.")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        TextBox2.Text = FolderBrowserDialog1.SelectedPath
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog9.FileName = Nothing
        OpenFileDialog9.ShowDialog()
        TextBox1.Text = OpenFileDialog9.FileName
    End Sub

#Region "Selezione dei flussi vuoti excel"
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox6.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox7.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox8.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox9.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox10.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox11.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox12.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.Filter = "File Excel (*.xls)|*.xls|(*.xlms)|*.xlms|(*.xlsx)|*.xlsx|(*.csv)|*.csv"
        OpenFileDialog1.ShowDialog()
        TextBox13.Text = OpenFileDialog1.FileName
    End Sub
#End Region

    ''' <summary>
    ''' Salva e chiudi
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        With My.Settings
            .PercorsoDB = Trim(TextBox1.Text)
            .OutPutXls = Trim(TextBox2.Text)
            .OutPutTxt = Trim(TextBox3.Text)
            .TipoOleDb = ComboBox1.SelectedIndex
            .conCredenziali = CheckBox2.CheckState
            .txtMod = CheckBox1.Checked
            .MostraExcel = CheckBox3.Checked
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
            .TipoFornitore = Trim(TextBox14.Text)
            .CodiceFisacaleFornitore = Trim(TextBox15.Text)
            .CodiceFiscaleProduttoreSW = Trim(TextBox16.Text)
            .CodiceCarica = Trim(TextBox17.Text)
            .DataInizioProcedura = Trim(TextBox18.Text)
            .DataFineProcedura = Trim(TextBox19.Text)
            .NumeroCAF = Trim(TextBox20.Text)
            .ImpegnoATrasmettere = Trim(TextBox21.Text)
            .DataImpegno = Trim(TextBox22.Text)
            .Save()
        End With
        MsgBox("Impostazioni salvate con successo.")
        Me.Dispose()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub TextBox20_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox20.Validating

        Dim stringappoggio = TextBox20.Text

        If stringappoggio.Length > 5 Then
            MsgBox("Il numero del CAF deve essere lungo 5 caratteri")
            TextBox20.Text = String.Empty
        ElseIf stringappoggio.Length < 5 Then
            Dim cinqueZeri = "00000"
            '' str = Left(str, length) equivalent
            'Str = Str.Substring(0, Math.Min(length, Str.Length))
            '' str = Right(str, length) equivalent
            'Str = Str.Substring(Math.Max(Str.Length, length) - length)
            TextBox20.Text = (cinqueZeri & stringappoggio).Substring(Math.Max((cinqueZeri & stringappoggio).Length, 5) - 5)
        End If

    End Sub

    Private Sub TextBox22_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox22.Validating
        Dim stringappoggio As String = TextBox22.Text

        If stringappoggio.Length > 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox22.Text = String.Empty
        ElseIf stringappoggio.Length < 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox22.Text = String.Empty
        ElseIf stringappoggio.Length = 8 Then
            For Each chrt In stringappoggio
                If Not IsNumeric(chrt) Then
                    MsgBox("La data deve essere espressa nella forma GGMMAAAA")
                    TextBox22.Text = String.Empty
                    Exit For
                End If
            Next
        End If


    End Sub

    Private Sub TextBox15_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox15.Validating
        Dim stringappoggio As String = TextBox15.Text

        If stringappoggio.Length > 16 Then
            MsgBox("Il codice fiscale deve essere al massimo di 16 caratteri")
            TextBox15.Text = String.Empty
        End If

    End Sub

    Private Sub TextBox18_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox18.Validating
        Dim stringappoggio As String = TextBox18.Text

        If stringappoggio.Length > 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox18.Text = String.Empty
        ElseIf stringappoggio.Length < 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox18.Text = String.Empty
        ElseIf stringappoggio.Length = 8 Then
            For Each chrt In stringappoggio
                If Not IsNumeric(chrt) Then
                    MsgBox("La data deve essere espressa nella forma GGMMAAAA")
                    TextBox18.Text = String.Empty
                    Exit For
                End If
            Next
        End If

    End Sub

    Private Sub TextBox19_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TextBox19.Validating
        Dim stringappoggio As String = TextBox19.Text

        If stringappoggio.Length > 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox19.Text = String.Empty
        ElseIf stringappoggio.Length < 8 Then
            MsgBox("La data deve essere espressa nella forma GGMMAAAA")
            TextBox19.Text = String.Empty
        ElseIf stringappoggio.Length = 8 Then
            For Each chrt In stringappoggio
                If Not IsNumeric(chrt) Then
                    MsgBox("La data deve essere espressa nella forma GGMMAAAA")
                    TextBox19.Text = String.Empty
                    Exit For
                End If
            Next
        End If

    End Sub

End Class

