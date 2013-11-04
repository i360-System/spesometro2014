Public Class UserControlMenuXLS

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        WorkflowBL.mainXls(0)
    End Sub

    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        With ComboBox1
            .Items.Add("Analitica")
            .Items.Add("Aggregazione di dati")
            .SelectedIndex = 0
        End With
        With ComboBox2
            .Items.Add("Fattura emesse - Quadro FE")
            .Items.Add("Note credito emesse - Quadro NE")
            .Items.Add("Fatture ricevute - Quadro FR")
            .Items.Add("Note credito ricevute - Quadro NR")
            '.items.add("All")
            .SelectedIndex = 0
        End With

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub
End Class
