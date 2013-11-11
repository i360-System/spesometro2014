Public Class LoginForm1

    ' TODO: inserire il codice per l'esecuzione dell'autenticazione personalizzata tramite il nome utente e la password forniti 
    ' (Vedere http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' L'entità personalizzata può essere quindi collegata all'entità del thread corrente nel modo seguente: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' dove CustomPrincipal è l'implementazione di IPrincipal utilizzata per eseguire l'autenticazione. 
    ' My.User restituirà quindi informazioni sull'identità incapsulate nell'oggetto CustomPrincipal,
    ' quali il nome utente, il nome visualizzato e così via.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If InitController.NotCampiVuotiAll(Me.Controls) Then
            Dim p As New List(Of String)
            p.AddRange({Trim(Me.UsernameTextBox.Text), Trim(Me.PasswordTextBox.Text)})
            If InitController.Credenziali(p) Then

                Me.Hide()
                MainForm.ShowDialog()

            Else

                MsgBox("Credenziali inserite non corrette.")

            End If

        Else

            MsgBox("Compilare correttamente i campi.")

        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Dispose()
    End Sub

End Class
